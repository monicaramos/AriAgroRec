Attribute VB_Name = "ModContabilizar"
Option Explicit


'===================================================================================
'CONTABILIZAR FACTURAS:
'Modulo para el traspaso de registros de cabecera y lineas de tablas de FACTURACION
'A las tablas de FACTURACION de Contabilidad
'====================================================================================

Private DtoGnral As Currency
Private DtoPPago As Currency
Private BaseImp As Currency
Private TotalFac As Currency
Private CCoste As String

Private conCtaAlt As Boolean 'el cliente utiliza cuentas alternativas

'Para pasar a contabilidad facturas de socio
Private AnyoFacPr As Integer 'año factura socio, es el ano de fecha_recepcion

Private IbanSoc As String
Private BancoSoc As Integer
Private SucurSoc As Integer
Private DigcoSoc As String
Private CtaBaSoc As String

Private Socio As String
Private CtaSocio As String ' cuenta contable del socio para la seccion que hemos introducido
Private Seccion As Integer
Private TipoFact As Integer
Private FecRecep As Date

Private FecVenci As Date
Private ForpaPosi As Integer
Private ForpaNega As Integer
Private CtaBanco As String
Private CtaReten As String
Private CtaAport As String
Private tipoMov As String
Private FacturaSoc As String
Private FecFactuSoc As Date
Private ImpReten As Currency
Private ImpAport As Currency
Private CodiIVA As String


Private CodTipomRECT As String
Private NumfactuRECT As String
Private FecfactuRECT As String


Private Variedades As String
Private TotalTesor As Currency


Private FacturaTRA As String
Private FecFactuTRA As Date
Private CtaTransporte As String ' cuenta contable del transportista

Private IbanTRA As String
Private BancoTRA As Integer
Private SucurTRA As Integer
Private DigcoTRA As String
Private CtaBaTRA As String

Dim vvIban As String
Dim vSoc As cSocio
Dim vTra As CTransportista

Private vTipoIva(2) As Currency
Private vPorcIva(2) As Currency
Private vPorcRec(2) As Currency
Private vBaseIva(2) As Currency
Private vImpIva(2) As Currency
Private vImpRec(2) As Currency

' para terceros
Private ForPago As String
Private mCodmacta As String
Private numfactu As String
Private fecfactu As String
Private CuentaPrev As String


Public Function CrearTMPFacturas(cadTabla As String, cadWHERE As String) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPFacturas = False
    
    Sql = "CREATE TEMPORARY TABLE tmpFactu ( "
    If cadTabla = "facturas" Then
        Sql = Sql & "codtipom char(3) NOT NULL default '',"
        Sql = Sql & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Else
        If cadTabla = "rfactsoc" Or cadTabla = "advfacturas" Or cadTabla = "rbodfacturas" Or cadTabla = "fvarcabfact" Or cadTabla = "fvarcabfactpro" Then
            Sql = Sql & "codtipom char(3) NOT NULL default '',"
            Sql = Sql & "numfactu int(7)  NOT NULL ,"
        Else
            If cadTabla = "rcabfactalmz" Then
                Sql = Sql & "tipofichero smallint(1) unsigned NOT NULL,"
                Sql = Sql & "codsocio smallint(3) unsigned NOT NULL default '0',"
                Sql = Sql & "numfactu int(7)  NOT NULL ,"
            Else
                If cadTabla = "rtelmovil" Then
                    Sql = Sql & "numserie varchar(1) NOT NULL,"
                    Sql = Sql & "numfactu int(7)     NOT NULL,"
                Else
                    If cadTabla = "rrecibpozos" Then
                        Sql = Sql & "codtipom char(3) NOT NULL,"
                        Sql = Sql & "numfactu int(7) unsigned NOT NULL,"
                    Else
                        If cadTabla = "rfacttra" Then
                            Sql = Sql & "codtipom char(3) NOT NULL default '',"
                            Sql = Sql & "numfactu int(7)  NOT NULL ,"
                        Else
                            Sql = Sql & "codsocio int(7) unsigned NOT NULL default '0',"
                            Sql = Sql & "numfactu varchar(10)  NOT NULL ,"
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    Sql = Sql & "fecfactu date NOT NULL default '0000-00-00' "
    
    If cadTabla = "rfacttra" Then
        Sql = Sql & ",codtrans varchar(10))"
    Else
        Sql = Sql & ")"
    End If
    
    conn.Execute Sql
     
     
    If cadTabla = "facturas" Or cadTabla = "advfacturas" Or cadTabla = "rbodfacturas" Or cadTabla = "fvarcabfact" Or cadTabla = "fvarcabfactpro" Then
        Sql = "SELECT codtipom, numfactu, fecfactu"
    Else
        If cadTabla = "rfactsoc" Then
            Sql = "SELECT codtipom, numfactu, fecfactu"
        Else
            If cadTabla = "rcabfactalmz" Then
                Sql = "SELECT tipofichero, codsocio, numfactu, fecfactu "
            Else
                If cadTabla = "rtelmovil" Then
                    Sql = "SELECT numserie, numfactu, fecfactu "
                Else
                    If cadTabla = "rfacttra" Then
                        Sql = "SELECT codtipom, numfactu, fecfactu, codtrans"
                    Else
                        If cadTabla = "rrecibpozos" Then
                            Sql = "SELECT DISTINCT codtipom, numfactu, fecfactu "
                        Else
                            Sql = "SELECT codsocio, numfactu, fecfactu"
                        End If
                    End If
                End If
            End If
        End If
    End If
    Sql = Sql & " FROM " & cadTabla
    Sql = Sql & " WHERE " & cadWHERE
    Sql = " INSERT INTO tmpFactu " & Sql
    conn.Execute Sql

    CrearTMPFacturas = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPFacturas = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpFactu;"
        conn.Execute Sql
    End If
End Function


Public Sub BorrarTMPFacturas()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmpFactu;"
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub InsertarTMPErrFac(MenError As String, cadWHERE As String)
Dim Sql As String

    On Error Resume Next
    Sql = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
    Sql = Sql & " Select *," & DBSet(Mid(MenError, 1, 200), "T") & " as error From tmpFactu "
    Sql = Sql & " WHERE " & Replace(cadWHERE, "rfactsoc", "tmpFactu")
    conn.Execute Sql
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub InsertarTMPErrFacSoc(MenError As String, cadWHERE As String)
Dim Sql As String

    On Error Resume Next
    Sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
    Sql = Sql & " Select *," & DBSet(Mid(MenError, 1, 200), "T") & " as error From tmpFactu "
    Sql = Sql & " WHERE " & Replace(cadWHERE, "rfactsoc", "tmpFactu")
    conn.Execute Sql
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub InsertarTMPErrFacFVAR(MenError As String, cadWHERE As String)
Dim Sql As String

    On Error Resume Next
    Sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
    Sql = Sql & " Select *," & DBSet(Mid(MenError, 1, 200), "T") & " as error From tmpFactu "
    Sql = Sql & " WHERE " & Replace(cadWHERE, "fvarcabfact", "tmpFactu")
    conn.Execute Sql
    
    If Err.Number <> 0 Then Err.Clear
End Sub





Public Function CrearTMPErrFact(cadTabla As String) As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPErrFact = False
    
    Sql = "CREATE TEMPORARY TABLE tmpErrFac ( "
    If cadTabla = "facturas" Or cadTabla = "rfactsoc" Or cadTabla = "rbodfacturas" Or cadTabla = "rrecibpozos" Or cadTabla = "fvarcabfact" Or cadTabla = "fvarcabfactpro" Then
        Sql = Sql & "codtipom char(3) NOT NULL default '',"
        Sql = Sql & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Else
        If cadTabla = "rcabfactalmz" Then
            Sql = Sql & "tipofichero smallint unsigned NOT NULL, "
            Sql = Sql & "numfactu int(7) NOT NULL ,"
        Else
            Sql = Sql & "codprove int(6) unsigned NOT NULL default '0',"
            Sql = Sql & "numfactu varchar(10) NOT NULL ,"
        End If
    End If
    Sql = Sql & "fecfactu date NOT NULL default '0000-00-00', "
    
    If cadTabla = "rcabfactalmz" Then Sql = Sql & "codsocio int(7) ,"
    
    Sql = Sql & "error varchar(200) NULL )"
    
    conn.Execute Sql
     
    CrearTMPErrFact = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrFact = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpErrFac;"
        conn.Execute Sql
    End If
End Function

Public Function CrearTMPErrComprob() As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPErrComprob = False
    
    Sql = "CREATE TEMPORARY TABLE tmperrcomprob ( "
    Sql = Sql & "error varchar(100) NULL )"
    conn.Execute Sql
     
    CrearTMPErrComprob = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrComprob = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmperrcomprob;"
        conn.Execute Sql
    End If
End Function


Public Sub BorrarTMPErrFact()
On Error Resume Next
    conn.Execute " DROP TABLE IF EXISTS tmpErrFac;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub BorrarTMPErrComprob()
On Error Resume Next
    conn.Execute " DROP TABLE IF EXISTS tmperrcomprob;"
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function ComprobarLetraSerie(cadTabla As String) As Boolean
'Para Facturas VENTA a clientes
'Comprueba que la letra del serie del tipo de movimiento es  correcta
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim B As Boolean
Dim cad As String, devuelve As String
Dim Sql2 As String
Dim Total As Long

On Error GoTo EComprobarLetra

    ComprobarLetraSerie = False
    
    'Comprobar que existe la letra de serie en contabilidad
    Select Case cadTabla
        Case "rfactsoc"
            'cargamos el RSConta con la tabla contadores de BD: Contabilidad
            'donde estan todas las letra de serie que existen en la contabilidad
            Sql = "Select distinct tiporegi from contadores"
            Set RSconta = New ADODB.Recordset
            RSconta.Open Sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
            If RSconta.EOF Then
                RSconta.Close
                Set RSconta = Nothing
                Exit Function
            End If
                
        
            'obtenemos los distintos tipos de movimiento que vamos a contabilizar
            'de las facturas seleccionadas
            Sql = "select distinct rfactsoc.codtipom from " & cadTabla
            Sql = Sql & " INNER JOIN tmpFactu ON rfactsoc.codtipom=tmpFactu.codtipom AND rfactsoc.numfactu=tmpFactu.numfactu AND rfactsoc.fecfactu=tmpFactu.fecfactu "
    '        SQL = SQL & cadWHERE
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            cad = ""
            B = True
            While Not Rs.EOF And B
                'comprobar que todas las letras serie existen en usuarios
                Sql = "letraser"
                Sql2 = "select letraser from usuarios.stipom where codtipom = " & DBSet(Rs!CodTipom, "T")
                'devuelve = DevuelveDesdeBDNew(cAgro, "stipom", "codtipom", "codtipom", RS!codtipom, "T", SQL)
                Total = TotalRegistrosConsulta(Sql2)
                If Total = 0 Then 'devuelve = "" Then
                    B = False
                    cad = Rs!CodTipom & " en BD de Ariagrorec."
                ElseIf DevuelveValor(Sql2) <> "" Then 'SQL <> "" Then
                    'comprobar que todas las letras serie existen en la contabilidad
                    devuelve = "tiporegi= " & DBSet(DevuelveValor(Sql2), "T") 'SQL, "T")
                    RSconta.MoveFirst
                    RSconta.Find (devuelve), , adSearchForward
                    If RSconta.EOF Then
                        'no encontrado
                        B = False
                        'Cad = SQL & " en BD de Contabilidad."
                        cad = DevuelveValor(Sql2) & " en BD de Contabilidad."
                    End If
                End If
                If B Then cad = cad & DBSet(Rs!CodTipom, "T") & ","
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
            RSconta.Close
            Set RSconta = Nothing
            
            If Not B Then 'Hay algun movimiento que no existe
                devuelve = "No existe el tipo de movimiento: " & cad & vbCrLf
                devuelve = devuelve & "Consulte con el administrador."
                MsgBox devuelve, vbExclamation
                Exit Function
            End If
            
            'Todos los Tipo de movimiento existen
            If cad <> "" Then
                cad = Mid(cad, 1, Len(cad) - 1) 'quitamos ult. coma
            
                'miramos si hay algun movimiento de factura que la letra serie sea nulo
                Sql = "select count(*) from usuarios.stipom "
                Sql = Sql & "where codtipom IN (" & cad & ") and (isnull(letraser) or letraser='')"
                If RegistrosAListar(Sql) > 0 Then
                    Sql = "Hay algun tipo de movimiento de Facturación que no tiene letra serie." & vbCrLf
                    Sql = Sql & "Comprobar en la tabla de tipos de movimiento: " & cad
                    MsgBox Sql, vbExclamation
                    Exit Function
                End If
            End If
            ComprobarLetraSerie = True
        Case "advfacturas"
            'cargamos el RSConta con la tabla contadores de BD: Contabilidad
            'donde estan todas las letra de serie que existen en la contabilidad
            Sql = "Select distinct tiporegi from contadores"
            Set RSconta = New ADODB.Recordset
            RSconta.Open Sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
            If RSconta.EOF Then
                RSconta.Close
                Set RSconta = Nothing
                Exit Function
            End If
                
        
            'obtenemos los distintos tipos de movimiento que vamos a contabilizar
            'de las facturas seleccionadas
            Sql = "select distinct advfacturas.codtipom from " & cadTabla
            Sql = Sql & " INNER JOIN tmpFactu ON advfacturas.codtipom=tmpFactu.codtipom AND advfacturas.numfactu=tmpFactu.numfactu AND advfacturas.fecfactu=tmpFactu.fecfactu "
    '        SQL = SQL & cadWHERE
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            cad = ""
            B = True
            While Not Rs.EOF And B
                'comprobar que todas las letras serie existen en usuarios
                Sql = "letraser"
                Sql2 = "select letraser from usuarios.stipom where codtipom = " & DBSet(Rs!CodTipom, "T")
                'devuelve = DevuelveDesdeBDNew(cAgro, "stipom", "codtipom", "codtipom", RS!codtipom, "T", SQL)
                Total = TotalRegistrosConsulta(Sql2)
                If Total = 0 Then 'devuelve = "" Then
                    B = False
                    cad = Rs!CodTipom & " en BD de Ariagrorec."
                ElseIf DevuelveValor(Sql2) <> "" Then 'SQL <> "" Then
                    'comprobar que todas las letras serie existen en la contabilidad
                    devuelve = "tiporegi= " & DBSet(DevuelveValor(Sql2), "T") 'SQL, "T")
                    RSconta.MoveFirst
                    RSconta.Find (devuelve), , adSearchForward
                    If RSconta.EOF Then
                        'no encontrado
                        B = False
                        'Cad = SQL & " en BD de Contabilidad."
                        cad = DevuelveValor(Sql2) & " en BD de Contabilidad."
                    End If
                End If
                If B Then cad = cad & DBSet(Rs!CodTipom, "T") & ","
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
            RSconta.Close
            Set RSconta = Nothing
            
            If Not B Then 'Hay algun movimiento que no existe
                devuelve = "No existe el tipo de movimiento: " & cad & vbCrLf
                devuelve = devuelve & "Consulte con el administrador."
                MsgBox devuelve, vbExclamation
                Exit Function
            End If
            
            'Todos los Tipo de movimiento existen
            If cad <> "" Then
                cad = Mid(cad, 1, Len(cad) - 1) 'quitamos ult. coma
            
                'miramos si hay algun movimiento de factura que la letra serie sea nulo
                Sql = "select count(*) from usuarios.stipom "
                Sql = Sql & "where codtipom IN (" & cad & ") and (isnull(letraser) or letraser='')"
                If RegistrosAListar(Sql) > 0 Then
                    Sql = "Hay algun tipo de movimiento de Facturación que no tiene letra serie." & vbCrLf
                    Sql = Sql & "Comprobar en la tabla de tipos de movimiento: " & cad
                    MsgBox Sql, vbExclamation
                    Exit Function
                End If
            End If
            ComprobarLetraSerie = True
    
        Case "rrecibpozos"
            'cargamos el RSConta con la tabla contadores de BD: Contabilidad
            'donde estan todas las letra de serie que existen en la contabilidad
            Sql = "Select distinct tiporegi from contadores"
            Set RSconta = New ADODB.Recordset
            RSconta.Open Sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
            If RSconta.EOF Then
                RSconta.Close
                Set RSconta = Nothing
                Exit Function
            End If
                
        
            'obtenemos los distintos tipos de movimiento que vamos a contabilizar
            'de las facturas seleccionadas
            Sql = "select distinct rrecibpozos.codtipom from " & cadTabla
            Sql = Sql & " INNER JOIN tmpFactu ON rrecibpozos.codtipom=tmpFactu.codtipom AND rrecibpozos.numfactu=tmpFactu.numfactu AND rrecibpozos.fecfactu=tmpFactu.fecfactu "
    '        SQL = SQL & cadWHERE
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            cad = ""
            B = True
            While Not Rs.EOF And B
                'comprobar que todas las letras serie existen en usuarios
                Sql = "letraser"
                Sql2 = "select letraser from usuarios.stipom where codtipom = " & DBSet(Rs!CodTipom, "T")
                'devuelve = DevuelveDesdeBDNew(cAgro, "stipom", "codtipom", "codtipom", RS!codtipom, "T", SQL)
                Total = TotalRegistrosConsulta(Sql2)
                If Total = 0 Then 'devuelve = "" Then
                    B = False
                    cad = Rs!CodTipom & " en BD de Ariagrorec."
                ElseIf DevuelveValor(Sql2) <> "" Then 'SQL <> "" Then
                    'comprobar que todas las letras serie existen en la contabilidad
                    devuelve = "tiporegi= " & DBSet(DevuelveValor(Sql2), "T") 'SQL, "T")
                    RSconta.MoveFirst
                    RSconta.Find (devuelve), , adSearchForward
                    If RSconta.EOF Then
                        'no encontrado
                        B = False
                        'Cad = SQL & " en BD de Contabilidad."
                        cad = DevuelveValor(Sql2) & " en BD de Contabilidad."
                    End If
                End If
                If B Then cad = cad & DBSet(Rs!CodTipom, "T") & ","
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
            RSconta.Close
            Set RSconta = Nothing
            
            If Not B Then 'Hay algun movimiento que no existe
                devuelve = "No existe el tipo de movimiento: " & cad & vbCrLf
                devuelve = devuelve & "Consulte con el administrador."
                MsgBox devuelve, vbExclamation
                Exit Function
            End If
            
            'Todos los Tipo de movimiento existen
            If cad <> "" Then
                cad = Mid(cad, 1, Len(cad) - 1) 'quitamos ult. coma
            
                'miramos si hay algun movimiento de factura que la letra serie sea nulo
                Sql = "select count(*) from usuarios.stipom "
                Sql = Sql & "where codtipom IN (" & cad & ") and (isnull(letraser) or letraser='')"
                If RegistrosAListar(Sql) > 0 Then
                    Sql = "Hay algun tipo de movimiento de Facturación que no tiene letra serie." & vbCrLf
                    Sql = Sql & "Comprobar en la tabla de tipos de movimiento: " & cad
                    MsgBox Sql, vbExclamation
                    Exit Function
                End If
            End If
            ComprobarLetraSerie = True
    
    
    
    
        Case "rbodfacturas"
            'cargamos el RSConta con la tabla contadores de BD: Contabilidad
            'donde estan todas las letra de serie que existen en la contabilidad
            Sql = "Select distinct tiporegi from contadores"
            Set RSconta = New ADODB.Recordset
            RSconta.Open Sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
            If RSconta.EOF Then
                RSconta.Close
                Set RSconta = Nothing
                Exit Function
            End If
                
        
            'obtenemos los distintos tipos de movimiento que vamos a contabilizar
            'de las facturas seleccionadas
            Sql = "select distinct rbodfacturas.codtipom from " & cadTabla
            Sql = Sql & " INNER JOIN tmpFactu ON rbodfacturas.codtipom=tmpFactu.codtipom AND rbodfacturas.numfactu=tmpFactu.numfactu AND rbodfacturas.fecfactu=tmpFactu.fecfactu "
    '        SQL = SQL & cadWHERE
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            cad = ""
            B = True
            While Not Rs.EOF And B
                'comprobar que todas las letras serie existen en usuarios
                Sql = "letraser"
                Sql2 = "select letraser from usuarios.stipom where codtipom = " & DBSet(Rs!CodTipom, "T")
                'devuelve = DevuelveDesdeBDNew(cAgro, "stipom", "codtipom", "codtipom", RS!codtipom, "T", SQL)
                Total = TotalRegistrosConsulta(Sql2)
                If Total = 0 Then 'devuelve = "" Then
                    B = False
                    cad = Rs!CodTipom & " en BD de Ariagrorec."
                ElseIf DevuelveValor(Sql2) <> "" Then 'SQL <> "" Then
                    'comprobar que todas las letras serie existen en la contabilidad
                    devuelve = "tiporegi= " & DBSet(DevuelveValor(Sql2), "T") 'SQL, "T")
                    RSconta.MoveFirst
                    RSconta.Find (devuelve), , adSearchForward
                    If RSconta.EOF Then
                        'no encontrado
                        B = False
                        'Cad = SQL & " en BD de Contabilidad."
                        cad = DevuelveValor(Sql2) & " en BD de Contabilidad."
                    End If
                End If
                If B Then cad = cad & DBSet(Rs!CodTipom, "T") & ","
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
            RSconta.Close
            Set RSconta = Nothing
            
            If Not B Then 'Hay algun movimiento que no existe
                devuelve = "No existe el tipo de movimiento: " & cad & vbCrLf
                devuelve = devuelve & "Consulte con el administrador."
                MsgBox devuelve, vbExclamation
                Exit Function
            End If
            
            'Todos los Tipo de movimiento existen
            If cad <> "" Then
                cad = Mid(cad, 1, Len(cad) - 1) 'quitamos ult. coma
            
                'miramos si hay algun movimiento de factura que la letra serie sea nulo
                Sql = "select count(*) from usuarios.stipom "
                Sql = Sql & "where codtipom IN (" & cad & ") and (isnull(letraser) or letraser='')"
                If RegistrosAListar(Sql) > 0 Then
                    Sql = "Hay algun tipo de movimiento de Facturación que no tiene letra serie." & vbCrLf
                    Sql = Sql & "Comprobar en la tabla de tipos de movimiento: " & cad
                    MsgBox Sql, vbExclamation
                    Exit Function
                End If
            End If
            ComprobarLetraSerie = True
    
    
        Case "tmpfactvarias"
            Sql = "Select distinct tiporegi from contadores where tiporegi = 'XX1'"
            Set RSconta = New ADODB.Recordset
            RSconta.Open Sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
            If RSconta.EOF Then
                RSconta.Close
                Set RSconta = Nothing
                Exit Function
            End If
            ComprobarLetraSerie = True
    
    
    
    
        Case "rtelmovil"
            'cargamos el RSConta con la tabla contadores de BD: Contabilidad
            'donde estan todas las letra de serie que existen en la contabilidad
            Sql = "Select distinct tiporegi from contadores"
            Set RSconta = New ADODB.Recordset
            RSconta.Open Sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
            If RSconta.EOF Then
                RSconta.Close
                Set RSconta = Nothing
                Exit Function
            End If
                
        
            'obtenemos las distintas letras de serie de las facturas seleccionadas
            Sql = "select distinct rtelmovil.numserie from " & cadTabla
            Sql = Sql & " INNER JOIN tmpFactu ON rtelmovil.numserie=tmpFactu.numserie AND rtelmovil.numfactu=tmpFactu.numfactu AND rtelmovil.fecfactu=tmpFactu.fecfactu "
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            cad = ""
            B = True
            While Not Rs.EOF And B
                'comprobar que todas las letras serie existen en la contabilidad
                devuelve = "tiporegi= " & DBSet(Rs!numserie, "T") 'SQL, "T")
                RSconta.MoveFirst
                RSconta.Find (devuelve), , adSearchForward
                If RSconta.EOF Then
                    'no encontrado
                    B = False
                    'Cad = SQL & " en BD de Contabilidad."
                    cad = Rs!numserie & " en BD de Contabilidad."
                End If
                
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
            RSconta.Close
            Set RSconta = Nothing
            
            If Not B Then 'Hay algun movimiento que no existe
                devuelve = "No existe la letra de serie: " & cad & vbCrLf
                devuelve = devuelve & "Consulte con el administrador."
                MsgBox devuelve, vbExclamation
                Exit Function
            End If
            
            ComprobarLetraSerie = True
    
    
    
        Case "fvarcabfact"
            'cargamos el RSConta con la tabla contadores de BD: Contabilidad
            'donde estan todas las letra de serie que existen en la contabilidad
            Sql = "Select distinct tiporegi from contadores"
            Set RSconta = New ADODB.Recordset
            RSconta.Open Sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
            If RSconta.EOF Then
                RSconta.Close
                Set RSconta = Nothing
                Exit Function
            End If
                
        
            'obtenemos los distintos tipos de movimiento que vamos a contabilizar
            'de las facturas seleccionadas
            Sql = "select distinct fvarcabfact.codtipom from " & cadTabla
            Sql = Sql & " INNER JOIN tmpFactu ON fvarcabfact.codtipom=tmpFactu.codtipom AND fvarcabfact.numfactu=tmpFactu.numfactu AND fvarcabfact.fecfactu=tmpFactu.fecfactu "
    '        SQL = SQL & cadWHERE
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            cad = ""
            B = True
            While Not Rs.EOF And B
                'comprobar que todas las letras serie existen en usuarios
                Sql = "letraser"
                Sql2 = "select letraser from usuarios.stipom where codtipom = " & DBSet(Rs!CodTipom, "T")
                'devuelve = DevuelveDesdeBDNew(cAgro, "stipom", "codtipom", "codtipom", RS!codtipom, "T", SQL)
                Total = TotalRegistrosConsulta(Sql2)
                If Total = 0 Then 'devuelve = "" Then
                    B = False
                    cad = Rs!CodTipom & " en BD de Ariagrorec."
                ElseIf DevuelveValor(Sql2) <> "" Then 'SQL <> "" Then
                    'comprobar que todas las letras serie existen en la contabilidad
                    devuelve = "tiporegi= " & DBSet(DevuelveValor(Sql2), "T") 'SQL, "T")
                    RSconta.MoveFirst
                    RSconta.Find (devuelve), , adSearchForward
                    If RSconta.EOF Then
                        'no encontrado
                        B = False
                        'Cad = SQL & " en BD de Contabilidad."
                        cad = DevuelveValor(Sql2) & " en BD de Contabilidad."
                    End If
                End If
                If B Then cad = cad & DBSet(Rs!CodTipom, "T") & ","
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
            RSconta.Close
            Set RSconta = Nothing
            
            If Not B Then 'Hay algun movimiento que no existe
                devuelve = "No existe el tipo de movimiento: " & cad & vbCrLf
                devuelve = devuelve & "Consulte con el administrador."
                MsgBox devuelve, vbExclamation
                Exit Function
            End If
            
            'Todos los Tipo de movimiento existen
            If cad <> "" Then
                cad = Mid(cad, 1, Len(cad) - 1) 'quitamos ult. coma
            
                'miramos si hay algun movimiento de factura que la letra serie sea nulo
                Sql = "select count(*) from usuarios.stipom "
                Sql = Sql & "where codtipom IN (" & cad & ") and (isnull(letraser) or letraser='')"
                If RegistrosAListar(Sql) > 0 Then
                    Sql = "Hay algun tipo de movimiento de Facturación que no tiene letra serie." & vbCrLf
                    Sql = Sql & "Comprobar en la tabla de tipos de movimiento: " & cad
                    MsgBox Sql, vbExclamation
                    Exit Function
                End If
            End If
            ComprobarLetraSerie = True
    
    
    End Select

EComprobarLetra:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Letra Serie", Err.Description
    End If
End Function



Public Function ComprobarNumFacturas_new(cadTabla As String, cadWConta) As Boolean
'Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
'vamos a contabilizar
Dim Sql As String
Dim SQLconta As String
Dim Rs As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim B As Boolean

    On Error GoTo ECompFactu

    ComprobarNumFacturas_new = False
    
    SQLconta = "SELECT count(*) FROM cabfact WHERE "
 
    
        'Seleccionamos las distintas facturas que vamos a facturar
    If cadTabla = "rtelmovil" Then
        Sql = "SELECT DISTINCT " & cadTabla & ".numserie," & cadTabla & ".numfactu," & cadTabla & ".fecfactu "
        Sql = Sql & " FROM " & cadTabla
        Sql = Sql & " INNER JOIN tmpFactu ON " & cadTabla & ".numserie=tmpFactu.numserie AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "

        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        B = True
        While Not Rs.EOF And B
            If vParamAplic.ContabilidadNueva Then
                Sql = "(numserie= " & DBSet(Rs!numserie, "T") & " AND numfactu=" & DBSet(Rs!numfactu, "N") & " AND anofactu=" & Year(Rs!fecfactu) & ")"
            Else
                Sql = "(numserie= " & DBSet(Rs!numserie, "T") & " AND codfaccl=" & DBSet(Rs!numfactu, "N") & " AND anofaccl=" & Year(Rs!fecfactu) & ")"
            End If
'            If SituarRSetMULTI(RSconta, SQL) Then
            Sql = SQLconta & Sql
            If RegistrosAListar(Sql, cConta) Then
                B = False
                Sql = "          Letra Serie: " & DBSet(Rs!numserie, "T") & vbCrLf
                Sql = Sql & "          Nº Fac.: " & Format(Rs!numfactu, "0000000") & vbCrLf
                Sql = Sql & "          Fecha: " & Rs!fecfactu
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not B Then
            Sql = "Ya existe la factura: " & vbCrLf & Sql
            Sql = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & Sql
            
            MsgBox Sql, vbExclamation
            ComprobarNumFacturas_new = False
        Else
            ComprobarNumFacturas_new = True
        End If
     
     
     Else
        Sql = "SELECT DISTINCT " & cadTabla & ".codtipom,letraser," & cadTabla & ".numfactu," & cadTabla & ".fecfactu "
        Sql = Sql & " FROM (" & cadTabla & " INNER JOIN usuarios.stipom ON " & cadTabla & ".codtipom=usuarios.stipom.codtipom) "
        Sql = Sql & " INNER JOIN tmpFactu ON " & cadTabla & ".codtipom=tmpFactu.codtipom AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "

        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        B = True
        While Not Rs.EOF And B
            If vParamAplic.ContabilidadNueva Then
                Sql = "(numserie= " & DBSet(Rs!letraser, "T") & " AND numfactu=" & DBSet(Rs!numfactu, "N") & " AND anofactu=" & Year(Rs!fecfactu) & ")"
            Else
                Sql = "(numserie= " & DBSet(Rs!letraser, "T") & " AND codfaccl=" & DBSet(Rs!numfactu, "N") & " AND anofaccl=" & Year(Rs!fecfactu) & ")"
            End If
'            If SituarRSetMULTI(RSconta, SQL) Then
            Sql = SQLconta & Sql
            If RegistrosAListar(Sql, cConta) Then
                B = False
                Sql = "          Letra Serie: " & DBSet(Rs!letraser, "T") & vbCrLf
                Sql = Sql & "          Nº Fac.: " & Format(Rs!numfactu, "0000000") & vbCrLf
                Sql = Sql & "          Fecha: " & Rs!fecfactu
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not B Then
            Sql = "Ya existe la factura: " & vbCrLf & Sql
            Sql = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & Sql
            
            MsgBox Sql, vbExclamation
            ComprobarNumFacturas_new = False
        Else
            ComprobarNumFacturas_new = True
        End If
    End If
    Exit Function
    
ECompFactu:
     If Err.Number <> 0 Then
        ComprobarNumFacturas_new = False
        MuestraError Err.Number, "Comprobar Nº Facturas", Err.Description
    End If
End Function




Public Function ComprobarCtaContable_new(cadTabla As String, Opcion As Byte, Optional Tipo As Byte, Optional Seccion As Integer, Optional cuenta As String) As Boolean
'Comprobar que todas las ctas contables de los distintos clientes de las facturas
'que vamos a contabilizar existan en la contabilidad
Dim Sql As String
Dim Rs As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim B As Boolean
Dim cadG As String
Dim SQLcuentas As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigit3 As String


    On Error GoTo ECompCta

    ComprobarCtaContable_new = False

    cadG = ""
    If Opcion = 7 Or Opcion = 9 Or Opcion = 10 Or Opcion = 11 Then
        'si hay analitica comprobar que todas las cuentas
        'empiezan por el digito que hay en conta.parametros.grupogto o .grupovta
        cadG = "grupovta"
        Sql = DevuelveDesdeBDNew(cConta, "parametros", "grupogto", "", "", "", cadG)
        If Sql <> "" And cadG <> "" Then
            Sql = " AND (codmacta like '" & Sql & "%' OR codmacta like '" & cadG & "%')"
        ElseIf Sql <> "" Then
            Sql = " AND (codmacta like '" & Sql & "%')"
        ElseIf cadG <> "" Then
            Sql = " AND (codmacta like '" & cadG & "%')"
        End If
        cadG = Sql
    End If


    SQLcuentas = "SELECT count(*) FROM cuentas WHERE apudirec='S' "
    If cadG <> "" Then SQLcuentas = SQLcuentas & cadG

    If Opcion = 1 Then
        If cadTabla = "rfactsoc" Then
            'Seleccionamos los distintos socios, cuentas que vamos a facturar
            Sql = "SELECT DISTINCT rfactsoc.codsocio, rsocios_seccion.codmacpro as codmacta "
            Sql = Sql & " FROM (rfactsoc INNER JOIN rsocios_seccion ON rfactsoc.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N") & ") "
            Sql = Sql & " INNER JOIN tmpFactu ON rfactsoc.codtipom=tmpFactu.codtipom AND rfactsoc.numfactu=tmpFactu.numfactu AND rfactsoc.fecfactu=tmpFactu.fecfactu "
        Else
            If cadTabla = "rcabfactalmz" Then
                'Seleccionamos los distintos socios, cuentas que vamos a facturar
                Sql = "SELECT DISTINCT rcabfactalmz.codsocio, rsocios_seccion.codmacpro as codmacta "
                Sql = Sql & " FROM (rcabfactalmz INNER JOIN rsocios_seccion ON rcabfactalmz.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N") & ") "
                Sql = Sql & " INNER JOIN tmpFactu ON rcabfactalmz.tipofichero=tmpFactu.tipofichero AND rcabfactalmz.numfactu=tmpFactu.numfactu AND rcabfactalmz.fecfactu=tmpFactu.fecfactu "
                Sql = Sql & " and rcabfactalmz.codsocio = tmpFactu.codsocio "
                
                '[Monica]29/07/2015: en el caso de catadau si es asociado tengo que mirar raiz asociado + codigo asociado
                '                                           si es socio entonces raiz socio + codigo socio
                If vParamAplic.Cooperativa = 0 Then
                    Sql = "SELECT DISTINCT rcabfactalmz.codsocio, if(rsocios.tiporelacion = 1, concat(rseccion.raiz_cliente_asociado,right(concat('00000',rsocios.nroasociado),5)), rsocios_seccion.codmacpro) as codmacta "
                    Sql = Sql & " FROM (((rcabfactalmz INNER JOIN rsocios_seccion ON rcabfactalmz.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N") & ") "
                    Sql = Sql & " INNER JOIN tmpFactu ON rcabfactalmz.tipofichero=tmpFactu.tipofichero AND rcabfactalmz.numfactu=tmpFactu.numfactu AND rcabfactalmz.fecfactu=tmpFactu.fecfactu "
                    Sql = Sql & " and rcabfactalmz.codsocio = tmpFactu.codsocio) "
                    Sql = Sql & " INNER JOIN rsocios ON rcabfactalmz.codsocio = rsocios.codsocio) "
                    Sql = Sql & " INNER JOIN rseccion on rseccion.codsecci = rsocios_seccion.codsecci "
                End If
            Else
                If cadTabla = "advfacturas" Then
                    'Seleccionamos los distintos socios, cuentas que vamos a facturar
                    Sql = "SELECT DISTINCT advfacturas.codsocio, rsocios_seccion.codmaccli as codmacta "
                    Sql = Sql & " FROM (advfacturas INNER JOIN rsocios_seccion ON advfacturas.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionADV, "N") & ") "
                    Sql = Sql & " INNER JOIN tmpFactu ON advfacturas.codtipom=tmpFactu.codtipom AND advfacturas.numfactu=tmpFactu.numfactu AND advfacturas.fecfactu=tmpFactu.fecfactu "
'                    SQL = SQL & " and advfacturas.codsocio = tmpFactu.codsocio "
                Else ' facturas de retirada de almazara
                    If cadTabla = "rbodfact1" Then
                        'Seleccionamos los distintos socios, cuentas que vamos a facturar
                        Sql = "SELECT DISTINCT rbodfacturas.codsocio, rsocios_seccion.codmaccli as codmacta "
                        Sql = Sql & " FROM (rbodfacturas INNER JOIN rsocios_seccion ON rbodfacturas.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N") & ") "
                        Sql = Sql & " INNER JOIN tmpFactu ON rbodfacturas.codtipom=tmpFactu.codtipom AND rbodfacturas.numfactu=tmpFactu.numfactu AND rbodfacturas.fecfactu=tmpFactu.fecfactu "
                    Else ' facturas de retirada de bodega
                        If cadTabla = "rbodfact2" Then
                            'Seleccionamos los distintos socios, cuentas que vamos a facturar
                            Sql = "SELECT DISTINCT rbodfacturas.codsocio, rsocios_seccion.codmaccli as codmacta "
                            Sql = Sql & " FROM (rbodfacturas INNER JOIN rsocios_seccion ON rbodfacturas.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionBodega, "N") & ") "
                            Sql = Sql & " INNER JOIN tmpFactu ON rbodfacturas.codtipom=tmpFactu.codtipom AND rbodfacturas.numfactu=tmpFactu.numfactu AND rbodfacturas.fecfactu=tmpFactu.fecfactu "
                        Else 'facturas de telefonia
                            If cadTabla = "rtelmovil" Then
                                Sql = "SELECT DISTINCT rtelmovil.codsocio, rsocios_seccion.codmaccli as codmacta "
                                Sql = Sql & " FROM (rtelmovil INNER JOIN rsocios_seccion ON rtelmovil.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.Seccionhorto, "N") & ") "
                                Sql = Sql & " INNER JOIN tmpFactu ON rtelmovil.numserie=tmpFactu.numserie AND rtelmovil.numfactu=tmpFactu.numfactu AND rtelmovil.fecfactu=tmpFactu.fecfactu "
                            Else
                                If cadTabla = "rrecibpozos" Then
                                    Sql = "SELECT DISTINCT rrecibpozos.codsocio, rsocios_seccion.codmaccli as codmacta "
                                    Sql = Sql & " FROM (rrecibpozos INNER JOIN rsocios_seccion ON rrecibpozos.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionPOZOS, "N") & ") "
                                    Sql = Sql & " INNER JOIN tmpFactu ON rrecibpozos.codtipom=tmpFactu.codtipom AND rrecibpozos.numfactu=tmpFactu.numfactu AND rrecibpozos.fecfactu=tmpFactu.fecfactu "
                                Else
                                    If cadTabla = "rfacttra" Then
                                        'Seleccionamos los distintos socios, cuentas que vamos a facturar
                                        Sql = "SELECT DISTINCT rfacttra.codtrans, rtransporte.codmacpro as codmacta "
                                        Sql = Sql & " FROM (rfacttra INNER JOIN rtransporte ON rfacttra.codtrans=rtransporte.codtrans) "
                                        Sql = Sql & " INNER JOIN tmpFactu ON rfacttra.codtipom=tmpFactu.codtipom AND rfacttra.numfactu=tmpFactu.numfactu AND rfacttra.fecfactu=tmpFactu.fecfactu "
                                    Else
                                        If cadTabla = "fvarcabfact" Then
                                            If Tipo = 0 Then ' seleccionamos primero los socios
                                                'Seleccionamos los distintos socios de facturas varias, cuentas que vamos a facturar
                                                Sql = "SELECT DISTINCT fvarcabfact.codsocio, rsocios_seccion.codmaccli as codmacta "
                                                Sql = Sql & " FROM (fvarcabfact INNER JOIN rsocios_seccion ON fvarcabfact.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N") & " and not fvarcabfact.codsocio is null and fvarcabfact.codsocio <> 0 ) "
                                                Sql = Sql & " INNER JOIN tmpFactu ON fvarcabfact.codtipom=tmpFactu.codtipom AND fvarcabfact.numfactu=tmpFactu.numfactu AND fvarcabfact.fecfactu=tmpFactu.fecfactu "
                                            Else
                                                'Seleccionamos los distintos clientes de facturas varias, cuentas que vamos a facturar
                                                Sql = "SELECT DISTINCT fvarcabfact.codclien, clientes.codmacta as codmacta "
                                                Sql = Sql & " FROM (fvarcabfact INNER JOIN clientes ON fvarcabfact.codclien=clientes.codclien and not fvarcabfact.codclien is null and fvarcabfact.codclien <> 0) "
                                                Sql = Sql & " INNER JOIN tmpFactu ON fvarcabfact.codtipom=tmpFactu.codtipom AND fvarcabfact.numfactu=tmpFactu.numfactu AND fvarcabfact.fecfactu=tmpFactu.fecfactu "
                                            End If
                                        Else
                                            If cadTabla = "fvarcabfactpro" Then
                                                'Seleccionamos los distintos socios de facturas varias, cuentas que vamos a facturar
                                                Sql = "SELECT DISTINCT fvarcabfactpro.codsocio, rsocios_seccion.codmacpro as codmacta "
                                                Sql = Sql & " FROM (fvarcabfactpro INNER JOIN rsocios_seccion ON fvarcabfactpro.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N") & " ) "
                                                Sql = Sql & " INNER JOIN tmpFactu ON fvarcabfactpro.codtipom=tmpFactu.codtipom AND fvarcabfactpro.numfactu=tmpFactu.numfactu AND fvarcabfactpro.fecfactu=tmpFactu.fecfactu "
                                            Else
                                                If cadTabla = "tmpfactvarias" Then
                                                    If Tipo = 0 Then ' seleccionamos primero los socios
                                                        'Seleccionamos los distintos socios de facturas varias, cuentas que vamos a facturar
                                                        Sql = "SELECT DISTINCT tmpfactvarias.codsoccli, rsocios_seccion.codmaccli as codmacta "
                                                        Sql = Sql & " FROM (tmpfactvarias INNER JOIN rsocios_seccion ON tmpfactvarias.codsoccli=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N") & " and not tmpfactvarias.codsoccli is null and tmpfactvarias.codsoccli <> 0 and tmpfactvarias.codusu = " & vUsu.Codigo & ") "
                                                    Else
                                                        'Seleccionamos los distintos clientes de facturas varias, cuentas que vamos a facturar
                                                        Sql = "SELECT DISTINCT tmpfactvarias.codsoccli, clientes.codmacta as codmacta "
                                                        Sql = Sql & " FROM (tmpfactvarias INNER JOIN clientes ON tmpfactvarias.codsoccli=clientes.codclien and not tmpfactvarias.codsoccli is null and tmpfactvarias.codsoccli <> 0 and tmpfactvarias.codusu = " & vUsu.Codigo & ") "
                                                    End If
                                                Else
                                                    'Seleccionamos los distintos socios terceros, cuentas que vamos a facturar
                                                    Sql = "SELECT DISTINCT rcafter.codsocio, rsocios_seccion.codmacpro as codmacta "
                                                    Sql = Sql & " FROM (rcafter INNER JOIN rsocios_seccion ON rcafter.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & vParamAplic.Seccionhorto & ") "
                                                    Sql = Sql & " INNER JOIN tmpFactu ON rcafter.codsocio=tmpFactu.codsocio AND rcafter.numfactu=tmpFactu.numfactu AND rcafter.fecfactu=tmpFactu.fecfactu "
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    ElseIf Opcion = 8 Then
        Sql = "SELECT distinct "
        If cadTabla = "rfactsoc" Then
            Select Case Tipo
                Case 1, 3, 7, 9  ' 1=anticipos , 3=anticipos venta campo, 7=anticipos almazara, 9=anticipo bodega
                    Sql = Sql & " rfactsoc_variedad.codvarie, variedades.ctaanticipo as codmacta from ((rfactsoc_variedad "
                Case 2, 4, 8, 10, 11 ' 2=liquidaciones, 4=liquidaciones venta campo, 8=liquidacion almazara, 10=liquidacion bodega
                    Sql = Sql & " rfactsoc_variedad.codvarie, variedades.ctaliquidacion as codmacta from ((rfactsoc_variedad "
                Case 6  '6=siniestros
                    Sql = Sql & " rfactsoc_variedad.codvarie, variedades.ctasiniestros as codmacta from ((rfactsoc_variedad "
                Case 12 ' le paso un tipo 12 a las liquidaciones de industria de terceros para comprobar las cuentas de terceros
                    Sql = Sql & " rfactsoc_variedad.codvarie, variedades.ctacomtercero as codmacta from ((rfactsoc_variedad "
                Case 13 ' facturas de acarreo recoleccion socio FTS
                    Sql = Sql & " rfactsoc_variedad.codvarie, variedades.ctaacarecol as codmacta from ((rfactsoc_variedad "
            End Select
            Sql = Sql & " INNER JOIN tmpFactu ON rfactsoc_variedad.codtipom=tmpFactu.codtipom AND rfactsoc_variedad.numfactu=tmpFactu.numfactu AND rfactsoc_variedad.fecfactu=tmpFactu.fecfactu) "
            Sql = Sql & "INNER JOIN variedades ON rfactsoc_variedad.codvarie=variedades.codvarie) "
        Else
            If cadTabla = "rfacttra" Then
                Sql = Sql & " rfacttra_albaran.codvarie, variedades.ctatransporte as codmacta from ((rfacttra_albaran "
                Sql = Sql & " INNER JOIN tmpFactu ON rfacttra_albaran.codtipom=tmpFactu.codtipom AND rfacttra_albaran.numfactu=tmpFactu.numfactu AND rfacttra_albaran.fecfactu=tmpFactu.fecfactu) "
                Sql = Sql & "INNER JOIN variedades ON rfacttra_albaran.codvarie=variedades.codvarie) "
            Else
                Sql = Sql & " rlifter.codvarie, variedades.ctacomtercero as codmacta from ((rlifter "
                Sql = Sql & " INNER JOIN tmpFactu ON rlifter.codsocio=tmpFactu.codsocio AND rlifter.numfactu=tmpFactu.numfactu AND rlifter.fecfactu=tmpFactu.fecfactu) "
                Sql = Sql & "INNER JOIN variedades ON rlifter.codvarie=variedades.codvarie) "
            End If
        End If
    ElseIf Opcion = 2 Then
            If cadTabla = "advfacturas" Then
                Sql = "SELECT distinct advartic.codfamia,"
                Sql = Sql & " advfamia.ctaventa as codmacta,advfamia.aboventa as ctaabono from ((advfacturas_lineas "
                Sql = Sql & " INNER JOIN tmpFactu ON advfacturas_lineas.codtipom=tmpFactu.codtipom AND advfacturas_lineas.numfactu=tmpFactu.numfactu AND advfacturas_lineas.fecfactu=tmpFactu.fecfactu) "
                Sql = Sql & "INNER JOIN advartic ON advfacturas_lineas.codartic=advartic.codartic) "
                Sql = Sql & "INNER JOIN advfamia ON advartic.codfamia = advfamia.codfamia "
            Else
                If cadTabla = "rbodfacturas" Then
                    Sql = "SELECT distinct rbodfacturas_lineas.codvarie, variedades.ctavtasotros as codmacta from (rbodfacturas_lineas "
                    Sql = Sql & " INNER JOIN tmpFactu ON rbodfacturas_lineas.codtipom=tmpFactu.codtipom AND rbodfacturas_lineas.numfactu=tmpFactu.numfactu AND rbodfacturas_lineas.fecfactu=tmpFactu.fecfactu) "
                    Sql = Sql & " INNER JOIN variedades ON rbodfacturas_lineas.codvarie = variedades.codvarie "
                Else
                    If cadTabla = "rbodfact1" Then ' rbodfacturas de almazara "FZA"
                        Sql = "select distinct " & vParamAplic.CtaVentasAlmz & " as codmacta "
                        Sql = Sql & " FROM rbodfacturas "
                    Else
                        If cadTabla = "rbodfact2" Then ' rbodfacturas de bodega "FAB"
                            Sql = "select distinct " & vParamAplic.CtaVentasBOD & " as codmacta "
                            Sql = Sql & " FROM rbodfacturas "
                        Else
                            If cadTabla = "rtelmovil" Then
                                Sql = "select distinct " & vParamAplic.CtaVentasTel & " as codmacta "
                                Sql = Sql & " FROM rtelmovil "
                            Else
                                If cadTabla = "rrecibpozos" Then
                                    Select Case Tipo
                                        Case 1   ' cuenta de ventas de consumo
                                            Sql = "select distinct " & vParamAplic.CtaVentasConsPOZ & " as codmacta "
                                            Sql = Sql & " FROM rrecibpozos "
                                        Case 2   ' cuenta de ventas de cuotas
                                            Sql = "select distinct " & vParamAplic.CtaVentasCuoPOZ & " as codmacta "
                                            Sql = Sql & " FROM rrecibpozos "
                                        Case 3   ' cuenta de ventas de talla
                                            Sql = "select distinct " & vParamAplic.CtaVentasTalPOZ & " as codmacta "
                                            Sql = Sql & " FROM rrecibpozos "
                                        Case 4   ' cuenta de ventas de mantenimiento
                                            Sql = "select distinct " & vParamAplic.CtaVentasMtoPOZ & " as codmacta "
                                            Sql = Sql & " FROM rrecibpozos "
                                        Case 5   ' cuenta de ventas de consumo a manta
                                            Sql = "select distinct " & vParamAplic.CtaVentasMantaPOZ & " as codmacta "
                                            Sql = Sql & " FROM rrecibpozos "
                                        '[Monica]21/01/2016: cuenta de recargos
                                        Case 6   ' cuenta de recargos
                                            Sql = "select distinct " & vParamAplic.CtaRecargosPOZ & " as codmacta "
                                            Sql = Sql & " FROM rrecibpozos "
                                    End Select
                                Else
                                    If cadTabla = "fvarcabfact" Then
                                        Sql = "select distinct fvarconce.codmacta as codmacta "
                                        Sql = Sql & " FROM ((fvarlinfact "
                                        Sql = Sql & " INNER JOIN tmpFactu ON fvarlinfact.codtipom=tmpFactu.codtipom AND fvarlinfact.numfactu=tmpFactu.numfactu AND fvarlinfact.fecfactu=tmpFactu.fecfactu) "
                                        Sql = Sql & "INNER JOIN fvarconce ON fvarlinfact.codconce=fvarconce.codconce) "
                                    Else
                                        If cadTabla = "fvarcabfactpro" Then
                                            Sql = "select distinct fvarconce.codmacpr as codmacta "
                                            Sql = Sql & " FROM ((fvarlinfactpro "
                                            Sql = Sql & " INNER JOIN tmpFactu ON fvarlinfactpro.codtipom=tmpFactu.codtipom AND fvarlinfactpro.numfactu=tmpFactu.numfactu AND fvarlinfactpro.fecfactu=tmpFactu.fecfactu) "
                                            Sql = Sql & "INNER JOIN fvarconce ON fvarlinfactpro.codconce=fvarconce.codconce) "
                                        Else
                                            If cadTabla = "tmpfactvarias" Then
                                                Sql = "select distinct fvarconce.codmacta as codmacta "
                                                Sql = Sql & " FROM (tmpfactvarias "
                                                Sql = Sql & "INNER JOIN fvarconce ON tmpfactvarias.codconce=fvarconce.codconce) "
                                            Else
                                                Sql = "select distinct " & vParamAplic.CtaVentasAlmz & " as codmacta "
                                                Sql = Sql & " FROM rcabfactalmz "
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
    ElseIf Opcion = 3 Then
            Sql = "select distinct " & vParamAplic.CtaGastosAlmz & " as codmacta "
            Sql = Sql & " FROM rcabfactalmz "
    ElseIf Opcion = 4 Then
        Sql = "select distinct " & DBSet(vParamAplic.CtaTerReten, "T") & " as codmacta from tcafpc "
    ElseIf Opcion = 7 Then
            If cadTabla = "rfactsoc" Then
                Select Case Tipo
                    Case 1, 3, 7, 9  ' 1=anticipos , 3=anticipos venta campo, 7=anticipos almazara, 9=anticipos bodega
                        Sql = " SELECT variedades.ctaanticipo as cuenta "
                        Sql = Sql & " FROM rfactsoc_variedad, variedades, tmpFactu  WHERE "
                        Sql = Sql & " rfactsoc_variedad.codtipom=tmpFactu.codtipom and rfactsoc_variedad.numfactu=tmpFactu.numfactu and rfactsoc_variedad.fecfactu=tmpFactu.fecfactu and "
                        Sql = Sql & " rfactsoc_variedad.codvarie=variedades.codvarie "
                        Sql = Sql & " group by 1 "
                    Case 2, 4, 6, 8, 10 ' 2=liquidaciones, 4=liquidaciones venta campo, 6=siniestros, 8=liquidacion almazara, 10=liquidacion bodega
                        Sql = " SELECT variedades.ctaliquidacion as cuenta "
                        Sql = Sql & " FROM rfactsoc_variedad, variedades, tmpFactu  WHERE "
                        Sql = Sql & " rfactsoc_variedad.codtipom=tmpFactu.codtipom and rfactsoc_variedad.numfactu=tmpFactu.numfactu and rfactsoc_variedad.fecfactu=tmpFactu.fecfactu and "
                        Sql = Sql & " rfactsoc_variedad.codvarie=variedades.codvarie "
                        Sql = Sql & " group by 1 "
                End Select
            Else
                If cadTabla = "advfacturas" Then
                    Sql = "SELECT distinct advartic.codfamia,"
                    Sql = Sql & " advfamia.ctaventa as cuenta,advfamia.aboventa as ctaabono from ((advfacturas_lineas "
                    Sql = Sql & " INNER JOIN tmpFactu ON advfacturas_lineas.codtipom=tmpFactu.codtipom AND advfacturas_lineas.numfactu=tmpFactu.numfactu AND advfacturas_lineas.fecfactu=tmpFactu.fecfactu) "
                    Sql = Sql & "INNER JOIN advartic ON advfacturas_lineas.codartic=advartic.codartic) "
                    Sql = Sql & "INNER JOIN advfamia ON advartic.codfamia = advfamia.codfamia "
                Else
                    If cadTabla = "rbodfacturas" Then
                        Sql = "SELECT distinct "
                        Sql = Sql & " variedades.ctavtasotros as cuenta from (rbodfacturas_lineas "
                        Sql = Sql & " INNER JOIN tmpFactu ON rbodfacturas_lineas.codtipom=tmpFactu.codtipom AND rbodfacturas_lineas.numfactu=tmpFactu.numfactu AND rbodfacturas_lineas.fecfactu=tmpFactu.fecfactu) "
                        Sql = Sql & "INNER JOIN variedades ON rbodfacturas_lineas.codvarie=variedades.codvarie "
                    Else
                        If cadTabla = "rbodfact1" Then ' facturas de retirada de almazara
                            Sql = "select distinct " & vParamAplic.CtaVentasAlmz & " as cuenta "
                            Sql = Sql & " FROM rbodfacturas "
                        Else
                            If cadTabla = "rbodfact2" Then ' facturas de retirada de bodega
                                Sql = "select distinct " & vParamAplic.CtaVentasBOD & " as cuenta "
                                Sql = Sql & " FROM rbodfacturas "
                            Else
                                If cadTabla = "rtelmovil" Then
                                    Sql = "select distinct " & vParamAplic.CtaVentasTel & " as cuenta "
                                    Sql = Sql & " FROM rtelmovil "
                                Else
                                    If cadTabla = "rrecibpozos" Then
                                        Sql = "select distinct " & vParamAplic.CtaVentasConsPOZ & " as cuenta "
                                        Sql = Sql & " FROM rrecibpozos "
                                    Else
                                        If cadTabla = "rfacttra" Then
                                            Sql = " SELECT variedades.ctatransporte as cuenta "
                                            Sql = Sql & " FROM rfacttra_albaran, variedades, tmpFactu  WHERE "
                                            Sql = Sql & " rfacttra_albaran.codtipom=tmpFactu.codtipom and rfacttra_albaran.numfactu=tmpFactu.numfactu and rfacttra_albaran.fecfactu=tmpFactu.fecfactu and "
                                            Sql = Sql & " rfacttra_albaran.codvarie=variedades.codvarie "
                                            Sql = Sql & " group by 1 "
                                        Else
                                            If cadTabla = "fvarcabfact" Then
                                                Sql = "SELECT distinct "
                                                Sql = Sql & " fvarconce.codmacta as cuenta from (fvarlinfact "
                                                Sql = Sql & " INNER JOIN tmpFactu ON fvarlinfact.codtipom=tmpFactu.codtipom AND fvarlinfact.numfactu=tmpFactu.numfactu AND fvarlinfact.fecfactu=tmpFactu.fecfactu) "
                                                Sql = Sql & "INNER JOIN fvarconce ON fvarlinfact.codconce=fvarconce.codconce "
                                            Else
                                                If cadTabla = "fvarcabfactpro" Then
                                                    Sql = "SELECT distinct "
                                                    Sql = Sql & " fvarconce.codmacpr as cuenta from (fvarlinfactpro "
                                                    Sql = Sql & " INNER JOIN tmpFactu ON fvarlinfactpro.codtipom=tmpFactu.codtipom AND fvarlinfactpro.numfactu=tmpFactu.numfactu AND fvarlinfactpro.fecfactu=tmpFactu.fecfactu) "
                                                    Sql = Sql & "INNER JOIN fvarconce ON fvarlinfactpro.codconce=fvarconce.codconce "
                                                Else
                                            '       terceros
                                                    Sql = " SELECT variedades.ctacomtercero as cuenta "
                                                    Sql = Sql & " FROM rlifter, variedades, tmpFactu  WHERE "
                                                    Sql = Sql & " rlifter.codsocio=tmpFactu.codsocio and rlifter.numfactu=tmpFactu.numfactu and rlifter.fecfactu=tmpFactu.fecfactu and "
                                                    Sql = Sql & " rlifter.codvarie=variedades.codvarie "
                                                    Sql = Sql & " group by 1 "
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                   End If
                End If
            End If
    ElseIf Opcion = 9 Then
            Sql = " select distinct " & vParamAplic.CtaVentasAlmz & " as cuenta "
            Sql = Sql & " from tmpFactu "
    ElseIf Opcion = 11 Then
            Sql = " select distinct " & vParamAplic.CtaGastosAlmz & " as cuenta "
            Sql = Sql & " from tmpFactu "
    ElseIf Opcion = 12 Then
            Sql = "SELECT rconcepgasto.codmacta as cuenta "
            Sql = Sql & " from (rconcepgasto INNER JOIN rfactsoc_gastos  ON rconcepgasto.codgasto = rfactsoc_gastos.codgasto) "
            Sql = Sql & " INNER JOIN tmpFactu ON rfactsoc_gastos.codtipom=tmpFactu.codtipom AND rfactsoc_gastos.numfactu=tmpFactu.numfactu AND rfactsoc_gastos.fecfactu=tmpFactu.fecfactu "
            
            '[Monica]06/06/2016: si es catadau y no hay cuenta contable no comprobamos nada
            If vParamAplic.Cooperativa = 0 Then
                Sql = Sql & " where not rconcepgasto.codmacta is null and rconcepgasto.codmacta <> '' "
            End If
    ElseIf Opcion = 13 Then
        Sql = "SELECT distinct "
        Sql = Sql & " rcafter.concepcargo, fvarconce.codmacpr as codmacta from ((rcafter "
        Sql = Sql & " INNER JOIN tmpFactu ON rcafter.codsocio=tmpFactu.codsocio AND rcafter.numfactu=tmpFactu.numfactu AND rcafter.fecfactu=tmpFactu.fecfactu) "
        Sql = Sql & "INNER JOIN fvarconce ON rcafter.concepcargo=fvarconce.codconce) "
    ElseIf Opcion = 14 Then
        'Seleccionamos los distintos socios asociados , cuentas que vamos a facturar
        Sql = "SELECT DISTINCT rfactsoc.codsocio, replace(codmacpro,mid(codmacpro,1,length(rseccion.raiz_cliente_socio)), rseccion.raiz_cliente_asociado) as codmacta "
        Sql = Sql & " FROM (((rfactsoc INNER JOIN rsocios_seccion ON rfactsoc.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N") & ") INNER JOIN rseccion ON rsocios_seccion.codsecci = rseccion.codsecci) INNER JOIN rsocios ON rsocios.codsocio = rsocios_seccion.codsocio and rsocios.tiporelacion = 1) "
        Sql = Sql & " INNER JOIN tmpFactu ON rfactsoc.codtipom=tmpFactu.codtipom AND rfactsoc.numfactu=tmpFactu.numfactu AND rfactsoc.fecfactu=tmpFactu.fecfactu "
    
    '[Monica]09/05/2017: cuenta de aportaciones
    ElseIf Opcion = 15 Then
        Sql = " select distinct " & DBSet(cuenta, "T") & " as cuenta "
        Sql = Sql & " from rparam "
    End If

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    B = True

    While Not Rs.EOF And B
        If Opcion < 4 Or Opcion = 8 Or Opcion = 13 Or Opcion = 14 Then
            Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!Codmacta, "T")
        ElseIf Opcion = 4 Then
            Sql = SQLcuentas & " AND codmacta= " & DBSet(vParamAplic.CtaTerReten, "T")
        ElseIf Opcion = 7 Then
            Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!cuenta, "T")
        ElseIf Opcion = 9 Or Opcion = 10 Or Opcion = 11 Or Opcion = 12 Or Opcion = 15 Then
            Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!cuenta, "T")
        End If


        If Not (RegistrosAListar(Sql, cConta) > 0) Then
        'si no lo encuentra
            B = False 'no encontrado
            If Opcion = 1 Then
                If cadTabla = "facturas" Then
                    Sql = DBLet(Rs!Codmacta, "T") & " del Socio " & Format(Rs!CodClien, "000000")
                Else
                    If cadTabla = "rfacttra" Then
                        Sql = DBLet(Rs!Codmacta, "T") & " del transportista " & DBLet(Rs!codTrans, "T")
                    Else
                        If cadTabla = "rfactsoc" Or cadTabla = "advfacturas" Or cadTabla = "rbodfact1" Or cadTabla = "rbodfact2" Or cadTabla = "rtelmovil" Or cadTabla = "rrecibpozos" Or cadTabla = "fvarcabfact" Or cadTabla = "fvarcabfactpro" Then
                            Sql = DBLet(Rs!Codmacta, "T") & " del Socio " & Format(Rs!Codsocio, "000000")
                        Else
                            If cadTabla = "tmpfactvarias" Then
                                Sql = DBLet(Rs!Codmacta, "T") & " del Socio " & Format(Rs!Codsoccli, "000000")
                            
                            Else
                                Sql = DBLet(Rs!Codmacta, "T") & " del Socio " & Format(Rs!Codsocio, "000000")
                            End If
                        End If
                    End If
                End If
            ElseIf Opcion = 2 Then
                If cadTabla = "advfacturas" Then
                    Sql = DBLet(Rs!Codmacta, "T") & " de la familia " & DBLet(Rs!codfamia, "N")
                Else
                    If cadTabla = "rbodfacturas" Then
                        Sql = DBLet(Rs!Codmacta, "T") & " de la variedad " & DBLet(Rs!codvarie, "N")
                    Else
                        If cadTabla = "rbodfact1" Then
                            Sql = DBLet(Rs!Codmacta, "T") & " de ventas de Almazara"
                        Else
                            If cadTabla = "rbodfact2" Then
                                Sql = DBLet(Rs!Codmacta, "T") & " de ventas de Bodega"
                            Else
                                If cadTabla = "rrecibpozos" Then
                                    Select Case Tipo
                                        Case 1
                                            Sql = DBLet(Rs!Codmacta, "T") & " de ventas consumo de Pozos"
                                        Case 2
                                            Sql = DBLet(Rs!Codmacta, "T") & " de ventas cuotas de Pozos"
                                        Case 3
                                            Sql = DBLet(Rs!Codmacta, "T") & " de ventas talla de Pozos"
                                        Case 4
                                            Sql = DBLet(Rs!Codmacta, "T") & " de ventas mantenimiento de Pozos"
                                        Case 5
                                            Sql = DBLet(Rs!Codmacta, "T") & " de vevntas consumo a manta Pozos"
                                    End Select
                                Else
                                    If cadTabla = "fvarcabfact" Then
                                        Sql = DBLet(Rs!Codmacta, "T") & " del concepto de factura varia cliente"
                                    Else
                                        If cadTabla = "fvarcabfactpro" Then
                                            Sql = DBLet(Rs!Codmacta, "T") & " del concepto de factura varia proveedor"
                                        Else
                                            If cadTabla = "rtelmovil" Then
                                                Sql = DBLet(Rs!Codmacta, "T") & " de ventas de Telefonia"
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            ElseIf Opcion = 4 Then
                Sql = vParamAplic.CtaTerReten
            ElseIf Opcion = 7 Then
                Sql = DBLet(Rs!cuenta, "T")
            ElseIf Opcion = 8 Then
                Sql = DBLet(Rs!Codmacta, "T") & " de la variedad " & Format(Rs!codvarie, "0000")
            ElseIf Opcion = 9 Then
                Sql = DBLet(Rs!cuenta, "T") & " de ventas de almazara "
            ElseIf Opcion = 11 Then
                Sql = DBLet(Rs!cuenta, "T") & " de gastos de almazara "
            ElseIf Opcion = 12 Then
                Sql = DBLet(Rs!cuenta, "T") & " de gasto de concepto a pie de factura "
            ElseIf Opcion = 13 Then
                Sql = DBLet(Rs!Codmacta, "T") & " del concepto de gasto "
            ElseIf Opcion = 14 Then
                Sql = DBLet(Rs!Codmacta, "T") & " del Socio Asociado " & Format(Rs!Codsocio, "000000")
            ElseIf Opcion = 15 Then
                Sql = DBLet(Rs!cuenta, "T") & " de Aportacion del Socio "
            End If
        End If

        If B And (Opcion = 2 Or Opcion = 7) Then
            If cadTabla = "advfacturas" Then
                'Comprobar que ademas de existir la cuenta de ventas exista tambien
                'la cuenta ABONO ventas (sfamia.aboventa)
                '---------------------------------------------
                Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!ctaabono, "T")
    '            RSconta.MoveFirst
    '            RSconta.Find (SQL), , adSearchForward
    '            If RSconta.EOF Then
                If Not (RegistrosAListar(Sql, cConta) > 0) Then
                    B = False 'no encontrado
                    If Opcion = 2 Then
                        Sql = DBLet(Rs!ctaabono, "T") & " de la familia " & Format(Rs!codfamia, "0000")
                    ElseIf Opcion = 7 Then
                        Sql = DBLet(Rs!ctaabono, "T")
                    End If
                End If
            End If
        End If

        Rs.MoveNext
    Wend

    If Not B Then
        If Not (Opcion = 7 Or Opcion = 9 Or Opcion = 10 Or Opcion = 11 Or Opcion = 12) Then
            Sql = "No existe la cta contable " & Sql
            If Opcion = 4 Then Sql = Sql & " de retención."
        Else
            Sql = "La cuenta " & Sql & " no es del nivel correcto. "
        End If
        Sql = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & Sql

        MsgBox Sql, vbExclamation
        ComprobarCtaContable_new = False
    Else
        ComprobarCtaContable_new = True
    End If
    
    Exit Function

ECompCta:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
    End If
End Function


Public Function ComprobarTiposIVA(cadTabla As String) As Boolean
'Comprobar que todos los Tipos de IVA de las distintas facturas (scafac.codigiva1, codigiv2,codigiv3)
'que vamos a contabilizar existan en la contabilidad
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim B As Boolean
Dim I As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVA = False
    
    Sql = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        For I = 1 To 3
            If cadTabla = "advfacturas" Then
                Sql = "SELECT DISTINCT advfacturas.codiiva" & I
                Sql = Sql & " FROM advfacturas "
                Sql = Sql & " INNER JOIN tmpFactu ON advfacturas.codtipom=tmpFactu.codtipom AND advfacturas.numfactu=tmpFactu.numfactu AND advfacturas.fecfactu=tmpFactu.fecfactu "
                Sql = Sql & " WHERE not isnull(codiiva" & I & ")"
'                SQL = SQL & " WHERE " & " codigiv" & i & " <> 0 "
            Else
                If cadTabla = "rbodfacturas" Then
                    Sql = "SELECT DISTINCT rbodfacturas.codiiva" & I
                    Sql = Sql & " FROM rbodfacturas "
                    Sql = Sql & " INNER JOIN tmpFactu ON rbodfacturas.codtipom=tmpFactu.codtipom AND rbodfacturas.numfactu=tmpFactu.numfactu AND rbodfacturas.fecfactu=tmpFactu.fecfactu "
                    Sql = Sql & " WHERE not isnull(codiiva" & I & ")"
                Else
                    If cadTabla = "scafpc" Then
                        Sql = "SELECT DISTINCT scafpc.tipoiva" & I
                        Sql = Sql & " FROM " & cadTabla
                        Sql = Sql & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
                        Sql = Sql & " WHERE not isnull(tipoiva" & I & ")"
        '                SQL = SQL & " WHERE " & " tipoiva" & i & " <> 0 "
                    Else
                        If cadTabla = "rrecibpozos" Then
                            Sql = "SELECT DISTINCT tipoiva"
                            Sql = Sql & " FROM " & cadTabla
                            Sql = Sql & " INNER JOIN tmpFactu ON rrecibpozos.codtipom=tmpFactu.codtipom AND rrecibpozos.numfactu=tmpFactu.numfactu AND rrecibpozos.fecfactu=tmpFactu.fecfactu "
                            Sql = Sql & " WHERE not isnull(tipoiva)"
            '                SQL = SQL & " WHERE " & " tipoiva" & i & " <> 0 "
                        Else
                            Sql = "SELECT DISTINCT rcafter.tipoiva" & I
                            Sql = Sql & " FROM " & cadTabla
                            Sql = Sql & " INNER JOIN tmpFactu ON rcafter.codsocio=tmpFactu.codsocio AND rcafter.numfactu=tmpFactu.numfactu AND rcafter.fecfactu=tmpFactu.fecfactu "
                            Sql = Sql & " WHERE not isnull(tipoiva" & I & ")"
            '                SQL = SQL & " WHERE " & " tipoiva" & i & " <> 0 "
                    
                        End If
                    End If
               End If
            End If
'            SQL = SQL & " WHERE " & cadWHERE & " AND codigiv" & i & " <> 0 "

            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            B = True
            While Not Rs.EOF And B
                Sql = "codigiva= " & DBSet(Rs.Fields(0), "N")
                RSconta.MoveFirst
                RSconta.Find (Sql), , adSearchForward
                If RSconta.EOF Then
                    B = False 'no encontrado
                    Sql = "Tipo de IVA: " & Rs.Fields(0)
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        
            If Not B Then
                Sql = "No existe el " & Sql
                Sql = "Comprobando Tipos de IVA en contabilidad..." & vbCrLf & vbCrLf & Sql
            
                MsgBox Sql, vbExclamation
                ComprobarTiposIVA = False
                Exit For
            Else
                ComprobarTiposIVA = True
            End If
        Next I
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompIVA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Tipo de IVA.", Err.Description
    End If
End Function


Public Function ComprobarIVA(cadTabla As String, Optional CodIva As String) As Boolean
'Comprobar que todos los Tipos de IVA de las distintas facturas (scafac.codigiva1, codigiv2,codigiv3)
'que vamos a contabilizar existan en la contabilidad
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim B As Boolean
Dim I As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarIVA = False
    
    Sql = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
            Select Case cadTabla
                Case "rfactsoc"
                    Sql = "SELECT DISTINCT rfactsoc.tipoiva"
                    Sql = Sql & " FROM rfactsoc "
                    Sql = Sql & " INNER JOIN tmpFactu ON rfactsoc.codtipom=tmpFactu.codtipom AND rfactsoc.numfactu=tmpFactu.numfactu AND rfactsoc.fecfactu=tmpFactu.fecfactu "
                    Sql = Sql & " WHERE not isnull(tipoiva)"
                Case "rcabfactalmz"
                    Sql = "SELECT DISTINCT rcabfactalmz.tipoiva"
                    Sql = Sql & " FROM rcabfactalmz "
                    Sql = Sql & " INNER JOIN tmpFactu ON rcabfactalmz.tipofichero=tmpFactu.tipofichero AND rcabfactalmz.numfactu=tmpFactu.numfactu AND rcabfactalmz.fecfactu=tmpFactu.fecfactu "
                    Sql = Sql & " and rcabfactalmz.codsocio = tmpFactu.codsocio "
                    Sql = Sql & " WHERE not isnull(tipoiva)"
                Case "rtelmovil"
                    Sql = "SELECT DISTINCT " & CodIva
                    Sql = Sql & " FROM rtelmovil  "
                Case "rfacttra"
                    Sql = "SELECT DISTINCT rfacttra.tipoiva"
                    Sql = Sql & " FROM rfacttra "
                    Sql = Sql & " INNER JOIN tmpFactu ON rfacttra.codtipom=tmpFactu.codtipom AND rfacttra.numfactu=tmpFactu.numfactu AND rfacttra.fecfactu=tmpFactu.fecfactu "
                    Sql = Sql & " WHERE not isnull(tipoiva)"
                
            End Select

            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            B = True
            While Not Rs.EOF And B
                Sql = "codigiva= " & DBSet(Rs.Fields(0), "N")
                RSconta.MoveFirst
                RSconta.Find (Sql), , adSearchForward
                If RSconta.EOF Then
                    B = False 'no encontrado
                    Sql = "Tipo de IVA: " & Rs.Fields(0)
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        
            If Not B Then
                Sql = "No existe el " & Sql
                Sql = "Comprobando Tipos de IVA en contabilidad..." & vbCrLf & vbCrLf & Sql
            
                MsgBox Sql, vbExclamation
                ComprobarIVA = False
            Else
                ComprobarIVA = True
            End If
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompIVA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar IVA.", Err.Description
    End If
End Function





Public Function ComprobarCCoste(cadCC As String) As Boolean
Dim Sql As String

    On Error GoTo ECCoste

    ComprobarCCoste = False
    Sql = vUsu.Login
    If Sql <> "" Then
        cadCC = DevuelveDesdeBDNew(cAgro, "straba", "codccost", "login", Sql, "T")
        If cadCC <> "" Then
            'comprobar que el Centro de Coste existe en la Contabilidad
            Sql = DevuelveDesdeBDNew(cConta, "cabccost", "codccost", "codccost", cadCC, "T")
            If Sql <> "" Then
                ComprobarCCoste = True
            Else
                Sql = "No existe el CC: " & cadCC
                Sql = "Comprobando Centros de Coste en contabilidad..." & vbCrLf & vbCrLf & Sql
                MsgBox Sql, vbExclamation
            End If
        Else 'el usuario no tiene asignado un centro de coste
            Sql = "El trabajador conectado no tiene asignado un centro de coste."
            Sql = "Comprobando Centros de Coste ..." & vbCrLf & vbCrLf & Sql
            MsgBox Sql, vbExclamation
        End If
    End If
    
ECCoste:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Centros de Coste", Err.Description
    End If
End Function


Public Function ComprobarCCoste_new(cadCC As String, cadTabla As String, Optional Opcion As Byte) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim B As Boolean

    On Error GoTo ECCoste

    ComprobarCCoste_new = False
    Select Case cadTabla
        Case "facturas" ' facturas de venta
            If Opcion = 1 Then
                Sql = "select distinct variedades.codccost from facturas_variedad, albaran_variedad, variedades, tmpFactu where "
                Sql = Sql & " albaran_variedad.codvarie=variedades.codvarie and "
                Sql = Sql & " facturas_variedad.codtipom=tmpFactu.codtipom AND facturas_variedad.numfactu=tmpFactu.numfactu AND facturas_variedad.fecfactu=tmpFactu.fecfactu and  "
                Sql = Sql & " albaran_variedad.numalbar = facturas_variedad.numalbar and "
                Sql = Sql & " albaran_variedad.numlinea = facturas_variedad.numlinealbar "
            Else
                Sql = Sql & " select distinct sfamia.codccost from facturas_envases, sartic, sfamia, tmpFactu where "
                Sql = Sql & " facturas_envases.codtipom=tmpFactu.codtipom AND facturas_envases.numfactu=tmpFactu.numfactu AND facturas_envases.fecfactu=tmpFactu.fecfactu and  "
                Sql = Sql & " facturas_envases.codartic = sartic.codartic and "
                Sql = Sql & " sartic.codfamia = sfamia.codfamia "
            End If
        Case "scafpc" ' facturas de compra
            Sql = Sql & " select distinct sfamia.codccost from slifpc, sartic, sfamia, tmpFactu where "
            Sql = Sql & " slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu and  "
            Sql = Sql & " slifpc.codartic = sartic.codartic and "
            Sql = Sql & " sartic.codfamia = sfamia.codfamia "
        
        Case "rfactsoc" ' facturas de socio
            Sql = "select distinct variedades.codccost from rfactsoc_variedad,  variedades, tmpFactu where "
            Sql = Sql & " rfactsoc_variedad.codvarie=variedades.codvarie and "
            Sql = Sql & " rfactsoc_variedad.codtipom=tmpFactu.codtipom AND rfactsoc_variedad.numfactu=tmpFactu.numfactu AND rfactsoc_variedad.fecfactu=tmpFactu.fecfactu "
        
        Case "rfacttra" ' facturas de transporte
            Sql = "select distinct variedades.codccost from rfacttra_albaran,  variedades, tmpFactu where "
            Sql = Sql & " rfacttra_albaran.codvarie=variedades.codvarie and "
            Sql = Sql & " rfacttra_albaran.codtipom=tmpFactu.codtipom AND rfacttra_albaran.numfactu=tmpFactu.numfactu AND rfacttra_albaran.fecfactu=tmpFactu.fecfactu "
        
        Case "rcafter" ' facturas de terceros
            Sql = "select distinct variedades.codccost from rlifter,  variedades, tmpFactu where "
            Sql = Sql & " rlifter.codvarie=variedades.codvarie and "
            Sql = Sql & " rlifter.codsocio=tmpFactu.codsocio AND rlifter.numfactu=tmpFactu.numfactu AND rlifter.fecfactu=tmpFactu.fecfactu "
    
        Case "advfacturas" ' facturas de venta adv
            Sql = Sql & "select distinct advfamia.codccost from advfacturas_lineas, advartic, advfamia, tmpFactu where "
            Sql = Sql & " advfacturas_lineas.codtipom=tmpFactu.codtipom AND advfacturas_lineas.numfactu=tmpFactu.numfactu AND advfacturas_lineas.fecfactu=tmpFactu.fecfactu and  "
            Sql = Sql & " advfacturas_lineas.codartic = advartic.codartic and "
            Sql = Sql & " advartic.codfamia = advfamia.codfamia "
        
        Case "rrecibpozos" ' recibos de consumo de pozos
            Sql = Sql & "select distinct " & DBSet(vParamAplic.CodCCostPOZ, "T") & " as codccost from rrecibpozos where 1=1 "
        
        Case "rbodfacturas" ' facturas de retirada de bodega / almazara
            Sql = "select distinct variedades.codccost from rbodfacturas_lineas, variedades, tmpFactu where "
            Sql = Sql & " rbodfacturas_lineas.codvarie=variedades.codvarie and rbodfacturas_lineas.codtipom=tmpFactu.codtipom and "
            Sql = Sql & " rbodfacturas_lineas.numfactu=tmpFactu.numfactu AND rbodfacturas_lineas.fecfactu=tmpFactu.fecfactu "
    
        Case "fvarcabfact" ' facturas de tipo clientes varias
            Sql = "select distinct fvarconce.codccost from fvarlinfact, fvarconce, tmpFactu where "
            Sql = Sql & " fvarlinfact.codconce=fvarconce.codconce and fvarlinfact.codtipom=tmpFactu.codtipom and "
            Sql = Sql & " fvarlinfact.numfactu=tmpFactu.numfactu AND fvarlinfact.fecfactu=tmpFactu.fecfactu "
    
        Case "fvarcabfactpro" ' facturas de tipo proveedor varias
            Sql = "select distinct fvarconce.codccost from fvarlinfactpro, fvarconce, tmpFactu where "
            Sql = Sql & " fvarlinfactpro.codconce=fvarconce.codconce and fvarlinfactpro.codtipom=tmpFactu.codtipom and "
            Sql = Sql & " fvarlinfactpro.numfactu=tmpFactu.numfactu AND fvarlinfactpro.fecfactu=tmpFactu.fecfactu "
    End Select
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    B = True

    While Not Rs.EOF And B
        Sql = DevuelveDesdeBDNew(cConta, "cabccost", "codccost", "codccost", DBLet(Rs.Fields(0).Value), "T")
        If Sql = "" Then
            B = False
            Sql2 = "Centro de Coste: " & Rs.Fields(0)
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
    If Not B Then
        Sql = "No existe el " & Sql2
        Sql = "Comprobando Centros de Coste en contabilidad..." & vbCrLf & vbCrLf & Sql
    
        MsgBox Sql, vbExclamation
        ComprobarCCoste_new = False
        Exit Function
    Else
        ComprobarCCoste_new = True
    End If
    
ECCoste:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Centros de Coste", Err.Description
    End If
End Function


Public Function ComprobarFormadePago(cadCC As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim B As Boolean

    On Error GoTo ECCoste

    ComprobarFormadePago = False
    
    Select Case cadCC
        Case "advfacturas"
            Sql = "select distinct advfacturas.codforpa from advfacturas, tmpFactu where "
            Sql = Sql & " advfacturas.codtipom=tmpFactu.codtipom AND advfacturas.numfactu=tmpFactu.numfactu AND advfacturas.fecfactu=tmpFactu.fecfactu  "
        Case "rbodfacturas"
            Sql = "select distinct rbodfacturas.codforpa from rbodfacturas, tmpFactu where "
            Sql = Sql & " rbodfacturas.codtipom=tmpFactu.codtipom AND rbodfacturas.numfactu=tmpFactu.numfactu AND rbodfacturas.fecfactu=tmpFactu.fecfactu  "
    End Select
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    B = True

    While Not Rs.EOF And B
        If vParamAplic.ContabilidadNueva Then
            Sql = DevuelveDesdeBDNew(cConta, "formapago", "codforpa", "codforpa", Rs.Fields(0).Value, "N")
        Else
            Sql = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", Rs.Fields(0).Value, "N")
        End If
        If Sql = "" Then
            B = False
            Sql2 = "Formas de Pago: " & Rs.Fields(0)
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
    If Not B Then
        Sql = "No existe la " & Sql2
        Sql = "Comprobando Formas de Pago en contabilidad..." & vbCrLf & vbCrLf & Sql
    
        MsgBox Sql, vbExclamation
        ComprobarFormadePago = False
        Exit Function
    Else
        ComprobarFormadePago = True
    End If
    
ECCoste:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Formas de Pago", Err.Description
    End If
End Function




Public Function PasarFacturaADV(cadWHERE As String, CodCCost As String, CtaBan As String, FecVen As String, TipoM As String, FecFac As Date, Observac As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariges.scafac --> conta.cabfact
' ariges.slifac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim B As Boolean
Dim cadMen As String
Dim Sql As String
Dim Mc As Contadores
Dim Obs As String


    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    
    'Insertar en la conta Cabecera Factura
    
    If TipoM <> "FIN" Then
        
        B = InsertarCabFactADV(cadWHERE, Observac, cadMen, vContaFra)
        cadMen = "Insertando Cab. Factura: " & cadMen
        
        If B Then
            CCoste = CodCCost
            'Insertar lineas de Factura en la Conta
            If vParamAplic.ContabilidadNueva Then
                B = InsertarLinFactADVContaNueva("advfacturas", cadWHERE, cadMen)
            Else
                B = InsertarLinFactADV("advfacturas", cadWHERE, cadMen)
            End If
            cadMen = "Insertando Lin. Factura: " & cadMen
    
            '++monica:añadida la parte de insertar en tesoreria
            If B Then
                B = InsertarEnTesoreriaNewADV(cadWHERE, CtaBan, FecVen, cadMen)
                cadMen = "Insertando en Tesoreria: " & cadMen
            End If
            
            If B Then
                If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
            End If

        End If
            '++
    Else
        ' No insertamos la factura sino un asiento en el diario
        Set Mc = New Contadores
        
        If Mc.ConseguirContador("0", (CDate(FecFac) <= CDate(vEmpresa.FechaFin)), True) = 0 Then
        
            Obs = "Contabilización Factura Interna de Fecha " & Format(FecFac, "dd/mm/yyyy")
        
            'Insertar en la conta Cabecera Asiento
            B = InsertarCabAsientoDia(vEmpresa.NumDiarioInt, Mc.Contador, CStr(Format(FecFac, "dd/mm/yyyy")), Obs, cadMen)
            cadMen = "Insertando Cab. Asiento: " & cadMen
        Else
            B = False
        End If
    
        If B Then
            Socio = DevuelveValor("select codsocio from advfacturas where " & cadWHERE)
            CtaSocio = DevuelveValor("select codmaccli from rsocios_seccion where codsocio = " & Socio & " and codsecci = " & vParamAplic.SeccionADV)
        
        
            B = InsertarLinAsientoFactInt("advfacturas", cadWHERE, cadMen, CtaSocio, Mc.Contador)
            cadMen = "Insertando Lin. Factura Interna: " & cadMen
        
            Set Mc = Nothing
        End If
    
        '++monica:añadida la parte de insertar en tesoreria
        If B Then
            CCoste = CodCCost
            B = InsertarEnTesoreriaNewADV(cadWHERE, CtaBan, FecVen, cadMen)
            cadMen = "Insertando en Tesoreria: " & cadMen
        End If
    
    End If

    If B Then
        'Poner intconta=1 en ariagro.facturas
        B = ActualizarCabFact("advfacturas", cadWHERE, cadMen)
        cadMen = "Actualizando Factura: " & cadMen
    End If
    
    
    
    
'    If Not b Then
'        Sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
'        Sql = Sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
'        Sql = Sql & " WHERE " & Replace(cadWhere, "facturas", "tmpFactu")
'        Conn.Execute Sql
'    End If
    
EContab:
    
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Contabilizando Factura ADV", Err.Description
    End If
    If B Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaADV = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaADV = False
        
        Sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
        Sql = Sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "advfacturas", "tmpFactu")
        conn.Execute Sql
    End If
End Function

Public Function InsertarCabAsientoDia(Diario As String, Asiento As String, Fecha As String, Obs As String, cadErr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String

    On Error GoTo EInsertar
       
    

    If vParamAplic.ContabilidadNueva Then
        cad = Format(Diario, "00") & ", " & DBSet(Fecha, "F") & "," & Format(Asiento, "000000") & ","
        cad = cad & DBSet(Obs, "T") & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARIAGRO RECOLECCION'"
        
        cad = "(" & cad & ")"
        
        'Insertar en la contabilidad
        Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) "
        Sql = Sql & " VALUES " & cad
    
    Else
        cad = Format(Diario, "00") & ", " & DBSet(Fecha, "F") & "," & Format(Asiento, "000000") & ","
        cad = cad & "''," & ValorNulo & "," & DBSet(Obs, "T")
        cad = "(" & cad & ")"
        
        'Insertar en la contabilidad
        Sql = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) "
        Sql = Sql & " VALUES " & cad
    
    End If
    
    ConnConta.Execute Sql
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabAsientoDia = False
        cadErr = Err.Description
    Else
        InsertarCabAsientoDia = True
    End If
End Function


Private Function InsertarLinAsientoFactInt(cadTabla As String, cadWHERE As String, cadErr As String, CtaSocio As String, Optional Contador As Long) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim numdocum As String
Dim Ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim I As Long
Dim B As Boolean
Dim cad As String
Dim cadMen As String
Dim FeFact As Date

Dim cadCampo As String

    On Error GoTo eInsertarLinAsientoFactInt

    InsertarLinAsientoFactInt = False

    TotalFac = DevuelveValor("select totalfac from advfacturas where " & cadWHERE)
    'utilizamos sfamia.ctaventa o sfamia.aboventa
    If TotalFac >= 0 Then
        cadCampo = "advfamia.ctaventa"
    Else
        cadCampo = "advfamia.aboventa"
    End If
    
    If vEmpresa.TieneAnalitica Then
        Sql = " SELECT usuarios.stipom.letraser,advfacturas_lineas.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, advfamia.codccost "
    Else
        Sql = " SELECT usuarios.stipom.letraser,advfacturas_lineas.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe "
    End If
    
    Sql = Sql & " FROM ((advfacturas_lineas inner join usuarios.stipom on advfacturas_lineas.codtipom=usuarios.stipom.codtipom) "
    Sql = Sql & " inner join advartic on advfacturas_lineas.codartic=advartic.codartic) "
    Sql = Sql & " inner join advfamia on advartic.codfamia=advfamia.codfamia "
    Sql = Sql & " WHERE " & Replace(cadWHERE, "advfacturas", "advfacturas_lineas")
    If vEmpresa.TieneAnalitica Then
        Sql = Sql & " GROUP BY 5,7 " '& cadCampo, codccost
    Else
        Sql = Sql & " GROUP BY 5 " '& cadCampo
    End If

    
    Set Rs = New ADODB.Recordset
    
    Rs.Open Sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
            
    I = 0
    ImporteD = 0
    ImporteH = 0
    
    numdocum = Format(Rs!numfactu, "0000000")
    Ampliacion = Rs.Fields(0).Value & "-" & Format(Rs!numfactu, "0000000")
    ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoInt, "N")) & " " & Ampliacion
    ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoInt, "N")) & " " & Ampliacion
    
    If Not Rs.EOF Then Rs.MoveFirst
    
    B = True
    
    
    
    While Not Rs.EOF And B
        I = I + 1
        
        FeFact = Rs!fecfactu
        
        cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Contador, "N") & ","
        cad = cad & DBSet(I, "N") & "," & DBSet(Rs!cuenta, "T") & "," & DBSet(numdocum, "T") & ","
        
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If Rs.Fields(5).Value < 0 Then
            ' importe al debe en positivo
            cad = cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Rs.Fields(5).Value * (-1), "N") & ","
            cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
        
            ImporteD = ImporteD + (CCur(Rs.Fields(5).Value) * (-1))
        Else
            ' importe al haber en positivo, cambiamos el signo
            cad = cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
            cad = cad & DBSet((Rs.Fields(5).Value), "N") & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
        
            ImporteH = ImporteH + CCur(Rs.Fields(5).Value)
        End If
        
        cad = "(" & cad & ")"
        
        B = InsertarLinAsientoDia(cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & I

        Rs.MoveNext
    Wend
    
    If B And I > 0 Then
        I = I + 1
                
        ' el Total es sobre la cuenta del cliente
        cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FeFact, "F") & "," & DBSet(Contador, "N") & ","
        cad = cad & DBSet(I, "N") & ","
        cad = cad & DBSet(CtaSocio, "T") & "," & DBSet(numdocum, "T") & ","
            
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If ImporteD - ImporteH > 0 Then
            ' importe al debe en positivo
            cad = cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliaciond, "T") & "," & ValorNulo & ","
            cad = cad & DBSet(ImporteD - ImporteH, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
        Else
            ' importe al haber en positivo, cambiamos el signo
            cad = cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliacionh, "T") & "," & DBSet(((ImporteD - ImporteH) * -1), "N") & ","
            cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
        End If
        
        cad = "(" & cad & ")"
        
        B = InsertarLinAsientoDia(cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & I
        
    End If
        
    Set Rs = Nothing
    InsertarLinAsientoFactInt = B
    Exit Function
    
eInsertarLinAsientoFactInt:
    cadErr = "Insertar Linea Asiento Factura Interna: " & Err.Description
    cadErr = cadErr & cadMen
    InsertarLinAsientoFactInt = False
End Function


Private Function InsertarLinAsientoFactIntPOZ(cadTabla As String, cadWHERE As String, cadErr As String, CtaSocio As String, Optional Contador As Long) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim numdocum As String
Dim Ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim I As Long
Dim B As Boolean
Dim cad As String
Dim cadMen As String
Dim FeFact As Date
Dim ImpConsumo As Currency
Dim ImpCuota As Currency
Dim totimp As Currency

Dim cadCampo As String

    On Error GoTo eInsertarLinAsientoFactInt

    InsertarLinAsientoFactIntPOZ = False

'============

    If vParamAplic.Cooperativa = 7 Then ' si la cooperativa es quatretonda
        If vEmpresa.TieneAnalitica Then
            Sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu,sum(round(precio1*consumo1,2)) as importeconsumo,sum(round(precio2*consumo2,2) + impcuota) as importecuota, " & DBSet(vParamAplic.CodCCostPOZ, "T") & " as codccost "
        Else
            Sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu,sum(round(precio1*consumo1,2)) as importeconsumo,sum(round(precio2*consumo2,2) + impcuota) as importecuota "
        End If
        
        Sql = Sql & " FROM rrecibpozos inner join usuarios.stipom on rrecibpozos.codtipom=usuarios.stipom.codtipom "
        Sql = Sql & " WHERE " & cadWHERE
        
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & " GROUP BY 1,2,3,4,7 " '& cadCampo, codccost
        Else
            Sql = Sql & " GROUP BY 1,2,3,4" '& cadCampo
        End If
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        
        numdocum = Format(Rs!numfactu, "0000000")
        Ampliacion = Rs.Fields(0).Value & "-" & Format(Rs!numfactu, "0000000")
        ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoInt, "N")) & " " & Ampliacion
        ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoInt, "N")) & " " & Ampliacion
        
        
        
        cad = ""
        I = 1
        totimp = 0
        If Not Rs.EOF Then
            
            ImpConsumo = DBLet(Rs!Importeconsumo, "N")
            ImpCuota = DBLet(Rs!importecuota, "N")
            totimp = totimp + ImpConsumo + ImpCuota
    
            B = True
            If ImpConsumo <> 0 Then
                I = I + 1
            
                FeFact = Rs!fecfactu
                
                cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Contador, "N") & ","
                cad = cad & DBSet(I, "N") & "," & DBSet(vParamAplic.CtaVentasConsPOZ, "T") & "," & DBSet(numdocum, "T") & ","
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If ImpConsumo < 0 Then
                    ' importe al debe en positivo
                    cad = cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(ImpConsumo * (-1), "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
                
                    ImporteD = ImporteD + (CCur(ImpConsumo) * (-1))
                Else
                    ' importe al haber en positivo, cambiamos el signo
                    cad = cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet(ImpConsumo, "N") & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + CCur(ImpConsumo)
                End If
                
                
                cad = "(" & cad & ")"
                
                B = InsertarLinAsientoDia(cad, cadMen)
                cadMen = "Insertando Lin. Asiento: " & I

            End If
            
            
            If B And ImpCuota <> 0 Then
                I = I + 1
            
                FeFact = Rs!fecfactu
                
                cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Contador, "N") & ","
                cad = cad & DBSet(I, "N") & "," & DBSet(vParamAplic.CtaVentasCuoPOZ, "T") & "," & DBSet(numdocum, "T") & ","
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If ImpCuota < 0 Then
                    ' importe al debe en positivo
                    cad = cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(ImpCuota * (-1), "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
                
                    ImporteD = ImporteD + (CCur(ImpCuota) * (-1))
                Else
                    ' importe al haber en positivo, cambiamos el signo
                    cad = cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet(ImpCuota, "N") & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + CCur(ImpCuota)
                End If
                
                cad = "(" & cad & ")"
                
                B = InsertarLinAsientoDia(cad, cadMen)
                cadMen = "Insertando Lin. Asiento: " & I
                
            End If
            
        
            If B And I > 1 Then
                I = 1
                        
                ' el Total es sobre la cuenta del cliente
                cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FeFact, "F") & "," & DBSet(Contador, "N") & ","
                cad = cad & DBSet(I, "N") & ","
                cad = cad & DBSet(CtaSocio, "T") & "," & DBSet(numdocum, "T") & ","
                    
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If ImporteD - ImporteH > 0 Then
                    ' importe al debe en positivo
                    cad = cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliaciond, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet(ImporteD - ImporteH, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                Else
                    ' importe al haber en positivo, cambiamos el signo
                    cad = cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliacionh, "T") & "," & DBSet(((ImporteD - ImporteH) * -1), "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                End If
                
                cad = "(" & cad & ")"
                
                B = InsertarLinAsientoDia(cad, cadMen)
                cadMen = "Insertando Lin. Asiento: " & I
            End If
        End If
        
    End If

    Set Rs = Nothing
    InsertarLinAsientoFactIntPOZ = B
    Exit Function
    
eInsertarLinAsientoFactInt:
    cadErr = "Insertar Linea Asiento Factura Interna Pozos: " & Err.Description
    cadErr = cadErr & cadMen
    InsertarLinAsientoFactIntPOZ = False
End Function






Public Function InsertarLinAsientoDia(cad As String, cadErr As String) As Boolean
' el Tipo me indica desde donde viene la llamada
' tipo = 0 srecau.codmacta
' tipo = 1 scaalb.codmacta

Dim Rs As ADODB.Recordset
Dim Aux As String
Dim Sql As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

    If vParamAplic.ContabilidadNueva Then
        Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
        Sql = Sql & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) "
        Sql = Sql & " VALUES " & cad
    
    Else
        
        Sql = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
        Sql = Sql & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) "
        Sql = Sql & " VALUES " & cad
        
    End If
    ConnConta.Execute Sql

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinAsientoDia = False
        cadErr = Err.Description
    Else
        InsertarLinAsientoDia = True
    End If
End Function



Public Function PasarFacturaBOD(cadWHERE As String, CodCCost As String, CtaBan As String, FecVen As String, Tipo As Byte, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariagro.rbodfacturas --> conta.cabfact
' ariagro.rbodfacturas_variedad --> conta.linfact
'Actualizar la tabla ariagro.rbodfacturas.inconta=1 para indicar que ya esta contabilizada
'Tipo : 0 = facturas de retirada de almazara
'       1 = facturas de retirada de bodega

Dim B As Boolean
Dim cadMen As String
Dim Sql As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    'Insertar en la conta Cabecera Factura
    B = InsertarCabFactBOD(cadWHERE, cadMen, Tipo, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    
    If B Then
        CCoste = CodCCost
        'Insertar lineas de Factura en la Conta
        If vParamAplic.ContabilidadNueva Then
            Select Case Tipo
                Case 0
                    B = InsertarLinFactBODContaNueva("rbodfact1", cadWHERE, cadMen)
                Case 1
                    B = InsertarLinFactBODContaNueva("rbodfact2", cadWHERE, cadMen)
            End Select
        
        Else
            Select Case Tipo
                Case 0
                    B = InsertarLinFactBOD("rbodfact1", cadWHERE, cadMen)
                Case 1
                    B = InsertarLinFactBOD("rbodfact2", cadWHERE, cadMen)
            End Select
        End If
        'b = InsertarLinFactBOD("rbodfacturas", cadWHERE, cadMen)
        cadMen = "Insertando Lin. Factura: " & cadMen

        '++monica:añadida la parte de insertar en tesoreria
        If B Then
            B = InsertarEnTesoreriaNewBOD(cadWHERE, CtaBan, FecVen, cadMen, Tipo)
            cadMen = "Insertando en Tesoreria: " & cadMen
        End If
        '++
        
        If B Then
            If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
        End If


        If B Then
            'Poner intconta=1 en ariagro.facturas
            B = ActualizarCabFact("rbodfacturas", cadWHERE, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
'    If Not b Then
'        Sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
'        Sql = Sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
'        Sql = Sql & " WHERE " & Replace(cadWhere, "facturas", "tmpFactu")
'        Conn.Execute Sql
'    End If
    
EContab:
    
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Contabilizando Factura Retirada", Err.Description
    End If
    If B Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaBOD = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaBOD = False
        
        Sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
        Sql = Sql & " Select tmpfactu.*," & DBSet(cadMen, "T") & " as error From tmpFactu "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "rbodfacturas", "tmpFactu")
        conn.Execute Sql
    End If
End Function


Public Function PasarFacturaTel(cadWHERE As String, CodCCost As String, CtaVtas As String, CodIva As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariagro.rbodfacturas --> conta.cabfact
' ariagro.rbodfacturas_variedad --> conta.linfact
'Actualizar la tabla ariagro.rbodfacturas.inconta=1 para indicar que ya esta contabilizada
'Tipo : 0 = facturas de retirada de almazara
'       1 = facturas de retirada de bodega

Dim B As Boolean
Dim cadMen As String
Dim Sql As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    CodiIVA = CodIva
    
    'Insertar en la conta Cabecera Factura
    B = InsertarCabFactTEL(cadWHERE, cadMen, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If B Then
        CCoste = CodCCost
        'Insertar lineas de Factura en la Conta
        If vParamAplic.ContabilidadNueva Then
            B = InsertarLinFactTELContaNueva(CtaVtas, cadWHERE, cadMen)
        Else
            B = InsertarLinFactTEL(CtaVtas, cadWHERE, cadMen)
        End If
        
        cadMen = "Insertando Lin. Factura: " & cadMen

'--Monica: quitado de momento
'        '++monica:añadida la parte de insertar en tesoreria
'        If b Then
'            b = InsertarEnTesoreriaNewBOD(cadWHERE, CtaBan, FecVen, cadMen, Tipo)
'            cadMen = "Insertando en Tesoreria: " & cadMen
'        End If
'
        '++
        If B Then
            If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
        End If
        
        If B Then
            'Poner intconta=1 en ariagro.facturas
            B = ActualizarCabFact("rtelmovil", cadWHERE, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
    
EContab:
    
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Contabilizando Factura Telefonia", Err.Description
    End If
    If B Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaTel = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaTel = False
        
        Sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
        Sql = Sql & " Select tmpfactu.*," & DBSet(cadMen, "T") & " as error From tmpFactu "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "rtelmovil", "tmpFactu")
        conn.Execute Sql
    End If
End Function





Private Function InsertarCabFactADV(cadWHERE As String, Observac As String, cadErr As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim IvaImp As Currency

Dim Sql2 As String
Dim CadenaInsertFaclin2 As String




    On Error GoTo EInsertar
    
    Sql = "SELECT stipom.letraser,numfactu,fecfactu, rsocios_seccion.codmaccli as codmacta,year(fecfactu) as anofaccl,"
    Sql = Sql & "baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    Sql = Sql & "totalfac,codiiva1,codiiva2,codiiva3, porcrec1, porcrec2, porcrec3, imporec1, imporec2, imporec3, "
    Sql = Sql & "advfacturas.codforpa, "
    Sql = Sql & "advfacturas.nomsocio, advfacturas.dirsocio,advfacturas.pobsocio,advfacturas.codpostal,advfacturas.prosocio,advfacturas.nifsocio"
    Sql = Sql & " FROM ((" & "advfacturas inner join " & "usuarios.stipom on advfacturas.codtipom=usuarios.stipom.codtipom) "
    Sql = Sql & "INNER JOIN rsocios ON advfacturas.codsocio=rsocios.codsocio) "
    Sql = Sql & "INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionADV, "N")
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
        vContaFra.NumeroFactura = DBLet(Rs!numfactu)
        vContaFra.Serie = DBLet(Rs!letraser)
        vContaFra.Anofac = DBLet(Rs!anofaccl)
        
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = 0
        DtoGnral = 0
        BaseImp = Rs!baseimp1 + CCur(DBLet(Rs!baseimp2, "N")) + CCur(DBLet(Rs!baseimp3, "N"))
        IvaImp = DBLet(Rs!impoiva1, "N") + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
        
        '---- Laura 10/10/2006:  añadir el totalfac para utilizarlo en insertar lineas
        TotalFac = Rs!TotalFac
        '----
        
        Sql = ""
        Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!fecfactu) & ","
        '[Monica]02/05/2012: añadido campo observaciones del frame, antes valor nulo
        Sql = Sql & DBSet(Observac, "T") & "," '& ValorNulo & ","
        
        
        If vParamAplic.ContabilidadNueva Then
            ' para el caso de las rectificativas
            Dim vTipM As String
            vTipM = DevuelveValor("select codtipom from stipom where letraser = " & DBSet(Rs!letraser, "T"))
            If vTipM = "FAR" Then
                Sql = Sql & "'D',"
            Else
                Sql = Sql & "'0',"
            End If
            
            Sql = Sql & "0," & DBSet(Rs!Codforpa, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(IvaImp, "N") & ","
            Sql = Sql & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!fecfactu, "F") & ","
            Sql = Sql & DBSet(Rs!nomsocio, "T") & "," & DBSet(Rs!dirsocio, "T") & "," & DBSet(Rs!pobsocio, "T") & "," & DBSet(Rs!codpostal, "T") & ","
            Sql = Sql & DBSet(Rs!prosocio, "T") & "," & DBSet(Rs!nifSocio, "T") & ",'ES',1"
        Else
            Sql = Sql & DBSet(Rs!baseimp1, "N") & "," & DBSet(Rs!baseimp2, "N", "S") & "," & DBSet(Rs!baseimp3, "N", "S") & "," & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", "S") & "," & DBSet(Rs!porciva3, "N", "S") & ","
            Sql = Sql & DBSet(Rs!porcrec1, "N", "S") & "," & DBSet(Rs!porcrec2, "N", "S") & "," & DBSet(Rs!porcrec3, "N", "S") & "," & DBSet(Rs!impoiva1, "N", "N") & "," & DBSet(Rs!impoiva2, "N", "S") & "," & DBSet(Rs!impoiva3, "N", "S") & ","
            Sql = Sql & DBSet(Rs!imporec1, "N", "S") & "," & DBSet(Rs!imporec2, "N", "S") & "," & DBSet(Rs!imporec3, "N", "S") & ","
            Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!codiiva1, "N") & "," & DBSet(Rs!codiiva2, "N", "S") & "," & DBSet(Rs!codiiva3, "N", "S") & ",0,"
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!fecfactu, "F")
        End If
        cad = cad & "(" & Sql & ")"
'        RS.MoveNext
    End If
    
    
    If vParamAplic.ContabilidadNueva Then
        Sql = "INSERT INTO factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
        Sql = Sql & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
        Sql = Sql & "codpais,codagente)"
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
'***
        CadenaInsertFaclin2 = ""
            
        
        'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
        'IVA 1, siempre existe
        Sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
        Sql2 = Sql2 & "1," & DBSet(Rs!baseimp1, "N") & "," & Rs!codiiva1 & "," & DBSet(Rs!porciva1, "N") & ","
        Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & ValorNulo
        CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
        
        'para las lineas
        vTipoIva(0) = Rs!codiiva1
        vPorcIva(0) = Rs!porciva1
        vPorcRec(0) = 0
        vImpIva(0) = Rs!impoiva1
        vImpRec(0) = 0
        vBaseIva(0) = Rs!baseimp1
        
        vTipoIva(1) = 0: vTipoIva(2) = 0
        
        If Not IsNull(Rs!porciva2) Then
            Sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
            Sql2 = Sql2 & "2," & DBSet(Rs!baseimp2, "N") & "," & Rs!codiiva2 & "," & DBSet(Rs!porciva2, "N") & ","
            Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva2, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
            vTipoIva(1) = Rs!codiiva2
            vPorcIva(1) = Rs!porciva2
            vPorcRec(1) = 0
            vImpIva(1) = Rs!impoiva2
            vImpRec(1) = 0
            vBaseIva(1) = Rs!baseimp2
        End If
        If Not IsNull(Rs!porciva3) Then
            Sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
            Sql2 = Sql2 & "3," & DBSet(Rs!baseimp3, "N") & "," & Rs!codiiva3 & "," & DBSet(Rs!porciva3, "N") & ","
            Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva3, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
            vTipoIva(2) = Rs!codiiva3
            vPorcIva(2) = Rs!porciva3
            vPorcRec(2) = 0
            vImpIva(2) = Rs!impoiva3
            vImpRec(2) = 0
            vBaseIva(2) = Rs!baseimp3
        End If

        Sql = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
        Sql = Sql & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
        ConnConta.Execute Sql
    Else
        'Insertar en la contabilidad
        Sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
        Sql = Sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
        Sql = Sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If
    
    Rs.Close
    Set Rs = Nothing
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactADV = False
        cadErr = Err.Description
    Else
        InsertarCabFactADV = True
    End If
End Function

Private Function InsertarCabFactBOD(cadWHERE As String, cadErr As String, Tipo As Byte, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Seccion As Integer
Dim IvaImp As Currency

Dim Sql2 As String
Dim CadenaInsertFaclin2 As String




    On Error GoTo EInsertar
    
    Select Case Tipo
        Case 0
            Seccion = vParamAplic.SeccionAlmaz
        Case 1
            Seccion = vParamAplic.SeccionBodega
    End Select
    
    
    Sql = "SELECT stipom.letraser,numfactu,fecfactu, rsocios_seccion.codmaccli as codmacta,year(fecfactu) as anofaccl,"
    Sql = Sql & "baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    Sql = Sql & "totalfac,codiiva1,codiiva2,codiiva3, porcrec1, porcrec2, porcrec3, imporec1, imporec2, imporec3, rbodfacturas.codforpa,  "
    Sql = Sql & "rbodfacturas.nomsocio, rbodfacturas.dirsocio,rbodfacturas.pobsocio,rbodfacturas.codpostal,rbodfacturas.prosocio,rbodfacturas.nifsocio"
    Sql = Sql & " FROM ((" & "rbodfacturas inner join " & "usuarios.stipom on rbodfacturas.codtipom=usuarios.stipom.codtipom) "
    Sql = Sql & "INNER JOIN rsocios ON rbodfacturas.codsocio=rsocios.codsocio) "
    Sql = Sql & "INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N")
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
        vContaFra.NumeroFactura = DBLet(Rs!numfactu)
        vContaFra.Serie = DBLet(Rs!letraser)
        vContaFra.Anofac = DBLet(Rs!anofaccl)
        
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = 0
        DtoGnral = 0
        BaseImp = Rs!baseimp1 + CCur(DBLet(Rs!baseimp2, "N")) + CCur(DBLet(Rs!baseimp3, "N"))
        IvaImp = DBLet(Rs!impoiva1, "N") + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
        
        '---- Laura 10/10/2006:  añadir el totalfac para utilizarlo en insertar lineas
        TotalFac = Rs!TotalFac
        '----
        
        Sql = ""
        Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!fecfactu) & "," & ValorNulo & ","
        
        If vParamAplic.ContabilidadNueva Then
            
            ' para el caso de las rectificativas
            Dim vTipM As String
            vTipM = DevuelveValor("select codtipom from stipom where letraser = " & DBSet(Rs!letraser, "T"))
            If vTipM = "FAR" Then
                Sql = Sql & "'D',"
            Else
                Sql = Sql & "'0',"
            End If
            
            Sql = Sql & "0," & DBSet(Rs!Codforpa, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(IvaImp, "N") & ","
            Sql = Sql & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(Rs!fecfactu, "F") & ","
            Sql = Sql & DBSet(Rs!nomsocio, "T") & "," & DBSet(Rs!dirsocio, "T") & "," & DBSet(Rs!pobsocio, "T") & "," & DBSet(Rs!codpostal, "T") & ","
            Sql = Sql & DBSet(Rs!prosocio, "T") & "," & DBSet(Rs!nifSocio, "T") & ",'ES',1"
            Sql = "INSERT INTO factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
            Sql = Sql & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
            Sql = Sql & "codpais,codagente)"
            Sql = Sql & " VALUES " & cad
            ConnConta.Execute Sql
    '***
            CadenaInsertFaclin2 = ""
                
            
            'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
            'IVA 1, siempre existe
            Sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
            Sql2 = Sql2 & "1," & DBSet(Rs!baseimp1, "N") & "," & Rs!codiiva1 & "," & DBSet(Rs!porciva1, "N") & ","
            Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
            
            'para las lineas
            vTipoIva(0) = Rs!TipoIVA1
            vPorcIva(0) = Rs!porciva1
            vPorcRec(0) = 0
            vImpIva(0) = Rs!impoiva1
            vImpRec(0) = 0
            vBaseIva(0) = Rs!baseimp1
            
            vTipoIva(1) = 0: vTipoIva(2) = 0
            
            If Not IsNull(Rs!porciva2) Then
                Sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
                Sql2 = Sql2 & "2," & DBSet(Rs!baseimp2, "N") & "," & Rs!codiiva2 & "," & DBSet(Rs!porciva2, "N") & ","
                Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva2, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
                vTipoIva(1) = Rs!TipoIVA2
                vPorcIva(1) = Rs!porciva2
                vPorcRec(1) = 0
                vImpIva(1) = Rs!impoiva2
                vImpRec(1) = 0
                vBaseIva(1) = Rs!baseimp2
            End If
            If Not IsNull(Rs!porciva3) Then
                Sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
                Sql2 = Sql2 & "3," & DBSet(Rs!baseimp3, "N") & "," & Rs!codiiva3 & "," & DBSet(Rs!porciva3, "N") & ","
                Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva3, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
                vTipoIva(2) = Rs!TipoIVA3
                vPorcIva(2) = Rs!porciva3
                vPorcRec(2) = 0
                vImpIva(2) = Rs!impoiva3
                vImpRec(2) = 0
                vBaseIva(2) = Rs!baseimp3
            End If
    
            Sql = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
            Sql = Sql & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
            ConnConta.Execute Sql
                
        Else
            Sql = Sql & DBSet(Rs!baseimp1, "N") & "," & DBSet(Rs!baseimp2, "N", "S") & "," & DBSet(Rs!baseimp3, "N", "S") & "," & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", "S") & "," & DBSet(Rs!porciva3, "N", "S") & ","
            Sql = Sql & DBSet(Rs!porcrec1, "N", "S") & "," & DBSet(Rs!porcrec2, "N", "S") & "," & DBSet(Rs!porcrec3, "N", "S") & "," & DBSet(Rs!impoiva1, "N", "N") & "," & DBSet(Rs!impoiva2, "N", "S") & "," & DBSet(Rs!impoiva3, "N", "S") & ","
            Sql = Sql & DBSet(Rs!imporec1, "N", "S") & "," & DBSet(Rs!imporec2, "N", "S") & "," & DBSet(Rs!imporec3, "N", "S") & ","
            Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!codiiva1, "N") & "," & DBSet(Rs!codiiva2, "N", "S") & "," & DBSet(Rs!codiiva3, "N", "S") & ",0,"
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
            cad = cad & "(" & Sql & ")"
        
            'Insertar en la contabilidad
            Sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,fecliqcl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
            Sql = Sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
            Sql = Sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien) "
            Sql = Sql & " VALUES " & cad
            ConnConta.Execute Sql
        
        End If
'        RS.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactBOD = False
        cadErr = Err.Description
    Else
        InsertarCabFactBOD = True
    End If
End Function



Private Function InsertarCabFactTEL(cadWHERE As String, cadErr As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Seccion As Integer
Dim PorcIva As String
Dim IvaImp As Currency

Dim Sql2 As String
Dim CadenaInsertFaclin2 As String



    On Error GoTo EInsertar
    
    Seccion = vParamAplic.Seccionhorto
    
    Sql = "SELECT numserie,numfactu,fecfactu, rsocios_seccion.codmaccli as codmacta,year(fecfactu) as anofaccl,"
    Sql = Sql & "baseimpo,cuotaiva,totalfac,"
    Sql = Sql & "rsocios.dirsocio,rsocios.pobsocio,rsocios.codpostal,rsocios.prosocio,rsocios.nifsocio "
    Sql = Sql & " FROM (rtelmovil "
    Sql = Sql & "INNER JOIN rsocios ON rtelmovil.codsocio=rsocios.codsocio) "
    Sql = Sql & "INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N")
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
        vContaFra.NumeroFactura = DBLet(Rs!numfactu)
        vContaFra.Anofac = DBLet(Rs!anofaccl)
        vContaFra.Serie = DBLet(Rs!numserie)
    
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = 0
        DtoGnral = 0
        BaseImp = Rs!baseimpo
        '---- Laura 10/10/2006:  añadir el totalfac para utilizarlo en insertar lineas
        TotalFac = Rs!TotalFac
        '----
        
        PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CodiIVA, "N")
        
        Sql = ""
        Sql = DBSet(Rs!numserie, "T") & "," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!fecfactu) & "," & ValorNulo & ","
        
        If vParamAplic.ContabilidadNueva Then
            ' para el caso de las rectificativas
            Dim vTipM As String
            vTipM = DevuelveValor("select codtipom from stipom where letraser = " & DBSet(Rs!letraser, "T"))
            If vTipM = "FAR" Then
                Sql = Sql & "'D',"
            Else
                Sql = Sql & "'0',"
            End If
            Dim FP As Currency
            FP = 0
            Sql = Sql & "0," & DBSet(FP, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(IvaImp, "N") & ","
            Sql = Sql & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,"
            Sql = Sql & DBSet(Rs!nomsocio, "T") & "," & DBSet(Rs!dirsocio, "T") & "," & DBSet(Rs!pobsocio, "T") & "," & DBSet(Rs!codpostal, "T") & ","
            Sql = Sql & DBSet(Rs!prosocio, "T") & "," & DBSet(Rs!nifSocio, "T") & ",'ES',1"
        
            cad = cad & "(" & Sql & ")"
        
        
            Sql = "INSERT INTO factcli (numserie,numfactu,fecfactu,fecliqcl,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
            Sql = Sql & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
            Sql = Sql & "codpais,codagente)"
            Sql = Sql & " VALUES " & cad
            ConnConta.Execute Sql
    '***
            CadenaInsertFaclin2 = ""
                
            
            'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
            'IVA 1, siempre existe
            Sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
            Sql2 = Sql2 & "1," & DBSet(Rs!baseimpo, "N") & "," & Rs!TipoIVA & "," & DBSet(Rs!porc_iva, "N") & ","
            Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!ImporIva, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
        
            Sql = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
            Sql = Sql & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
            ConnConta.Execute Sql
            
        
        Else
        
            Sql = Sql & DBSet(Rs!baseimpo, "N") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(PorcIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N", "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(CodiIVA, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            cad = cad & "(" & Sql & ")"
        
            'Insertar en la contabilidad
            Sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,fecliqcl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
            Sql = Sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
            Sql = Sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien) "
            Sql = Sql & " VALUES " & cad
            ConnConta.Execute Sql
    
        End If
    
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactTEL = False
        cadErr = Err.Description
    Else
        InsertarCabFactTEL = True
    End If
End Function






Private Function InsertarLinFact_new(cadTabla As String, cadWHERE As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String

    On Error GoTo EInLinea
    

    If cadTabla = "facturas" Then 'VENTAS
         'comprobar si el cliente utiliza cuenta alternativa
        If conCtaAlt Then
            'utilizamos sfamia.ctavent1 o sfamia.abovent1
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctavent1"
            Else
                cadCampo = "sfamia.abovent1" 'si es negativa es un abono
            End If
        Else
            'utilizamos sfamia.ctaventa o sfamia.aboventa
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctaventa"
            Else
                cadCampo = "sfamia.aboventa"
            End If
        End If
'   select concat(raizctavtas, right(concat('000000',codvarie),5)) as cuenta from variedades
        numnivel = DevuelveDesdeBDNew(cConta, "empresa", "numnivel", "codempre", vParamAplic.NumeroConta, "N")
        NumDigit = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & numnivel, "codempre", vParamAplic.NumeroConta, "N")
        NumDigit3 = DevuelveDesdeBDNew(cConta, "empresa", "numdigi3", "codempre", vParamAplic.NumeroConta, "N")
        
'        NumDigitAnt = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & NumNivel - 1, "codempre", vParamAplic.NumeroConta, "N")
        
'        CadCampo1 = "concat(concat(variedades.raizctavtas,tipomer.digicont), right(concat('0000000000',albaran_variedad.codvarie)," & (CCur(NumDigit) - CCur(NumDigit3) - 1) & "))" 'CCur(NumDigitAnt) + 1) & "))"
        CadCampo1 = "CASE tipomer.tiptimer WHEN 0 THEN ctavtasinterior WHEN 1 THEN ctavtasexportacion WHEN 2 THEN ctavtasindustria WHEN 3 THEN ctavtasretirada WHEN 4 THEN ctavtasotros END"
        
        
        
        If vEmpresa.TieneAnalitica Then
            Sql = " SELECT stipom.letraser,facturas_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sfamia.codccost "
        Else
            Sql = " SELECT stipom.letraser,facturas_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe "
        End If
        
        Sql = Sql & " FROM ((facturas_envases inner join stipom on facturas_envases.codtipom=stipom.codtipom) "
        Sql = Sql & " inner join sartic on facturas_envases.codartic=sartic.codartic) "
        Sql = Sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "facturas", "facturas_envases")
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & " GROUP BY 5,7 " '& cadCampo, codccost
        Else
            Sql = Sql & " GROUP BY 5 " '& cadCampo
        End If
        Sql = Sql & "Union "
        
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & " SELECT stipom.letraser,facturas_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe, variedades.codccost "
        Else
            Sql = Sql & " SELECT stipom.letraser,facturas_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe "
        End If
        Sql = Sql & " FROM (((((facturas_variedad inner join stipom on facturas_variedad.codtipom=stipom.codtipom) "
        Sql = Sql & " inner join albaran on facturas_variedad.numalbar = albaran.numalbar) "
        Sql = Sql & " inner join tipomer on albaran.codtimer = tipomer.codtimer) "
        Sql = Sql & " inner join albaran_variedad on facturas_variedad.numalbar = albaran_variedad.numalbar and facturas_variedad.numlinealbar = albaran_variedad.numlinea) "
        Sql = Sql & " inner join variedades on albaran_variedad.codvarie=variedades.codvarie) "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "facturas", "facturas_variedad")
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & " GROUP BY 5,7 " '& cadCampo1, codccost
        Else
            Sql = Sql & " GROUP BY 5 " '& cadCampo1
        End If
        
    Else
        If cadTabla = "scafpc" Then 'COMPRAS
            'utilizamos sfamia.ctaventa o sfamia.aboventa
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctacompr"
            Else
                cadCampo = "sfamia.abocompr"
            End If
            If vEmpresa.TieneAnalitica Then
                Sql = " SELECT slifpc.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sfamia.codccost"
            Else
                Sql = " SELECT slifpc.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe"
            End If
            Sql = Sql & " FROM (slifpc  "
            Sql = Sql & " inner join sartic on slifpc.codartic=sartic.codartic) "
            Sql = Sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
            Sql = Sql & " WHERE " & Replace(cadWHERE, "scafpc", "slifpc")
            Sql = Sql & " GROUP BY " & cadCampo
            If vEmpresa.TieneAnalitica Then
                Sql = Sql & ", sfamia.codccost "
            End If
        Else ' FACTURAS DE TRANSPORTE
            'utilizamos sparam.ctaventa o sparam.aboventa
'            If TotalFac >= 0 Then
'                cadCampo = vParamAplic.CtaTraReten
'            Else
'                cadCampo = vParamAplic.CtaAboTrans
'            End If
'            Sql = " SELECT tlifpc.codtrans,numfactu,fecfactu,'" & cadCampo & "' as cuenta,sum(importel) as importe "
'            Sql = Sql & " FROM tlifpc  "
'            Sql = Sql & " WHERE " & Replace(cadWhere, "tcafpc", "tlifpc")
'            Sql = Sql & " GROUP BY '" & cadCampo & "'"


'++monica: si tipomercado = 1(exportacion) cogemos  variedades.ctatraexporta
'          si tipomercado <> 1 (distinto de exportacion) cogemos  variedades.ctatrainterior
            If vEmpresa.TieneAnalitica Then
                 Sql = " SELECT 2, variedades.ctacomtercero as cuenta, sum(rlifter.importel) as importe, variedades.codccost "
            Else
                 Sql = " SELECT 2, variedades.ctacomtercero as cuenta, sum(rlifter.importel) as importe "
            End If
             Sql = Sql & " FROM rlifter, variedades "
             Sql = Sql & " WHERE " & Replace(cadWHERE, "rcafter", "rlifter") & " and"
             Sql = Sql & " rlifter.codvarie = variedades.codvarie "
             Sql = Sql & " group by 1,2 "
             
             '[Monica]23/09/2013: concepto de gasto
             Sql = Sql & " union "
             If vEmpresa.TieneAnalitica Then
                Sql = Sql & " select 1, fvarconce.codmacpr as cuenta, rcafter.impcargo as importe, '' "
             Else
                Sql = Sql & " select 1, fvarconce.codmacpr as cuenta, rcafter.impcargo as importe "
             End If
             Sql = Sql & " FROM rcafter, fvarconce "
             Sql = Sql & " WHERE " & cadWHERE & " and"
             Sql = Sql & " rcafter.concepcargo = fvarconce.codconce "
             Sql = Sql & " group by 1,2 "
             
             Sql = Sql & " order by 1,2 "


        End If
    End If
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
        ImpLinea = Rs!Importe - CCur(CalcularPorcentaje(Rs!Importe, DtoPPago, 2))
        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(Rs!Importe, DtoGnral, 2))
        'ImpLinea = Round(ImpLinea, 2)
        '----
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        If cadTabla = "facturas" Then 'VENTAS a clientes
            Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
            Sql = Sql & DBSet(Rs!cuenta, "T")
'            If Not conCtaAlt Then 'cliente no tiene cuenta alternativa
'                If ImpLinea >= 0 Then
'                    SQL = SQL & DBSet(RS!ctaventa, "T")
'                Else
'                    SQL = SQL & DBSet(RS!aboventa, "T")
'                End If
'            Else
'                If ImpLinea >= 0 Then
'                    SQL = SQL & DBSet(RS!ctavent1, "T")
'                Else
'                    SQL = SQL & DBSet(RS!abovent1, "T")
'                End If
'            End If
        Else
            If cadTabla = "scafpc" Then 'COMPRAS
                'Laura 24/10/2006
                'SQL = numRegis & "," & Year(RS!FecFactu) & "," & i & ","
                Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
                
    '            If ImpLinea >= 0 Then
                    Sql = Sql & DBSet(Rs!cuenta, "T")
    '            Else
    '                SQL = SQL & DBSet(RS!abocompr, "T")
    '            End If
            Else 'TRANSPORTE
                Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
                Sql = Sql & DBSet(Rs!cuenta, "T")
            End If
        End If
        
        Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
            If cadTabla = "rcafter" Then
                If DBLet(Rs!CodCCost, "T") = "----" Then
                    Sql = Sql & DBSet(CCoste, "T")
                Else
                    Sql = Sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBLet(Rs!CodCCost, "T")
                End If
            Else
                Sql = Sql & DBSet(Rs!CodCCost, "T")
                CCoste = DBSet(Rs!CodCCost, "T")
            End If
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        cad = cad & "(" & Sql & ")" & ","
        
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)
        Sql2 = Sql2 & DBSet(totimp, "N") & ","
        '[Monica]19/09/2013: Fallaba por el valor nulo de coste
        If CCoste = "" Or CCoste = ValorNulo Then
            Sql2 = Sql2 & ValorNulo
        Else
            Sql2 = Sql2 & DBSet(CCoste, "T")
        End If
        If SQLaux <> "" Then 'hay mas de una linea
            cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If


    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        If cadTabla = "facturas" Then
            Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Else
            Sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        End If
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact_new = False
        cadErr = Err.Description
    Else
        InsertarLinFact_new = True
    End If
End Function


Private Function InsertarLinFact_newContaNueva(cadTabla As String, cadWHERE As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String

Dim NumeroIVA As Byte
Dim k As Byte
Dim HayQueAjustar As Boolean
Dim ImpImva As Currency
Dim ImpREC As Currency
Dim Intracom As String
Dim vIntracom As Integer
Dim EsIntracom As Boolean

    On Error GoTo EInLinea
    
    Intracom = DevuelveValor("select intracom from rcafter where " & cadWHERE)
    EsIntracom = (CInt(Intracom) = 1)
    If vEmpresa.TieneAnalitica Then
         If EsIntracom Then
            Sql = " SELECT 2, variedades.ctacomtercero as cuenta, " & DBSet(vParamAplic.CodIvaIntra, "N") & " codigiva, sum(rlifter.importel) as importe, variedades.codccost "
         Else
            Sql = " SELECT 2, variedades.ctacomtercero as cuenta, variedades.codigiva, sum(rlifter.importel) as importe, variedades.codccost "
         End If
    Else
        If EsIntracom Then
            Sql = " SELECT 2, variedades.ctacomtercero as cuenta, " & DBSet(vParamAplic.CodIvaIntra, "N") & " codigiva, sum(rlifter.importel) as importe "
        Else
            Sql = " SELECT 2, variedades.ctacomtercero as cuenta, variedades.codigiva, sum(rlifter.importel) as importe "
        End If
    End If
     
     Sql = Sql & " FROM rlifter, variedades "
     Sql = Sql & " WHERE " & Replace(cadWHERE, "rcafter", "rlifter") & " and"
     Sql = Sql & " rlifter.codvarie = variedades.codvarie "
     Sql = Sql & " group by 1,2,3 "
     
     '[Monica]23/09/2013: concepto de gasto
     Sql = Sql & " union "
     If vEmpresa.TieneAnalitica Then
        If Intracom Then
            Sql = Sql & " select 1, fvarconce.codmacpr as cuenta, " & DBSet(vParamAplic.CodIvaIntra, "N") & " codigiva, rcafter.impcargo as importe, '' "
        Else
            Sql = Sql & " select 1, fvarconce.codmacpr as cuenta, fvarconce.tipoiva codigiva, rcafter.impcargo as importe, '' "
        End If
     Else
        If EsIntracom Then
            Sql = Sql & " select 1, fvarconce.codmacpr as cuenta, " & DBSet(vParamAplic.CodIvaIntra, "N") & " codigiva, rcafter.impcargo as importe "
        Else
            Sql = Sql & " select 1, fvarconce.codmacpr as cuenta, fvarconce.tipoiva codigiva, rcafter.impcargo as importe "
        End If
     End If
     Sql = Sql & " FROM rcafter, fvarconce "
     Sql = Sql & " WHERE " & cadWHERE & " and"
     Sql = Sql & " rcafter.concepcargo = fvarconce.codconce "
     Sql = Sql & " group by 1,2,3 "
     
     Sql = Sql & " order by 3,1,2 "


    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenKeyset, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
        ImpLinea = Rs!Importe - CCur(CalcularPorcentaje(Rs!Importe, DtoPPago, 2))
        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(Rs!Importe, DtoGnral, 2))
        'ImpLinea = Round(ImpLinea, 2)
        '----
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        Sql = DBSet(SerieFraPro, "T") & "," & NumRegis & "," & DBSet(FecRecep, "F") & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(Rs!cuenta, "T")
        Sql = Sql & ","
        
        If vEmpresa.TieneAnalitica Then
            If cadTabla = "rcafter" Then
                If DBLet(Rs!CodCCost, "T") = "----" Then
                    Sql = Sql & DBSet(CCoste, "T")
                Else
                    Sql = Sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBLet(Rs!CodCCost, "T")
                End If
            Else
                Sql = Sql & DBSet(Rs!CodCCost, "T")
                CCoste = DBSet(Rs!CodCCost, "T")
            End If
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        '$$$
       'Vemos que tipo de IVA es en el vector de importes
        NumeroIVA = 127
        For k = 0 To 2
            If Rs!Codigiva = vTipoIva(k) Then
                NumeroIVA = k
                Exit For
            End If
        Next
        If NumeroIVA > 100 Then Err.Raise 513, "Error obteniendo IVA: " & Rs!Codigiva
        
        vBaseIva(NumeroIVA) = vBaseIva(NumeroIVA) - ImpLinea   'Para ajustar el importe y que no haya descuadre
        
        'Caluclo el importe de IVA y el de recargo de equivalencia
        ImpImva = vPorcIva(NumeroIVA) / 100
        ImpImva = Round2(ImpLinea * ImpImva, 2)
        If vPorcRec(NumeroIVA) = 0 Then
            ImpREC = 0
        Else
            ImpREC = vPorcRec(NumeroIVA) / 100
            ImpREC = Round2(ImpLinea * ImpREC, 2)
        End If
        vImpIva(NumeroIVA) = vImpIva(NumeroIVA) - ImpImva
        vImpRec(NumeroIVA) = vImpRec(NumeroIVA) - ImpREC
        
        
        
        HayQueAjustar = False
        If vBaseIva(NumeroIVA) <> 0 Or vImpIva(NumeroIVA) <> 0 Or vImpRec(NumeroIVA) <> 0 Then
            'falta importe.
            'Puede ser que hayan mas lineas, o haya descuadre. Como esta ordenado por tipo de iva
            Rs.MoveNext
            If Rs.EOF Then
                'No hay mas lineas
                'Hay que ajustar SI o SI
                HayQueAjustar = True
            Else
                'Si que hay mas lineas.
                'Son del mismo tipo de IVA
                If Rs!Codigiva <> vTipoIva(NumeroIVA) Then
                    'NO es el mismo tipo de IVA
                    'Hay que ajustar
                    HayQueAjustar = True
                End If
            End If
            Rs.MovePrevious
        End If
        
        Sql = Sql & "," & vTipoIva(NumeroIVA) & "," & DBSet(vPorcIva(NumeroIVA), "N") & "," & DBSet(vPorcRec(NumeroIVA), "N", "S") & ","
        

        
        If HayQueAjustar Then
            If vBaseIva(NumeroIVA) <> 0 Then ImpLinea = ImpLinea + vBaseIva(NumeroIVA)
            If vImpIva(NumeroIVA) <> 0 Then ImpImva = ImpImva + vImpIva(NumeroIVA)
            If vImpRec(NumeroIVA) <> 0 Then ImpREC = ImpREC + vImpRec(NumeroIVA)
        End If
        
        
        
        
        Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        ' baseimpo , impoiva, imporec, aplicret, CodCCost
        Sql = Sql & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpREC, "N", "S")
        Sql = Sql & ",0"

        cad = cad & "(" & Sql & ")" & ","
        
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
'    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
'    'de la factura
'    If totimp <> BaseImp Then
''        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
'        'en SQL esta la ult linea introducida
'        totimp = BaseImp - totimp
'        totimp = ImpLinea + totimp '(+- diferencia)
'        Sql2 = Sql2 & DBSet(totimp, "N") & ","
'        '[Monica]19/09/2013: Fallaba por el valor nulo de coste
'        If CCoste = "" Or CCoste = ValorNulo Then
'            Sql2 = Sql2 & ValorNulo
'        Else
'            Sql2 = Sql2 & DBSet(CCoste, "T")
'        End If
'        If SQLaux <> "" Then 'hay mas de una linea
'            Cad = SQLaux & "(" & Sql2 & ")" & ","
'        Else 'solo una linea
'            Cad = "(" & Sql2 & ")" & ","
'        End If
'
''        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
''        cad = Replace(cad, SQL, Aux)
'    End If


    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO factpro_lineas (numserie,numregis,fecharec,anofactu,numlinea,codmacta,codccost,codigiva,porciva,porcrec,baseimpo,impoiva,imporec,aplicret) "
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact_newContaNueva = False
        cadErr = Err.Description
    Else
        InsertarLinFact_newContaNueva = True
    End If
End Function






Private Function InsertarLinFactADV(cadTabla As String, cadWHERE As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String

    On Error GoTo EInLinea
    

    If cadTabla = "advfacturas" Then 'VENTAS a socios
        'utilizamos sfamia.ctaventa o sfamia.aboventa
        If TotalFac >= 0 Then
            cadCampo = "advfamia.ctaventa"
        Else
            cadCampo = "advfamia.aboventa"
        End If
        
        If vEmpresa.TieneAnalitica Then
            Sql = " SELECT usuarios.stipom.letraser,advfacturas_lineas.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, advfamia.codccost "
        Else
            Sql = " SELECT usuarios.stipom.letraser,advfacturas_lineas.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe "
        End If
        
        Sql = Sql & " FROM ((advfacturas_lineas inner join usuarios.stipom on advfacturas_lineas.codtipom=usuarios.stipom.codtipom) "
        Sql = Sql & " inner join advartic on advfacturas_lineas.codartic=advartic.codartic) "
        Sql = Sql & " inner join advfamia on advartic.codfamia=advfamia.codfamia "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "advfacturas", "advfacturas_lineas")
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & " GROUP BY 5,7 " '& cadCampo, codccost
        Else
            Sql = Sql & " GROUP BY 5 " '& cadCampo
        End If
        
    End If
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
' --monica:no hay descuentos
'        '---- Laura: 10/10/2006
'        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
'        ImpLinea = RS!Importe - CCur(CalcularPorcentaje(RS!Importe, DtoPPago, 2))
'        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
'        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(RS!Importe, DtoGnral, 2))
'        'ImpLinea = Round(ImpLinea, 2)
'        '----
'        TotImp = TotImp + ImpLinea
        ImpLinea = DBLet(Rs!Importe, "N")
        totimp = totimp + DBLet(Rs!Importe, "N")

        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        If cadTabla = "advfacturas" Then 'VENTAS a socios
            Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
            Sql = Sql & DBSet(Rs!cuenta, "T")
        End If
        
        Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & DBSet(Rs!CodCCost, "T")
            CCoste = DBSet(Rs!CodCCost, "T")
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        cad = cad & "(" & Sql & ")" & ","
        
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)
        Sql2 = Sql2 & DBSet(totimp, "N") & ","
        If CCoste = "" Or CCoste = ValorNulo Then
            Sql2 = Sql2 & ValorNulo
        Else
            Sql2 = Sql2 & DBSet(CCoste, "T")
        End If
        If SQLaux <> "" Then 'hay mas de una linea
            cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If


    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactADV = False
        cadErr = Err.Description
    Else
        InsertarLinFactADV = True
    End If
End Function


Private Function InsertarLinFactADVContaNueva(cadTabla As String, cadWHERE As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String


Dim NumeroIVA As Byte
Dim k As Integer
Dim HayQueAjustar As Boolean
Dim ImpImva As Currency
Dim ImpREC As Currency


    On Error GoTo EInLinea
    

    If cadTabla = "advfacturas" Then 'VENTAS a socios
        'utilizamos sfamia.ctaventa o sfamia.aboventa
        If TotalFac >= 0 Then
            cadCampo = "advfamia.ctaventa"
        Else
            cadCampo = "advfamia.aboventa"
        End If
        
        If vEmpresa.TieneAnalitica Then
            Sql = " SELECT usuarios.stipom.letraser,advfacturas_lineas.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,advfacturas_lineas.codigiva, tiposiva.porceiva porciva,  tiposiva.porcerec porcrec,sum(importel) as importe, advfamia.codccost "
        Else
            Sql = " SELECT usuarios.stipom.letraser,advfacturas_lineas.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,advfacturas_lineas.codigiva, tiposiva.porceiva porciva,  tiposiva.porcerec porcrec,sum(importel) as importe "
        End If
        
        Sql = Sql & " FROM (((advfacturas_lineas inner join usuarios.stipom on advfacturas_lineas.codtipom=usuarios.stipom.codtipom) "
        Sql = Sql & " inner join advartic on advfacturas_lineas.codartic=advartic.codartic) "
        Sql = Sql & " inner join advfamia on advartic.codfamia=advfamia.codfamia) "
        Sql = Sql & " inner join ariconta" & vParamAplic.NumeroConta & ".tiposiva on ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = advfacturas_lineas.codigiva "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "advfacturas", "advfacturas_lineas")
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & " GROUP BY 5,6,7,8, 10 " '& cadCampo, codccost
        Else
            Sql = Sql & " GROUP BY 5,6,7,8 " '& cadCampo
        End If
        Sql = Sql & " ORDER BY 6,5 "
    End If
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenKeyset, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
' --monica:no hay descuentos
'        '---- Laura: 10/10/2006
'        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
'        ImpLinea = RS!Importe - CCur(CalcularPorcentaje(RS!Importe, DtoPPago, 2))
'        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
'        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(RS!Importe, DtoGnral, 2))
'        'ImpLinea = Round(ImpLinea, 2)
'        '----
'        TotImp = TotImp + ImpLinea
        ImpLinea = DBLet(Rs!Importe, "N")
        totimp = totimp + DBLet(Rs!Importe, "N")

        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        If cadTabla = "advfacturas" Then 'VENTAS a socios
            Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
            Sql = Sql & DBSet(Rs!cuenta, "T") & ","
        End If
        
        'Vemos que tipo de IVA es en el vector de importes
        NumeroIVA = 127
        For k = 0 To 2
            If Rs!Codigiva = vTipoIva(k) Then
                NumeroIVA = k
                Exit For
            End If
        Next
        If NumeroIVA > 100 Then Err.Raise 513, "Error obteniendo IVA: " & Rs!Codigiva
        
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & DBSet(Rs!CodCCost, "T")
            CCoste = DBSet(Rs!CodCCost, "T")
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        Sql = Sql & "," & DBSet(Rs!fecfactu, "F")
        
        
        vBaseIva(NumeroIVA) = vBaseIva(NumeroIVA) - ImpLinea   'Para ajustar el importe y que no haya descuadre
        
        'Caluclo el importe de IVA y el de recargo de equivalencia
        ImpImva = vPorcIva(NumeroIVA) / 100
        ImpImva = Round2(ImpLinea * ImpImva, 2)
        If vPorcRec(NumeroIVA) = 0 Then
            ImpREC = 0
        Else
            ImpREC = vPorcRec(NumeroIVA) / 100
            ImpREC = Round2(ImpLinea * ImpREC, 2)
        End If
        vImpIva(NumeroIVA) = vImpIva(NumeroIVA) - ImpImva
        vImpRec(NumeroIVA) = vImpRec(NumeroIVA) - ImpREC
        
        
        
        HayQueAjustar = False
        If vBaseIva(NumeroIVA) <> 0 Or vImpIva(NumeroIVA) <> 0 Or vImpRec(NumeroIVA) <> 0 Then
            'falta importe.
            'Puede ser que hayan mas lineas, o haya descuadre. Como esta ordenado por tipo de iva
            Rs.MoveNext
            If Rs.EOF Then
                'No hay mas lineas
                'Hay que ajustar SI o SI
                HayQueAjustar = True
            Else
                'Si que hay mas lineas.
                'Son del mismo tipo de IVA
                If Rs!Codigiva <> vTipoIva(NumeroIVA) Then
                    'NO es el mismo tipo de IVA
                    'Hay que ajustar
                    HayQueAjustar = True
                End If
            End If
            Rs.MovePrevious
        End If
        
        Sql = Sql & "," & vTipoIva(NumeroIVA) & "," & DBSet(vPorcIva(NumeroIVA), "N") & "," & DBSet(vPorcRec(NumeroIVA), "N", "S") & ","
        
        If HayQueAjustar Then
            If vBaseIva(NumeroIVA) <> 0 Then ImpLinea = ImpLinea + vBaseIva(NumeroIVA)
            If vImpIva(NumeroIVA) <> 0 Then ImpImva = ImpImva + vImpIva(NumeroIVA)
            If vImpRec(NumeroIVA) <> 0 Then ImpREC = ImpREC + vImpRec(NumeroIVA)
        End If

        
        
        ' baseimpo , impoiva, imporec
        Sql = Sql & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpREC, "N", "S")
        
   
        
        cad = cad & "(" & Sql & ")" & ","
        
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
'    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
'    'de la factura
'    If totimp <> BaseImp Then
''        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
'        'en SQL esta la ult linea introducida
'        totimp = BaseImp - totimp
'        totimp = ImpLinea + totimp '(+- diferencia)
'        Sql2 = Sql2 & DBSet(totimp, "N") & ","
'        If CCoste = "" Or CCoste = ValorNulo Then
'            Sql2 = Sql2 & ValorNulo
'        Else
'            Sql2 = Sql2 & DBSet(CCoste, "T")
'        End If
'        If SQLaux <> "" Then 'hay mas de una linea
'            Cad = SQLaux & "(" & Sql2 & ")" & ","
'        Else 'solo una linea
'            Cad = "(" & Sql2 & ")" & ","
'        End If
'
''        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
''        cad = Replace(cad, SQL, Aux)
'    End If


    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO factcli_lineas (numserie,numfactu,anofactu,numlinea,codmacta,codccost,fecfactu,codigiva,porciva,porcrec,baseimpo,impoiva,imporec) "
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactADVContaNueva = False
        cadErr = Err.Description
    Else
        InsertarLinFactADVContaNueva = True
    End If
End Function




Private Function InsertarLinFactBOD(cadTabla As String, cadWHERE As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String

    On Error GoTo EInLinea
    
        
    Select Case cadTabla
        Case "rbodfacturas" 'facturas de retirada de almazara y bodega
            'utilizamos variedades.ctaventa o variedades.aboventa
            If TotalFac >= 0 Then
                cadCampo = "variedades.ctavtasotros"
            Else
                cadCampo = "variedades.aboventa"
            End If
            
        Case "rbodfact1" ' lineas de variedades de almazara
            cadCampo = vParamAplic.CtaVentasAlmz
        
        Case "rbodfact2" ' lineas de variedades de bodega
            cadCampo = vParamAplic.CtaVentasBOD
    End Select
    
    If vEmpresa.TieneAnalitica Then
        Sql = " SELECT usuarios.stipom.letraser,rbodfacturas_lineas.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, variedades.codccost "
    Else
        Sql = " SELECT usuarios.stipom.letraser,rbodfacturas_lineas.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe "
    End If
    
    Sql = Sql & " FROM (rbodfacturas_lineas inner join usuarios.stipom on rbodfacturas_lineas.codtipom=usuarios.stipom.codtipom) "
    Sql = Sql & " inner join variedades on rbodfacturas_lineas.codvarie=variedades.codvarie "
    Sql = Sql & " WHERE " & Replace(cadWHERE, "rbodfacturas", "rbodfacturas_lineas")
    If vEmpresa.TieneAnalitica Then
        Sql = Sql & " GROUP BY 5,7 " '& cadCampo, codccost
    Else
        Sql = Sql & " GROUP BY 5 " '& cadCampo
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
' --monica:no hay descuentos
'        '---- Laura: 10/10/2006
'        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
'        ImpLinea = RS!Importe - CCur(CalcularPorcentaje(RS!Importe, DtoPPago, 2))
'        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
'        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(RS!Importe, DtoGnral, 2))
'        'ImpLinea = Round(ImpLinea, 2)
'        '----
'        TotImp = TotImp + ImpLinea
        ImpLinea = DBLet(Rs!Importe, "N")
        totimp = totimp + DBLet(Rs!Importe, "N")

        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
        Sql = Sql & DBSet(Rs!cuenta, "T")
        
        Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & DBSet(Rs!CodCCost, "T")
            CCoste = DBSet(Rs!CodCCost, "T")
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        cad = cad & "(" & Sql & ")" & ","
        
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)
        Sql2 = Sql2 & DBSet(totimp, "N") & ","
        If CCoste = "" Or CCoste = ValorNulo Then
            Sql2 = Sql2 & ValorNulo
        Else
            Sql2 = Sql2 & DBSet(CCoste, "T")
        End If
        If SQLaux <> "" Then 'hay mas de una linea
            cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If


    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactBOD = False
        cadErr = Err.Description
    Else
        InsertarLinFactBOD = True
    End If
End Function


Private Function InsertarLinFactBODContaNueva(cadTabla As String, cadWHERE As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String

Dim NumeroIVA As Byte
Dim k As Integer
Dim HayQueAjustar As Boolean

Dim ImpImva As Currency
Dim ImpREC As Currency



    On Error GoTo EInLinea
    
        
    Select Case cadTabla
        Case "rbodfacturas" 'facturas de retirada de almazara y bodega
            'utilizamos variedades.ctaventa o variedades.aboventa
            If TotalFac >= 0 Then
                cadCampo = "variedades.ctavtasotros"
            Else
                cadCampo = "variedades.aboventa"
            End If
            
        Case "rbodfact1" ' lineas de variedades de almazara
            cadCampo = vParamAplic.CtaVentasAlmz
        
        Case "rbodfact2" ' lineas de variedades de bodega
            cadCampo = vParamAplic.CtaVentasBOD
    End Select
    
    If vEmpresa.TieneAnalitica Then
        Sql = " SELECT usuarios.stipom.letraser,rbodfacturas_lineas.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,rbodfacturas_lineas.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec,sum(importel) as importe, variedades.codccost "
    Else
        Sql = " SELECT usuarios.stipom.letraser,rbodfacturas_lineas.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,rbodfacturas_lineas.codigiva, tiposiva.porceiva porciva, tiposiva.porcerec porcrec,sum(importel) as importe "
    End If
    
    Sql = Sql & " FROM ((rbodfacturas_lineas inner join usuarios.stipom on rbodfacturas_lineas.codtipom=usuarios.stipom.codtipom) "
    Sql = Sql & " inner join variedades on rbodfacturas_lineas.codvarie=variedades.codvarie) "
    Sql = Sql & " inner join ariconta" & vParamAplic.NumeroConta & ".tiposiva on ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = rbodfacturas_lineas.codigiva "
    
    Sql = Sql & " WHERE " & Replace(cadWHERE, "rbodfacturas", "rbodfacturas_lineas")
    If vEmpresa.TieneAnalitica Then
        Sql = Sql & " GROUP BY 5,7 " '& cadCampo, codccost
    Else
        Sql = Sql & " GROUP BY 5 " '& cadCampo
    End If
    Sql = Sql & " ORDER BY 6,5 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenKeyset, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
' --monica:no hay descuentos
'        '---- Laura: 10/10/2006
'        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
'        ImpLinea = RS!Importe - CCur(CalcularPorcentaje(RS!Importe, DtoPPago, 2))
'        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
'        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(RS!Importe, DtoGnral, 2))
'        'ImpLinea = Round(ImpLinea, 2)
'        '----
'        TotImp = TotImp + ImpLinea
        ImpLinea = DBLet(Rs!Importe, "N")
        totimp = totimp + DBLet(Rs!Importe, "N")

        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
        Sql = Sql & DBSet(Rs!cuenta, "T")
        
        'Vemos que tipo de IVA es en el vector de importes
        NumeroIVA = 127
        For k = 0 To 2
            If Rs!Codigiva = vTipoIva(k) Then
                NumeroIVA = k
                Exit For
            End If
        Next
        If NumeroIVA > 100 Then Err.Raise 513, "Error obteniendo IVA: " & Rs!Codigiva
        
        
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & DBSet(Rs!CodCCost, "T")
            CCoste = DBSet(Rs!CodCCost, "T")
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        Sql = Sql & "," & DBSet(Rs!fecfactu, "F")
        
        
        vBaseIva(NumeroIVA) = vBaseIva(NumeroIVA) - ImpLinea   'Para ajustar el importe y que no haya descuadre
        HayQueAjustar = False
        If vBaseIva(NumeroIVA) <> 0 Then
            'falta importe.
            'Puede ser que hayan mas lineas, o haya descuadre. Como esta ordenado por tipo de iva
            Rs.MoveNext
            If Rs.EOF Then
                'No hay mas lineas
                'Hay que ajustar SI o SI
                HayQueAjustar = True
            Else
                'Si que hay mas lineas.
                'Son del mismo tipo de IVA
                If Rs!Codigiva <> vTipoIva(0) Then
                    'NO es el mismo tipo de IVA
                    'Hay que ajustar
                    HayQueAjustar = True
                End If
            End If
            Rs.MovePrevious
        End If
        
        Sql = Sql & "," & vTipoIva(NumeroIVA) & "," & DBSet(vPorcIva(NumeroIVA), "N") & "," & DBSet(vPorcRec(NumeroIVA), "N", "S") & ","
        
        If HayQueAjustar Then
            Stop
        Else
        
        End If

        
        'Caluclo el importe de IVA y el de recargo de equivalencia
        ImpImva = vPorcIva(NumeroIVA) / 100
        ImpImva = Round2(ImpLinea * ImpImva, 2)
        If vPorcRec(NumeroIVA) = 0 Then
            ImpREC = 0
        Else
            ImpREC = vPorcRec(NumeroIVA) / 100
            ImpREC = Round2(ImpLinea * ImpREC, 2)
        End If
        vImpIva(NumeroIVA) = vImpIva(NumeroIVA) - ImpImva
        vImpRec(NumeroIVA) = vImpRec(NumeroIVA) - ImpREC
        
        
        ' baseimpo , impoiva, imporec
        Sql = Sql & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpREC, "N", "S")
        
        cad = cad & "(" & Sql & ")" & ","
        
        I = I + 1
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
'    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
'    'de la factura
'    If totimp <> BaseImp Then
''        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
'        'en SQL esta la ult linea introducida
'        totimp = BaseImp - totimp
'        totimp = ImpLinea + totimp '(+- diferencia)
'        Sql2 = Sql2 & DBSet(totimp, "N") & ","
'        If CCoste = "" Or CCoste = ValorNulo Then
'            Sql2 = Sql2 & ValorNulo
'        Else
'            Sql2 = Sql2 & DBSet(CCoste, "T")
'        End If
'        If SQLaux <> "" Then 'hay mas de una linea
'            Cad = SQLaux & "(" & Sql2 & ")" & ","
'        Else 'solo una linea
'            Cad = "(" & Sql2 & ")" & ","
'        End If
'
''        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
''        cad = Replace(cad, SQL, Aux)
'    End If

    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO factcli_lineas (numserie,numfactu,anofactu,numlinea,codmacta,codccost,fecfactu,codigiva,porciva,porcrec,baseimpo,impoiva,imporec) "
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactBODContaNueva = False
        cadErr = Err.Description
    Else
        InsertarLinFactBODContaNueva = True
    End If
End Function







Private Function InsertarLinFactTEL(CtaVtas As String, cadWHERE As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String

    On Error GoTo EInLinea
    
        
    cadCampo = CtaVtas
    
    If vEmpresa.TieneAnalitica Then
        Sql = " SELECT numserie,numfactu,fecfactu," & cadCampo & " as cuenta,baseimpo as importe, " & CCoste
    Else
        Sql = " SELECT numserie,numfactu,fecfactu," & cadCampo & " as cuenta,baseimpo as importe "
    End If
    
    Sql = Sql & " FROM rtelmovil "
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    If Not Rs.EOF Then
        SQLaux = cad
        
        ImpLinea = DBLet(Rs!Importe, "N")
        totimp = totimp + DBLet(Rs!Importe, "N")

        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = "'" & Rs!numserie & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
        Sql = Sql & DBSet(Rs!cuenta, "T")
        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & DBSet(Rs!CodCCost, "T")
            CCoste = DBSet(Rs!CodCCost, "T")
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        cad = cad & "(" & Sql & ")"
        
        I = I + 1
        Rs.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
    

    'Insertar en la contabilidad
    If cad <> "" Then
        Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactTEL = False
        cadErr = Err.Description
    Else
        InsertarLinFactTEL = True
    End If
End Function


Private Function InsertarLinFactTELContaNueva(CtaVtas As String, cadWHERE As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String

    On Error GoTo EInLinea
    
        
    cadCampo = CtaVtas
    
    If vEmpresa.TieneAnalitica Then
        Sql = " SELECT numserie,numfactu,fecfactu," & cadCampo & " as cuenta,baseimpo as importe, " & CCoste
    Else
        Sql = " SELECT numserie,numfactu,fecfactu," & cadCampo & " as cuenta,baseimpo as importe "
    End If
    
    Sql = Sql & " FROM rtelmovil "
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    If Not Rs.EOF Then
        SQLaux = cad
        
        ImpLinea = DBLet(Rs!Importe, "N")
        totimp = totimp + DBLet(Rs!Importe, "N")

        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = "'" & Rs!numserie & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
        Sql = Sql & DBSet(Rs!cuenta, "T")
        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & DBSet(Rs!CodCCost, "T")
            CCoste = DBSet(Rs!CodCCost, "T")
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        cad = cad & "(" & Sql & ")"
        
        I = I + 1
        Rs.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
    

    'Insertar en la contabilidad
    If cad <> "" Then
        Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactTELContaNueva = False
        cadErr = Err.Description
    Else
        InsertarLinFactTELContaNueva = True
    End If
End Function





Private Function InsertarLinFactSoc(cadTabla As String, cadWHERE As String, cadErr As String, Tipo As Byte, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String
Dim LineaVariedad As Integer

Dim vSocio As cSocio
Dim Socio As String
Dim TipoAnt As Byte
Dim TipoFact As String

Dim ImpAnticipo As Currency
    
    On Error GoTo EInLinea
    
    TipoAnt = Tipo
'    TipoFactAnt = TipoFact
    
    If Tipo = 11 Then ' si es una factura rectificativa cojo el tipo de movimiento de la factura que rectifico
        Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(CodTipomRECT, "T"))
        
        TipoFact = CodTipomRECT

    Else
' Estoy aqui: en liquidacion de industria

'select if(rsocios.tipoprod = 1, variedades.ctacomtercero, variedades.ctaliquidacion) as cuenta
'From rsocios, Variedades, rfactsoc, rfactsoc_variedad
'where rsocios.codsocio= rfactsoc.codsocio and mid(rfactsoc.codtipom,1,3) = "FLI" and
'rfactsoc.codtipom= rfactsoc_variedad.codtipom and
'rfactsoc.numfactu = rfactsoc_variedad.codtipom and
'rfactsoc.fecfactu = rfactsoc_variedad.fecfactu and
'rfactsoc_variedad.codvarie = Variedades.codvarie

        ' [Monica] 29/12/2009 si es liquidacion de industria miramos que cuenta coger dependiendo de que el socio sea
        ' tercero o no lo sea
        Sql = "select mid(rfactsoc.codtipom,1,3) from " & cadTabla & " where " & cadWHERE
        TipoFact = DevuelveValor(Sql)
    
    End If
    
    If Tipo = 2 And TipoFact = "FLI" Then
        Sql = "select rfactsoc.codsocio from " & cadTabla & " where " & cadWHERE
        Socio = DevuelveValor(Sql)
        
        Set vSocio = New cSocio
        If vSocio.LeerDatos(Socio) Then
            If vEmpresa.TieneAnalitica Then
                If vSocio.TipoProd = 1 Then
                    Sql = " SELECT 1, variedades.ctacomtercero as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Else
                    Sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                End If
            Else
                If vSocio.TipoProd = 1 Then
                    Sql = " SELECT 1, variedades.ctacomtercero as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                Else
                    Sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                End If
            End If
            
            '[Monica]14/11/2014: solo en el caso de Catadau en liquidacion de industria cogemos la ctacomtercero
            If vParamAplic.Cooperativa = 0 Then
                If vEmpresa.TieneAnalitica Then
                    Sql = " SELECT 1, variedades.ctacomtercero as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Else
                    Sql = " SELECT 1, variedades.ctacomtercero as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                End If
            End If
            
            Sql = Sql & " FROM rfactsoc_variedad, variedades "
            Sql = Sql & " WHERE " & Replace(cadWHERE, "rfactsoc", "rfactsoc_variedad") & " and"
            Sql = Sql & " rfactsoc_variedad.codvarie = variedades.codvarie "
            Sql = Sql & " group by 1,2 "
            Sql = Sql & " order by 1,2 "
            
        Else
            InsertarLinFactSoc = False
            Exit Function
        End If
    Else
    ' fin de lo añadido
    
        If vEmpresa.TieneAnalitica Then
            Select Case Tipo
                Case 1, 3, 7, 9
                     Sql = " SELECT 1, variedades.ctaanticipo as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Case 2, 4, 8, 10
                     Sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Case 6 ' siniestros
                     Sql = " SELECT 1, variedades.ctasiniestros as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
            End Select
            If TipoFact = "FTS" Then
                Sql = " SELECT 1, variedades.ctaacarecol as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
            End If
        Else
            Select Case Tipo
                Case 1, 3, 7, 9
                     Sql = " SELECT 1, variedades.ctaanticipo as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                Case 2, 4, 8, 10
                     Sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                Case 6 ' siniestros
                     Sql = " SELECT 1, variedades.ctasiniestros as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
            End Select
            '[Monica]16/07/2014: añadido el caso de tipo transporte tercero de Picassent
            If TipoFact = "FTS" Or TipoFact = "FTT" Then
                Sql = " SELECT 1, variedades.ctaacarecol as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
            End If
        End If
        Sql = Sql & " FROM rfactsoc_variedad, variedades "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "rfactsoc", "rfactsoc_variedad") & " and"
        Sql = Sql & " rfactsoc_variedad.codvarie = variedades.codvarie "
        Sql = Sql & " group by 1,2 "
        Sql = Sql & " order by 1,2 "

    End If



    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        ImpLinea = Rs!Importe
        
        ' si se trata de una liquidacion hemos de descontar los anticipos por variedad
        ' que no sean anticipo de gastos
        If (Tipo = 2 Or Tipo = 4 Or Tipo = 8 Or Tipo = 10) And TipoFact <> "FTS" Then
            Sql3 = "select sum(baseimpo) from rfactsoc_anticipos, variedades "
            Sql3 = Sql3 & " where " & Replace(cadWHERE, "rfactsoc", "rfactsoc_anticipos")
            Sql3 = Sql3 & " and rfactsoc_anticipos.codvarieanti = variedades.codvarie "
            Sql3 = Sql3 & " and variedades.ctaliquidacion = " & DBSet(Rs!cuenta, "N")
            
            ImpAnticipo = DevuelveValor(Sql3)
            
            ImpLinea = ImpLinea - ImpAnticipo
        End If
        '----
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(Rs!cuenta, "T")
        
        Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
                If DBLet(Rs!CodCCost, "T") = "----" Then
                    Sql = Sql & DBSet(CCoste, "T")
                Else
                    Sql = Sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBLet(Rs!CodCCost, "T")
                End If
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)
        Sql2 = Sql2 & DBSet(totimp, "N") & ","
        If CCoste = "" Or CCoste = ValorNulo Then
            Sql2 = Sql2 & ValorNulo
        Else
            Sql2 = Sql2 & DBSet(CCoste, "T")
        End If
        If SQLaux <> "" Then 'hay mas de una linea
            cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If

    ' las retenciones si las hay
    If ImpReten <> 0 Then
        Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(CtaSocio, "T")
        Sql = Sql & "," & DBSet(ImpReten, "N") & ","
        Sql = Sql & ValorNulo ' no llevan centro de coste
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
        
        Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(CtaReten, "T")
        Sql = Sql & "," & DBSet(ImpReten * (-1), "N") & ","
        Sql = Sql & ValorNulo ' no llevan centro de coste
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
    End If
    
    ' las aportaciones de fondo operativo si las hay
    If ImpAport <> 0 Then
        Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(CtaSocio, "T")
        Sql = Sql & "," & DBSet(ImpAport, "N") & ","
        Sql = Sql & ValorNulo ' no llevan centro de coste
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
    
        Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(CtaAport, "T")
        Sql = Sql & "," & DBSet(ImpAport * (-1), "N") & ","
        Sql = Sql & ValorNulo ' no llevan centro de coste
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
    End If
    
    '[Monica]20/12/2013: si es montifrut no descontamos el descuento que tengo grabado a pie
        '[Monica]09/03/2015: para el caso de Catadau tampoco se tienen que insertar las bases correspondientes a gastos
            '[Monica]13/04/2016: levanto el control de que no se contabilicen los gastos en Catadau
    If vParamAplic.Cooperativa <> 12 Then  'And vParamAplic.Cooperativa <> 0 Then
        ' insertamos todos los gastos a pie de factura rfactsoc_gastos
        Sql = " SELECT rconcepgasto.codmacta as cuenta, sum(rfactsoc_gastos.importe) as importe "
        Sql = Sql & " from rconcepgasto INNER JOIN rfactsoc_gastos ON rconcepgasto.codgasto = rfactsoc_gastos.codgasto "
        Sql = Sql & " where " & Replace(cadWHERE, "rfactsoc", "rfactsoc_gastos")
        
        '[Monica]06/06/2016: si es catadau solo los que tengan cuenta
        If vParamAplic.Cooperativa = 0 Then
            Sql = Sql & " and not rconcepgasto.codmacta is null and rconcepgasto.codmacta <> '' "
        End If
        
        Sql = Sql & " group by 1 "
        Sql = Sql & " order by 1 "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        While Not Rs.EOF
            Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
            Sql = Sql & DBSet(CtaSocio, "T")
            Sql = Sql & "," & DBSet(Rs!Importe, "N") & ","
            Sql = Sql & ValorNulo ' no llevan centro de coste
            
            cad = cad & "(" & Sql & ")" & ","
            I = I + 1
        
            Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
            Sql = Sql & DBSet(Rs!cuenta, "T")
            Sql = Sql & "," & DBSet(Rs!Importe * (-1), "N") & ","
            Sql = Sql & ValorNulo ' no llevan centro de coste
            
            cad = cad & "(" & Sql & ")" & ","
            I = I + 1
        
            Rs.MoveNext
        Wend
        Set Rs = Nothing
    End If
    
    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If
    
    Tipo = TipoAnt

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactSoc = False
        cadErr = Err.Description
    Else
        InsertarLinFactSoc = True
    End If
End Function



Private Function InsertarLinFactSocContaNueva(cadTabla As String, cadWHERE As String, cadErr As String, Tipo As Byte, FecRecep As Date, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim SqlAux2 As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String
Dim LineaVariedad As Integer

Dim vSocio As cSocio
Dim Socio As String
Dim TipoAnt As Byte
Dim TipoFact As String

Dim ImpAnticipo As Currency

Dim vTipoIvaAux As Currency
Dim vImpIvaAux As Currency
Dim vPorIvaAux As Currency
Dim impiva As Currency
Dim TotImpIVA As Currency

    On Error GoTo EInLinea
    
    
    
    TipoAnt = Tipo
'    TipoFactAnt = TipoFact
    
    If Tipo = 11 Then ' si es una factura rectificativa cojo el tipo de movimiento de la factura que rectifico
        Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(CodTipomRECT, "T"))
        
        TipoFact = CodTipomRECT

    Else
' Estoy aqui: en liquidacion de industria

'select if(rsocios.tipoprod = 1, variedades.ctacomtercero, variedades.ctaliquidacion) as cuenta
'From rsocios, Variedades, rfactsoc, rfactsoc_variedad
'where rsocios.codsocio= rfactsoc.codsocio and mid(rfactsoc.codtipom,1,3) = "FLI" and
'rfactsoc.codtipom= rfactsoc_variedad.codtipom and
'rfactsoc.numfactu = rfactsoc_variedad.codtipom and
'rfactsoc.fecfactu = rfactsoc_variedad.fecfactu and
'rfactsoc_variedad.codvarie = Variedades.codvarie

        ' [Monica] 29/12/2009 si es liquidacion de industria miramos que cuenta coger dependiendo de que el socio sea
        ' tercero o no lo sea
        Sql = "select mid(rfactsoc.codtipom,1,3) from " & cadTabla & " where " & cadWHERE
        TipoFact = DevuelveValor(Sql)
    
    End If
    
    If Tipo = 2 And TipoFact = "FLI" Then
        Sql = "select rfactsoc.codsocio from " & cadTabla & " where " & cadWHERE
        Socio = DevuelveValor(Sql)
        
        Set vSocio = New cSocio
        If vSocio.LeerDatos(Socio) Then
            If vEmpresa.TieneAnalitica Then
                If vSocio.TipoProd = 1 Then
                    Sql = " SELECT 1, variedades.ctacomtercero as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Else
                    Sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                End If
            Else
                If vSocio.TipoProd = 1 Then
                    Sql = " SELECT 1, variedades.ctacomtercero as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                Else
                    Sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                End If
            End If
            
            '[Monica]14/11/2014: solo en el caso de Catadau en liquidacion de industria cogemos la ctacomtercero
            If vParamAplic.Cooperativa = 0 Then
                If vEmpresa.TieneAnalitica Then
                    Sql = " SELECT 1, variedades.ctacomtercero as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Else
                    Sql = " SELECT 1, variedades.ctacomtercero as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                End If
            End If
            
            Sql = Sql & " FROM rfactsoc_variedad, variedades "
            Sql = Sql & " WHERE " & Replace(cadWHERE, "rfactsoc", "rfactsoc_variedad") & " and"
            Sql = Sql & " rfactsoc_variedad.codvarie = variedades.codvarie "
            Sql = Sql & " group by 1,2 "
            Sql = Sql & " order by 1,2 "
            
        Else
            InsertarLinFactSocContaNueva = False
            Exit Function
        End If
    Else
    ' fin de lo añadido
    
        If vEmpresa.TieneAnalitica Then
            Select Case Tipo
                Case 1, 3, 7, 9
                     Sql = " SELECT 1, variedades.ctaanticipo as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Case 2, 4, 8, 10
                     Sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Case 6 ' siniestros
                     Sql = " SELECT 1, variedades.ctasiniestros as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
            End Select
            If TipoFact = "FTS" Then
                Sql = " SELECT 1, variedades.ctaacarecol as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
            End If
        Else
            Select Case Tipo
                Case 1, 3, 7, 9
                     Sql = " SELECT 1, variedades.ctaanticipo as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                Case 2, 4, 8, 10
                     Sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                Case 6 ' siniestros
                     Sql = " SELECT 1, variedades.ctasiniestros as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
            End Select
            '[Monica]16/07/2014: añadido el caso de tipo transporte tercero de Picassent
            If TipoFact = "FTS" Or TipoFact = "FTT" Then
                Sql = " SELECT 1, variedades.ctaacarecol as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
            End If
        End If
        Sql = Sql & " FROM rfactsoc_variedad, variedades "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "rfactsoc", "rfactsoc_variedad") & " and"
        Sql = Sql & " rfactsoc_variedad.codvarie = variedades.codvarie "
        Sql = Sql & " group by 1,2 "
        Sql = Sql & " order by 1,2 "

    End If

    SqlAux2 = "select rfactsoc.tipoiva from " & cadTabla & " where " & cadWHERE
    vTipoIvaAux = DevuelveValor(SqlAux2)
    
    SqlAux2 = "select rfactsoc.porc_iva from " & cadTabla & " where " & cadWHERE
    vPorIvaAux = DevuelveValor(SqlAux2)
    
    SqlAux2 = "select rfactsoc.imporiva from " & cadTabla & " where " & cadWHERE
    vImpIvaAux = DevuelveValor(SqlAux2)


    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    TotImpIVA = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        ImpLinea = Rs!Importe
        
        ' si se trata de una liquidacion hemos de descontar los anticipos por variedad
        ' que no sean anticipo de gastos
        If (Tipo = 2 Or Tipo = 4 Or Tipo = 8 Or Tipo = 10) And TipoFact <> "FTS" Then
            Sql3 = "select sum(baseimpo) from rfactsoc_anticipos, variedades "
            Sql3 = Sql3 & " where " & Replace(cadWHERE, "rfactsoc", "rfactsoc_anticipos")
            Sql3 = Sql3 & " and rfactsoc_anticipos.codvarieanti = variedades.codvarie "
            Sql3 = Sql3 & " and variedades.ctaliquidacion = " & DBSet(Rs!cuenta, "N")
            
            ImpAnticipo = DevuelveValor(Sql3)
            
            ImpLinea = ImpLinea - ImpAnticipo
        End If
        '----
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        Sql = DBSet(SerieFraPro, "T") & "," & NumRegis & "," & DBSet(FecRecep, "F") & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(Rs!cuenta, "T")
        Sql = Sql & ","
        
        If vEmpresa.TieneAnalitica Then
            If DBLet(Rs!CodCCost, "T") = "----" Then
                Sql = Sql & DBSet(CCoste, "T")
            Else
                Sql = Sql & DBSet(Rs!CodCCost, "T")
                CCoste = DBLet(Rs!CodCCost, "T")
            End If
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        'tipo de iva, porcentaje iva y porcentaje recargo
        Sql = Sql & "," & vTipoIvaAux
        Sql = Sql & "," & vPorIvaAux
        Sql = Sql & "," & ValorNulo
        Sql = Sql & "," & DBSet(ImpLinea, "N")
        
        Sql2 = Sql & "," 'nos guardamos la linea sin el importe iva por si a la última hay q descontarle para q coincida con total factura
        
        impiva = Round(ImpLinea * vPorIvaAux / 100, 2)
        
        TotImpIVA = TotImpIVA + impiva
        
        Sql = Sql & "," & DBSet(impiva, "N") & ","
        
        ' llevan retencion
        Sql = Sql & ValorNulo & ",1"
        
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If TotImpIVA <> vImpIvaAux Then
'        MsgBox "FALTA cuadrar importes de iva!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = vImpIvaAux - TotImpIVA
        totimp = impiva + totimp '(+- diferencia)
        Sql2 = Sql2 & DBSet(totimp, "N") & ","
        Sql2 = Sql2 & ValorNulo & ",1"
        If SQLaux <> "" Then 'hay mas de una linea
            cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If

    ' las retenciones si las hay
    ' las aportaciones de fondo operativo si las hay
    If ImpAport <> 0 Then
'        Sql = NumRegis & "," & AnyoFacPr & "," & i & ","
'        Sql = Sql & DBSet(CtaSocio, "T")
'        Sql = Sql & "," & DBSet(ImpAport, "N") & ","
'        Sql = Sql & ValorNulo ' no llevan centro de coste
'
'        Cad = Cad & "(" & Sql & ")" & ","
'        i = i + 1
    
    
        Sql = DBSet(SerieFraPro, "T") & "," & NumRegis & "," & DBSet(FecRecep, "F") & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(CtaSocio, "T")
        Sql = Sql & ","
        Sql = Sql & ValorNulo ' no llevan centro de coste
        
        'tipo de iva, porcentaje iva y porcentaje recargo
        Sql = Sql & "," & vTipoIvaAux
        Sql = Sql & "," & vPorIvaAux
        Sql = Sql & "," & ValorNulo
        Sql = Sql & "," & DBSet(ImpAport, "N")
        
        impiva = Round(ImpAport * vPorIvaAux / 100, 2)
        
        Sql = Sql & "," & DBSet(impiva, "N") & ","
        
        ' llevan retencion
        Sql = Sql & ValorNulo & ",0"
        
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
    
'*****
    
'        Sql = NumRegis & "," & AnyoFacPr & "," & i & ","
'        Sql = Sql & DBSet(CtaAport, "T")
'        Sql = Sql & "," & DBSet(ImpAport * (-1), "N") & ","
'        Sql = Sql & ValorNulo ' no llevan centro de coste
'
'        Cad = Cad & "(" & Sql & ")" & ","
'        i = i + 1
    
        Sql = DBSet(SerieFraPro, "T") & "," & NumRegis & "," & DBSet(FecRecep, "F") & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(CtaAport, "T")
        Sql = Sql & ","
        Sql = Sql & ValorNulo ' no llevan centro de coste
        
        'tipo de iva, porcentaje iva y porcentaje recargo
        Sql = Sql & "," & vTipoIvaAux
        Sql = Sql & "," & vPorIvaAux
        Sql = Sql & "," & ValorNulo
        Sql = Sql & "," & DBSet(ImpAport * (-1), "N")
        
        impiva = Round(ImpAport * (-1) * vPorIvaAux / 100, 2)
        
        Sql = Sql & "," & DBSet(impiva, "N") & ","
        
        ' llevan retencion
        Sql = Sql & ValorNulo & ",0"
        
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
    
    End If
    
    '[Monica]20/12/2013: si es montifrut no descontamos el descuento que tengo grabado a pie
        '[Monica]09/03/2015: para el caso de Catadau tampoco se tienen que insertar las bases correspondientes a gastos
            '[Monica]13/04/2016: levanto el control de que no se contabilicen los gastos en Catadau
    If vParamAplic.Cooperativa <> 12 Then  'And vParamAplic.Cooperativa <> 0 Then
        ' insertamos todos los gastos a pie de factura rfactsoc_gastos
        Sql = " SELECT rconcepgasto.codmacta as cuenta, sum(rfactsoc_gastos.importe) as importe "
        Sql = Sql & " from rconcepgasto INNER JOIN rfactsoc_gastos ON rconcepgasto.codgasto = rfactsoc_gastos.codgasto "
        Sql = Sql & " where " & Replace(cadWHERE, "rfactsoc", "rfactsoc_gastos")
        
        '[Monica]06/06/2016: si es catadau solo los que tengan cuenta
        If vParamAplic.Cooperativa = 0 Then
            Sql = Sql & " and not rconcepgasto.codmacta is null and rconcepgasto.codmacta <> '' "
        End If
        
        Sql = Sql & " group by 1 "
        Sql = Sql & " order by 1 "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        While Not Rs.EOF
            Sql = DBSet(SerieFraPro, "T") & "," & NumRegis & "," & DBSet(FecRecep, "F") & "," & AnyoFacPr & "," & I & ","
            Sql = Sql & DBSet(CtaSocio, "T") & ","
            
            Sql = Sql & ValorNulo ' no llevan centro de coste
            
            'tipo de iva, porcentaje iva y porcentaje recargo
            Sql = Sql & "," & vTipoIvaAux
            Sql = Sql & "," & vPorIvaAux
            Sql = Sql & "," & ValorNulo
            Sql = Sql & "," & DBSet(Rs!Importe, "N")
            
            impiva = Round(DBLet(Rs!Importe, "N") * vPorIvaAux / 100, 2)
            
            Sql = Sql & "," & DBSet(impiva, "N") & ","
            
            ' llevan retencion
            Sql = Sql & ValorNulo & ",0"
            
            cad = cad & "(" & Sql & ")" & ","
            I = I + 1
            
            
            
            Sql = DBSet(SerieFraPro, "T") & "," & NumRegis & "," & DBSet(FecRecep, "F") & "," & AnyoFacPr & "," & I & ","
            Sql = Sql & DBSet(Rs!cuenta, "T") & ","
            
            Sql = Sql & ValorNulo ' no llevan centro de coste
            
            'tipo de iva, porcentaje iva y porcentaje recargo
            Sql = Sql & "," & vTipoIvaAux
            Sql = Sql & "," & vPorIvaAux
            Sql = Sql & "," & ValorNulo
            Sql = Sql & "," & DBSet(Rs!Importe * (-1), "N")
            
            impiva = Round(DBLet(Rs!Importe, "N") * (-1) * vPorIvaAux / 100, 2)
            
            Sql = Sql & "," & DBSet(impiva, "N") & ","
            
            ' llevan retencion
            Sql = Sql & ValorNulo & ",0"
            
            cad = cad & "(" & Sql & ")" & ","
            I = I + 1
        
            Rs.MoveNext
        Wend
        Set Rs = Nothing
    End If
    
    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO factpro_lineas (numserie,numregis,fecharec,anofactu,numlinea,codmacta,codccost,codigiva,porciva,porcrec,baseimpo,impoiva,imporec,aplicret) "
        
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If
    
    Tipo = TipoAnt

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactSocContaNueva = False
        cadErr = Err.Description
    Else
        InsertarLinFactSocContaNueva = True
    End If
End Function


Private Function ActualizarCabFact(cadTabla As String, cadWHERE As String, cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim Sql As String

    On Error GoTo EActualizar
    
    Select Case cadTabla
        Case "rrecibpozos"
    
            Sql = "UPDATE " & cadTabla & " SET contabilizado=1 "
            
        Case Else
            Sql = "UPDATE " & cadTabla & " SET intconta=1"
            
    End Select
    Sql = Sql & " WHERE " & cadWHERE

    conn.Execute Sql
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCabFact = False
        cadErr = Err.Description
    Else
        ActualizarCabFact = True
    End If
End Function



Private Function ActualizarCabFactSoc(cadTabla As String, cadWHERE As String, cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim Sql As String

    On Error GoTo EActualizar
    
        
    Sql = "UPDATE " & cadTabla & " SET contabilizado=1 "
    Sql = Sql & " WHERE " & cadWHERE

    conn.Execute Sql
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCabFactSoc = False
        cadErr = Err.Description
    Else
        ActualizarCabFactSoc = True
    End If
End Function



'----------------------------------------------------------------------
' FACTURAS SOCIOS
'----------------------------------------------------------------------

Public Function PasarFacturaSoc(cadWHERE As String, CodCCost As String, FechaFin As String, Seccion As String, TipoFact As Byte, FecRecep As Date, FecVto As Date, ForpaPos As String, ForpaNeg As String, CtaBanc As String, CtaRete As String, CtaApor As String, TipoM As String, ByRef vContaFra As cContabilizarFacturas, IvaRea As Integer) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariagro.rfactsoc --> conta.cabfactprov
' ariagro.rfactsoc_variedad --> conta.linfactprov
'Actualizar la tabla ariagro.rfactsoc.contabilizada=1 para indicar que ya esta contabilizada
Dim B As Boolean
Dim cadMen As String
Dim Sql As String
Dim Mc As Contadores
Dim Obs As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
        
    Set Mc = New Contadores
        
    '[Monica]09/11/2016: nueva clase de socio
    Set vSoc = New cSocio
    
    Sql = "select codsocio from rfactsoc where " & cadWHERE
    vSoc.LeerDatos DevuelveValor(Sql)
            
        
        
    '[Monica]09/11/2016: cargamos primero las variables
    '**************************************************
    FecVenci = FecVto
    ForpaPosi = ForpaPos
    ForpaNega = ForpaNeg
    CtaBanco = CtaBanc
    CtaReten = CtaRete
    CtaAport = CtaApor
    tipoMov = TipoM    ' codtipom de la factura de socio
    
    '[Monica]09/05/2013: si la cooperativa es Montifrut, las formas de pago estan en la propia factura
    If vParamAplic.Cooperativa = 12 Then
        ForpaPosi = DevuelveValor("select codforpa from rfactsoc where " & cadWHERE)
        ForpaNega = ForpaPosi
    End If
    '**************************************************
        
        
    '[Monica]29/04/2011: INTERNAS
    If EsFacturaInterna(cadWHERE) Then
        CtaReten = CtaRete
        CtaAport = CtaApor
        ' Insertamos en el diario
        B = InsertarAsientoDiario(cadWHERE, cadMen, Mc, CDate(FechaFin), Seccion, TipoFact, FecRecep, TipoM)
        cadMen = "Insertando Factura en Diario: " & cadMen
    Else
       '---- Insertar en la conta Cabecera Factura
        B = InsertarCabFactSoc(cadWHERE, cadMen, Mc, CDate(FechaFin), Seccion, TipoFact, FecRecep, TipoM, ForpaPos, ForpaNeg, vContaFra, IvaRea)
        cadMen = "Insertando Cab. Factura: " & cadMen
    End If
    
    If B Then
'[Monica]09/11/2016: lo hemos hecho arriba
'        FecVenci = FecVto
'        ForpaPosi = ForpaPos
'        ForpaNega = ForpaNeg
'        CtaBanco = CtaBanc
'        CtaReten = CtaRete
'        CtaAport = CtaApor
'        tipoMov = TipoM    ' codtipom de la factura de socio
'
'        '[Monica]09/05/2013: si la cooperativa es Montifrut, las formas de pago estan en la propia factura
'        If vParamAplic.Cooperativa = 12 Then
'            ForpaPosi = DevuelveValor("select codforpa from rfactsoc where " & cadWHERE)
'            ForpaNega = ForpaPosi
'        End If
        
'01-06-2009
        B = InsertarEnTesoreriaSoc(cadWHERE, cadMen, FacturaSoc, FecFactuSoc)
        cadMen = "Insertando en Tesoreria: " & cadMen

        If B Then
            CCoste = CodCCost
            '[Monica]29/04/2011: INTERNAS
            If Not EsFacturaInterna(cadWHERE) Then
                '---- Insertar lineas de Factura en la Conta
                If vParamAplic.ContabilidadNueva Then
                    B = InsertarLinFactSocContaNueva("rfactsoc", cadWHERE, cadMen, TipoFact, FecRecep, Mc.Contador)
                Else
                    B = InsertarLinFactSoc("rfactsoc", cadWHERE, cadMen, TipoFact, Mc.Contador)
                End If
                cadMen = "Insertando Lin. Factura: " & cadMen
            End If
            
            If B Then
                If Not EsFacturaInterna(cadWHERE) Then
                    If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaProv(vContaFra.NumeroFactura, vContaFra.Anofac)
                End If
            
                '---- Poner intconta=1 en ariges.scafac
                B = ActualizarCabFactSoc("rfactsoc", cadWHERE, cadMen)
                cadMen = "Actualizando Factura Socio: " & cadMen
            End If
        End If
    End If
    
    
EContab:
    Set vSoc = Nothing
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Contabilizando Factura Socio", Err.Description
    End If
    If B Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaSoc = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaSoc = False
        If Not B Then
            InsertarTMPErrFacSoc cadMen, cadWHERE
        End If
    End If
End Function



Private Function InsertarCabFactSoc(cadWHERE As String, cadErr As String, ByRef Mc As Contadores, FechaFin As Date, Secci As String, Tipo As Byte, FecRec As Date, TipoM As String, FPPos As String, FPNeg As String, ByRef vContaFra As cContabilizarFacturas, IvaRea As Integer) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Concepto As String
Dim letraser As String
Dim Socio As String
Dim TipoOpera As Integer
Dim Aux As String

Dim Sql2 As String
Dim CadenaInsertFaclin2 As String


    On Error GoTo EInsertar
    
    '[Monica]09/05/2013: en el caso de Montifrut cuando contabilizo la fecha de recepcion va a ser la fecha de factura
    If vParamAplic.Cooperativa = 12 Then
        Sql = " SELECT codtipom, fecfactu,year(fecfactu) as anofacpr,fecfactu,numfactu,rsocios_seccion.codmacpro,"
    Else
        Sql = " SELECT codtipom, fecfactu,year(" & DBSet(FecRec, "F") & ") as anofacpr," & DBSet(FecRec, "F") & ",numfactu,rsocios_seccion.codmacpro,"
    End If
    
    Sql = Sql & "baseimpo, tipoiva, porc_iva,imporiva,basereten,porc_ret,impreten,baseaport,porc_apo,impapor,totalfac,"
    Sql = Sql & "rectif_codtipom, rectif_numfactu, rectif_fecfactu,"
    Sql = Sql & "rsocios.codsocio, rsocios.nomsocio, rsocios.codbanco, rsocios.codsucur, rsocios.digcontr, rsocios.cuentaba, rsocios.iban "
    
    '[Monica]27/01/2012: Si han introducido el nro de factura recibido es el que hay que llevar a conta
    Sql = Sql & ", rfactsoc.numfacrec "
    
    Sql = Sql & ", rsocios.dirsocio, rsocios.pobsocio, rsocios.codpostal, rsocios.prosocio, rsocios.nifsocio "
    '[Monica]02/05/2017: tipoirpf
    Sql = Sql & ", rfactsoc.tipoirpf "
    
    
    Sql = Sql & " FROM (" & "rfactsoc "
    Sql = Sql & "INNER JOIN rsocios ON rfactsoc.codsocio=rsocios.codsocio) "
    Sql = Sql & " INNER JOIN rsocios_seccion ON rfactsoc.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Secci, "N")
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
        '[Monica]09/05/2013: si la cooperativa es Montifrut la fecha de recepcion es la de factura
        If vParamAplic.Cooperativa = 12 Then
            FecRec = DBLet(Rs!fecfactu, "F")
            
            If DBLet(Rs!CodTipom, "T") = "FRS" Then
                Mc.Contador = (CInt(Mid(Year(FecRec), 3, 2) & "1") * 100000) + DBLet(Rs!numfactu, "N")
            Else
                '[Monica]13/05/2013: nro de registro introducido + nro de factura
                Mc.Contador = (CInt(Mid(Year(FecRec), 3, 2)) * 1000000) + DBLet(Rs!numfactu, "N")
            End If
            
            vContaFra.NumeroFactura = Mc.Contador
            vContaFra.Anofac = Year(FecRec)
            
            
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            BaseImp = DBLet(Rs!baseimpo, "N")
            TotalFac = BaseImp + DBLet(Rs!ImporIva, "N")
            AnyoFacPr = Rs!anofacpr
            
            ImpReten = DBLet(Rs!ImpReten, "N")
            ImpAport = DBLet(Rs!impapor, "N")
            
            '[Monica]27/01/2012:Si han introducido el nro de factura recibido es el que hay que llevar a conta
            If DBLet(Rs!numfacrec, "T") <> "" Then
                FacturaSoc = DBLet(Rs!numfacrec, "T")
            Else
                letraser = ""
                letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(TipoM, "T"))
            
                FacturaSoc = letraser & "-" & DBLet(Rs!numfactu, "N")
            End If
            
            FecFactuSoc = DBLet(Rs!fecfactu, "F")
            
            CodTipomRECT = DBLet(Rs!rectif_codtipom, "T")
            NumfactuRECT = DBLet(Rs!rectif_numfactu, "T")
            FecfactuRECT = DBLet(Rs!rectif_fecfactu, "T")
            
            CtaSocio = Rs!codmacpro
            Seccion = Secci
            TipoFact = Tipo
            FecRecep = FecRec
            BancoSoc = DBLet(Rs!CodBanco, "N")
            SucurSoc = DBLet(Rs!CodSucur, "N")
            DigcoSoc = DBLet(Rs!digcontr, "T")
            CtaBaSoc = DBLet(Rs!CuentaBa, "T")
            IbanSoc = DBLet(Rs!Iban, "T")
            TotalTesor = DBLet(Rs!TotalFac, "N")
            
            
            Variedades = VariedadesFactura(cadWHERE)
            
            Select Case TipoFact
                Case 1, 7, 9 ' anticipo
                    Concepto = "ANTICIPO SOCIO"
                Case 2, 8, 10 ' liquidacion
                    Concepto = "LIQUIDACION SOCIO"
                Case 6
                    Concepto = "SINIESTRO"
                Case 11
                    Concepto = "Rectificativa"
                Case Else
                    Concepto = ""
            End Select
            
            Sql = ""
            
            If vParamAplic.ContabilidadNueva Then Sql = "'" & SerieFraPro & "',"
            Sql = Sql & Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & AnyoFacPr & "," & DBSet(FecRecep, "F") & "," & DBSet(FecRecep, "F") & "," & DBSet(FacturaSoc, "T") & "," & DBSet(CtaSocio, "T") & "," & DBSet(Concepto, "T") & ","
            
            If Not vParamAplic.ContabilidadNueva Then
                Sql = Sql & DBSet(BaseImp, "N") & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & DBSet(Rs!porc_iva, "N") & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!ImporIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & DBSet(TotalFac, "N") & "," & DBSet(Rs!TipoIVA, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                cad = cad & "(" & Sql & ")"
            
                'Insertar en la contabilidad
                Sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,fecliqpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                Sql = Sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                Sql = Sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,nodeducible) "
                Sql = Sql & " VALUES " & cad
                ConnConta.Execute Sql
            
            Else
                Sql = Sql & DBSet(Rs!nomsocio, "T") & "," & DBSet(Rs!dirsocio, "T", "S") & ","
                Sql = Sql & DBSet(Rs!codpostal, "T", "S") & "," & DBSet(Rs!pobsocio, "T", "S") & "," & DBSet(Rs!prosocio, "T", "S") & ","
                Sql = Sql & DBSet(Rs!nifSocio, "F", "S") & ",'ES',"
                
                If DBLet(Rs!TotalFac) < 0 Then
                    Sql = Sql & DBSet(FPNeg, "N") & ","
                Else
                    Sql = Sql & DBSet(FPPos, "N") & ","
                End If
                
                '$$$
                '[Monica]02/05/2017: Solo en el caso de que el tipo de iva sea REA
                If DBLet(Rs!TipoIVA, "N") = IvaRea Then
                    TipoOpera = 5 ' REA
                    
                    '[Monica]21/04/2017: antes tenia un 0 en Aux
                    Aux = "X"
                    'codopera,codconce340,codintra
                    Sql = Sql & TipoOpera & "," & DBSet(Aux, "T") & "," & ValorNulo & ","
                    
                Else
                    TipoOpera = 0 ' general
                    
                    '[Monica]21/04/2017: antes tenia un 0 en Aux
                    Aux = "0"
                    'codopera,codconce340,codintra
                    Sql = Sql & TipoOpera & "," & DBSet(Aux, "T") & "," & ValorNulo & ","
                
                End If
                
                '[Monica]10/11/2016: en totalfac llevabamos base + impiva pq antes retencion estaba en lineas
                '                    en la nueva conta está en la cabecera
                TotalFac = TotalFac - ImpReten
                
                'para las lineas
                'factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
                'IVA 1, siempre existe
                Aux = "'" & SerieFraPro & "'," & Mc.Contador & "," & DBSet(Rs!FecRecep, "F") & "," & Rs!anofacpr & ","
                
                Sql2 = Aux & "1," & DBSet(BaseImp, "N") & "," & DBSet(Rs!TipoIVA, "N") & "," & DBSet(Rs!porc_iva, "N") & ","
                Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!ImporIva, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
                    
                'Los totales
                'totbases,totbasesret,totivas,totrecargo,totfacpr,
                Sql = Sql & DBSet(BaseImp, "N") & "," & DBSet(Rs!BaseReten, "N", "S") & ","
                'totivas
                Sql = Sql & DBSet(Rs!ImporIva, "N") & "," & DBSet(TotalFac, "N") & ","
                If DBLet(Rs!porc_ret, "N") <> 0 Then
                    Sql = Sql & DBSet(Rs!porc_ret, "N") & "," & DBSet(Rs!ImpReten, "N") & "," & DBSet(vParamAplic.CtaRetenSoc, "T") & ","
                    
                    '[Monica]03/05/2017: si es una factura de transporte de socio (idem a una fra de transportista)
'               si retencion : Si REA + modulos ----> tipo retencion = 2 (act.agricola)
'                              Si no REA + modulos--> tipo retencion = 1 (act.profesional)
'                              si E.D.  ------------> tipo retencion = 4 (act.empresarial)
                    If DBLet(Rs!CodTipom, "T") = "FTS" Then
                        '[Monica]03/05/2017: tipo de retencion
                        If Rs!TipoIVA = IvaRea And Rs!TipoIRPF = 0 Then Sql = Sql & "2"
                        If Rs!TipoIVA <> IvaRea And Rs!TipoIRPF = 0 Then Sql = Sql & "1"
                        If Rs!TipoIRPF = 1 Then Sql = Sql & "4"
                    Else
                        '[Monica]03/05/2017: dependiendo del tipo de irpf
'                si retencion : Si modulos --> tipo retencion = 2 (act.agricola)
'                               si E.D. -----> tipo retencion = 4 (act.empresarial)
'                               si Entidad --> tipo retencion = 0 (sin retencion)
                        Select Case Rs!TipoIRPF
                            Case 0
                                Sql = Sql & "2" ' si modulos entonces act.agricola
                            Case 1
                                Sql = Sql & "4" ' si e.d entonces act.empresarial
                            Case 2
                                Sql = Sql & "0" ' si entidad --> nada
                        End Select
                    End If
                Else
                    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                End If
                
                cad = cad & "(" & Sql & ")"
            
                'Insertar en la contabilidad
                Sql = "INSERT INTO factpro(numserie,numregis,fecfactu,anofactu,fecharec,fecliqpr,numfactu,codmacta,observa,nommacta,"
                Sql = Sql & "dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,codforpa,codopera,codconce340,codintra,"
                Sql = Sql & "totbases,totbasesret,totivas,totfacpr,retfacpr,trefacpr,cuereten,tiporeten)"
                Sql = Sql & " VALUES " & cad
                ConnConta.Execute Sql
            
                'Las  lineas de IVA
                Sql = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)"
                Sql = Sql & " VALUES " & CadenaInsertFaclin2
                ConnConta.Execute Sql
            
            End If
            
            
            'añadido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            Sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            Sql = Sql & ",'" & DevNombreSQL(FacturaSoc) & " @ " & Format(Rs!fecfactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!nomsocio) & "'," & Rs!Codsocio & ")"
            conn.Execute Sql

            FacturaSoc = DBLet(Rs!numfactu, "N")
            
        Else
        
            If Mc.ConseguirContador("1", (FecRec <= CDate(FechaFin) - 365), True) = 0 Then
                vContaFra.NumeroFactura = Mc.Contador
                vContaFra.Anofac = Year(FecRec)
            
                'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
                BaseImp = DBLet(Rs!baseimpo, "N")
                TotalFac = BaseImp + DBLet(Rs!ImporIva, "N")
                AnyoFacPr = Rs!anofacpr
                
                ImpReten = DBLet(Rs!ImpReten, "N")
                ImpAport = DBLet(Rs!impapor, "N")
                
                '[Monica]27/01/2012:Si han introducido el nro de factura recibido es el que hay que llevar a conta
                If DBLet(Rs!numfacrec, "T") <> "" Then
                    FacturaSoc = DBLet(Rs!numfacrec, "T")
                Else
                    letraser = ""
                    letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(TipoM, "T"))
                
                    FacturaSoc = letraser & "-" & DBLet(Rs!numfactu, "N")
                End If
                FecFactuSoc = DBLet(Rs!fecfactu, "F")
                
                CodTipomRECT = DBLet(Rs!rectif_codtipom, "T")
                NumfactuRECT = DBLet(Rs!rectif_numfactu, "T")
                FecfactuRECT = DBLet(Rs!rectif_fecfactu, "T")
                
                CtaSocio = Rs!codmacpro
                Seccion = Secci
                TipoFact = Tipo
                FecRecep = FecRec
                IbanSoc = DBLet(Rs!Iban, "T")
                BancoSoc = DBLet(Rs!CodBanco, "N")
                SucurSoc = DBLet(Rs!CodSucur, "N")
                DigcoSoc = DBLet(Rs!digcontr, "T")
                CtaBaSoc = DBLet(Rs!CuentaBa, "T")
                TotalTesor = DBLet(Rs!TotalFac, "N")
                
                '[Monica]08/04/2015: en el caso de catadau vemos si el socio es un asociado para reemplazar la cuenta
                If vParamAplic.Cooperativa = 0 Then
                   Sql = "select rsocios.tiporelacion from rsocios where codsocio in (select codsocio from rfactsoc where " & cadWHERE & ")"
                   If DevuelveValor(Sql) = 1 Then
                       
                       Sql = "select codsocio from rfactsoc where " & cadWHERE
                       Socio = DevuelveValor(Sql)
                       
                       Sql = "select replace(codmacpro,mid(codmacpro,1,length(rseccion.raiz_cliente_socio)), rseccion.raiz_cliente_asociado) "
                       Sql = Sql & " from (rsocios_seccion inner join rseccion on rsocios_seccion.codsecci = rseccion.codsecci) inner join rsocios on rsocios_seccion .codsocio = rsocios.codsocio "
                       Sql = Sql & " where rsocios_seccion.codsocio = " & DBSet(Socio, "N")
                       Sql = Sql & " and rseccion.codsecci = " & DBSet(Seccion, "N")
    
                       CtaSocio = DevuelveValor(Sql)
                   End If
                End If
                
                
                Variedades = VariedadesFactura(cadWHERE)
                
                Select Case TipoFact
                    Case 1, 7, 9 ' anticipo
                        Concepto = "ANTICIPO SOCIO"
                    Case 2, 8, 10 ' liquidacion
                        Concepto = "LIQUIDACION SOCIO"
                    Case 6
                        Concepto = "SINIESTRO"
                    Case 11
                        Concepto = "Rectificativa"
                    Case Else
                        Concepto = ""
                End Select
                    
                    
                Sql = ""
                If vParamAplic.ContabilidadNueva Then Sql = "'" & SerieFraPro & "',"
                Sql = Sql & Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & AnyoFacPr & "," & DBSet(FecRecep, "F") & "," & DBSet(FecRecep, "F") & "," & DBSet(FacturaSoc, "T") & "," & DBSet(CtaSocio, "T") & "," & DBSet(Concepto, "T") & ","
                
                
                If Not vParamAplic.ContabilidadNueva Then
                    Sql = Sql & DBSet(BaseImp, "N") & "," & ValorNulo & "," & ValorNulo & ","
                    Sql = Sql & DBSet(Rs!porc_iva, "N") & "," & ValorNulo & "," & ValorNulo & ","
                    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!ImporIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
                    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                    Sql = Sql & DBSet(TotalFac, "N") & "," & DBSet(Rs!TipoIVA, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
                    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                    cad = cad & "(" & Sql & ")"
                
                    'Insertar en la contabilidad
                    Sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,fecliqpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                    Sql = Sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                    Sql = Sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,nodeducible) "
                    Sql = Sql & " VALUES " & cad
                    ConnConta.Execute Sql
                Else
                    Sql = Sql & DBSet(Rs!nomsocio, "T") & "," & DBSet(Rs!dirsocio, "T", "S") & ","
                    Sql = Sql & DBSet(Rs!codpostal, "T", "S") & "," & DBSet(Rs!pobsocio, "T", "S") & "," & DBSet(Rs!prosocio, "T", "S") & ","
                    Sql = Sql & DBSet(Rs!nifSocio, "F", "S") & ",'ES',"
                    If DBLet(Rs!TotalFac) < 0 Then
                        Sql = Sql & DBSet(FPNeg, "N") & ","
                    Else
                        Sql = Sql & DBSet(FPPos, "N") & ","
                    End If
                
                    '$$$
                    '[Monica]02/05/2017: Solo en el caso de modulos
                    If DBLet(Rs!TipoIVA, "N") = IvaRea Then
                        TipoOpera = 5 ' REA
                        
                        '[Monica]21/04/2017: antes tenia un 0 en Aux
                        Aux = "X"
'                        If Rs!TotalFac < 0 Then Aux = "D"
                        'codopera,codconce340,codintra
                        Sql = Sql & TipoOpera & "," & DBSet(Aux, "T") & "," & ValorNulo & ","
                    Else
                    
                        TipoOpera = 0 ' general
                        
                        '[Monica]21/04/2017: antes tenia un 0 en Aux
                        Aux = "0"
'                        If Rs!TotalFac < 0 Then Aux = "D"
                        'codopera,codconce340,codintra
                        Sql = Sql & TipoOpera & "," & DBSet(Aux, "T") & "," & ValorNulo & ","
                    End If
                    
                    '[Monica]10/11/2016: en totalfac llevabamos base + impiva pq antes retencion estaba en lineas
                    '                    en la nueva conta está en la cabecera
                    TotalFac = TotalFac - ImpReten
                    
                    'para las lineas
                    'factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
                    'IVA 1, siempre existe
                    Aux = "'" & SerieFraPro & "'," & Mc.Contador & "," & DBSet(FecRecep, "F") & "," & Rs!anofacpr & ","
                    
                    Sql2 = Aux & "1," & DBSet(BaseImp, "N") & "," & DBSet(Rs!TipoIVA, "N") & "," & DBSet(Rs!porc_iva, "N") & ","
                    Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!ImporIva, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
                        
                    'Los totales
                    'totbases,totbasesret,totivas,totrecargo,totfacpr,
                    Sql = Sql & DBSet(BaseImp, "N") & "," & DBSet(Rs!BaseReten, "N", "S") & ","
                    'totivas
                    Sql = Sql & DBSet(Rs!ImporIva, "N") & "," & DBSet(TotalFac, "N") & ","
                    If DBLet(Rs!porc_ret, "N") <> 0 Then
                        Sql = Sql & DBSet(Rs!porc_ret, "N") & "," & DBSet(Rs!ImpReten, "N") & "," & DBSet(vParamAplic.CtaRetenSoc, "T") & ","
                        
                        '[Monica]03/05/2017: dependiendo del tipo de irpf
                        Select Case Rs!TipoIRPF
                            Case 0
                                Sql = Sql & "2" ' si modulos entonces act.agricola
                            Case 1
                                Sql = Sql & "4" ' si e.d entonces act.empresarial
                            Case 2
                                Sql = Sql & "0" ' si entidad --> nada
                        End Select
                    Else
                        Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                    End If
                    cad = cad & "(" & Sql & ")"
                
                    'Insertar en la contabilidad
                    Sql = "INSERT INTO factpro(numserie,numregis,fecfactu,anofactu,fecharec,fecliqpr,numfactu,codmacta,observa,nommacta,"
                    Sql = Sql & "dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,codforpa,codopera,codconce340,codintra,"
                    Sql = Sql & "totbases,totbasesret,totivas,totfacpr,retfacpr,trefacpr,cuereten,tiporeten)"
                    Sql = Sql & " VALUES " & cad
                    ConnConta.Execute Sql
                
                    'Las  lineas de IVA
                    Sql = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)"
                    Sql = Sql & " VALUES " & CadenaInsertFaclin2
                    ConnConta.Execute Sql
                        
                        
                End If
                
                
                'añadido como david para saber que numero de registro corresponde a cada factura
                'Para saber el numreo de registro que le asigna a la factrua
                Sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
                Sql = Sql & ",'" & DevNombreSQL(FacturaSoc) & " @ " & Format(Rs!fecfactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!nomsocio) & "'," & Rs!Codsocio & ")"
                conn.Execute Sql
    
                FacturaSoc = DBLet(Rs!numfactu, "N")
            End If
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactSoc = False
        cadErr = Err.Description
    Else
        InsertarCabFactSoc = True
    End If
End Function



Private Function InsertarAsientoDiario(cadWHERE As String, cadErr As String, ByRef Mc As Contadores, FechaFin As Date, Secci As String, Tipo As Byte, FecRec As Date, TipoM As String) As Boolean
' la contabilizacion de las facturas internas de horto, se han de insertar en el diario no en el registro de iva de proveedor
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Concepto As String
Dim letraser As String
Dim Obs As String
Dim cadMen As String
Dim B As Boolean
'Dim CtaSocio As String


    On Error GoTo EInsertar
       
    Sql = " SELECT fecfactu,year(" & DBSet(FecRec, "F") & ") as anofacpr," & DBSet(FecRec, "F") & ",numfactu,rsocios_seccion.codmacpro,"
    Sql = Sql & "baseimpo, tipoiva, porc_iva,imporiva,basereten,porc_ret,impreten,baseaport,porc_apo,impapor,totalfac,"
    Sql = Sql & "rectif_codtipom, rectif_numfactu, rectif_fecfactu,"
    Sql = Sql & "rsocios.codsocio, rsocios.nomsocio, rsocios.codbanco, rsocios.codsucur, rsocios.digcontr, rsocios.cuentaba "
    Sql = Sql & " FROM (" & "rfactsoc "
    Sql = Sql & "INNER JOIN rsocios ON rfactsoc.codsocio=rsocios.codsocio) "
    Sql = Sql & " INNER JOIN rsocios_seccion ON rfactsoc.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Secci, "N")
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
        '[Monica]17/02/2017: hay que coger el nro de asiento antes : Mc.ConseguirContador("1", (FecRec <= CDate(FechaFin) - 365), True) = 0 Then
        If Mc.ConseguirContador("0", (FecRec <= CDate(FechaFin) - 365), True) = 0 Then
        
            BaseImp = DBLet(Rs!baseimpo, "N")
            TotalFac = BaseImp + DBLet(Rs!ImporIva, "N")
            AnyoFacPr = Rs!anofacpr
            
            ImpReten = DBLet(Rs!ImpReten, "N")
            ImpAport = DBLet(Rs!impapor, "N")
            
            letraser = ""
            letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(TipoM, "T"))
            
            FacturaSoc = letraser & "-" & DBLet(Rs!numfactu, "N")
            FecFactuSoc = DBLet(Rs!fecfactu, "F")
            
            CodTipomRECT = DBLet(Rs!rectif_codtipom, "T")
            NumfactuRECT = DBLet(Rs!rectif_numfactu, "T")
            FecfactuRECT = DBLet(Rs!rectif_fecfactu, "T")
            
            CtaSocio = Rs!codmacpro
            Seccion = Secci
            TipoFact = Tipo
            FecRecep = FecRec
            BancoSoc = DBLet(Rs!CodBanco, "N")
            SucurSoc = DBLet(Rs!CodSucur, "N")
            DigcoSoc = DBLet(Rs!digcontr, "T")
            CtaBaSoc = DBLet(Rs!CuentaBa, "T")
            TotalTesor = DBLet(Rs!TotalFac, "N")
            
            '[Monica]08/04/2015: en el caso de catadau vemos si el socio es un asociado para reemplazar la cuenta
            If vParamAplic.Cooperativa = 0 Then
               Sql = "select rsocios.tiporelacion from rsocios where codsocio in (select codsocio from rfactsoc where " & cadWHERE & ")"
               If DevuelveValor(Sql) = 1 Then
                   
                   Sql = "select codsocio from rfactsoc where " & cadWHERE
                   Socio = DevuelveValor(Sql)
                   
                   Sql = "select replace(codmacpro,mid(codmacpro,1,length(rseccion.raiz_cliente_socio)), rseccion.raiz_cliente_asociado) "
                   Sql = Sql & " from (rsocios_seccion inner join rseccion on rsocios_seccion.codsecci = rseccion.codsecci) inner join rsocios on rsocios_seccion .codsocio = rsocios.codsocio "
                   Sql = Sql & " where rsocios_seccion.codsocio = " & DBSet(Socio, "N")
                   Sql = Sql & " and rseccion.codsecci = " & DBSet(Seccion, "N")

                   CtaSocio = DevuelveValor(Sql)
               End If
            End If
            
            Variedades = VariedadesFactura(cadWHERE)
            
            Obs = "Contabilización Factura Interna de Fecha " & Format(FecRecep, "dd/mm/yyyy")
        
            'Insertar en la conta Cabecera Asiento
            B = InsertarCabAsientoDia(vEmpresa.NumDiarioInt, Mc.Contador, CStr(Format(FecRecep, "dd/mm/yyyy")), Obs, cadMen)
            cadMen = "Insertando Cab. Asiento: " & cadMen
    
            If B Then
                Socio = DevuelveValor("select codsocio from rfactsoc where " & cadWHERE)
'                CtaSocio = DevuelveValor("select codmacpro from rsocios_seccion where codsocio = " & Socio & " and codsecci = " & vParamAplic.SeccionHorto)
            
                B = InsertarLinAsientoFactIntProv("rfactsoc", cadWHERE, cadMen, Tipo, CtaSocio, Mc.Contador)
                cadMen = "Insertando Lin. Factura Interna: " & cadMen
            
                Set Mc = Nothing
            End If
            
            FacturaSoc = DBLet(Rs!numfactu, "N")
            
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarAsientoDiario = False
        cadErr = Err.Description
    Else
        InsertarAsientoDiario = True
    End If
End Function



Private Function InsertarLinAsientoFactIntProv(cadTabla As String, cadWHERE As String, cadErr As String, Tipo As Byte, CtaSocio As String, Optional Contador As Long) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim numdocum As String
Dim Ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim I As Long
Dim B As Boolean
Dim cad As String
Dim cadMen As String
Dim FeFact As Date

Dim cadCampo As String

Dim vSocio As cSocio
Dim Socio As String
Dim TipoAnt As Byte
Dim TipoFact As String

Dim totimp As Currency
Dim SQLaux As String
Dim ImpLinea As String
Dim Sql3 As String
Dim ImpAnticipo As Currency
Dim NumFact As Long

    On Error GoTo EInLinea
    
    InsertarLinAsientoFactIntProv = False
    
    TipoAnt = Tipo
'    TipoFactAnt = TipoFact
    
    If Tipo = 11 Then ' si es una factura rectificativa cojo el tipo de movimiento de la factura que rectifico
        Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(CodTipomRECT, "T"))
        
        TipoFact = CodTipomRECT

    Else
        ' [Monica] 29/12/2009 si es liquidacion de industria miramos que cuenta coger dependiendo de que el socio sea
        ' tercero o no lo sea
        Sql = "select mid(rfactsoc.codtipom,1,3) from " & cadTabla & " where " & cadWHERE
        TipoFact = DevuelveValor(Sql)
    
    End If
    
    FeFact = FecFactuSoc
    NumFact = DevuelveValor("select numfactu from rfactsoc where " & cadWHERE)
    
    If Tipo = 2 And TipoFact = "FLI" Then
        Sql = "select rfactsoc.codsocio from " & cadTabla & " where " & cadWHERE
        Socio = DevuelveValor(Sql)
        
        Set vSocio = New cSocio
        If vSocio.LeerDatos(Socio) Then
            If vEmpresa.TieneAnalitica Then
                If vSocio.TipoProd = 1 Then
                    Sql = " SELECT 1, variedades.ctacomtercero as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Else
                    Sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                End If
            Else
                If vSocio.TipoProd = 1 Then
                    Sql = " SELECT 1, variedades.ctacomtercero as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                Else
                    Sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                End If
            End If
            Sql = Sql & " FROM rfactsoc_variedad, variedades "
            Sql = Sql & " WHERE " & Replace(cadWHERE, "rfactsoc", "rfactsoc_variedad") & " and"
            Sql = Sql & " rfactsoc_variedad.codvarie = variedades.codvarie "
            Sql = Sql & " group by 1,2 "
            Sql = Sql & " order by 1,2 "
            
        Else
            InsertarLinAsientoFactIntProv = False
            Exit Function
        End If
    Else
    ' fin de lo añadido
    
        If vEmpresa.TieneAnalitica Then
            Select Case Tipo
                Case 1, 3, 7, 9
                     Sql = " SELECT 1, variedades.ctaanticipo as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Case 2, 4, 8, 10
                     Sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Case 6 ' siniestros
                     Sql = " SELECT 1, variedades.ctasiniestros as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
            End Select
            If TipoFact = "FTS" Then
                Sql = " SELECT 1, variedades.ctaacarecol as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
            End If
        Else
            Select Case Tipo
                Case 1, 3, 7, 9
                     Sql = " SELECT 1, variedades.ctaanticipo as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                Case 2, 4, 8, 10
                     Sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                Case 6 ' siniestros
                     Sql = " SELECT 1, variedades.ctasiniestros as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
            End Select
            If TipoFact = "FTS" Then
                Sql = " SELECT 1, variedades.ctaacarecol as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
            End If
        End If
        Sql = Sql & " FROM rfactsoc_variedad, variedades "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "rfactsoc", "rfactsoc_variedad") & " and"
        Sql = Sql & " rfactsoc_variedad.codvarie = variedades.codvarie "
        Sql = Sql & " group by 1,2 "
        Sql = Sql & " order by 1,2 "

    End If

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    I = 0
    ImporteD = 0
    ImporteH = 0
    
    numdocum = Format(NumFact, "0000000")
    Ampliacion = FacturaSoc  'TipoFact & "-" & Format(NumFact, "0000000")
    ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoIntProv, "N")) & " " & Ampliacion
    ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoIntProv, "N")) & " " & Ampliacion
    
    If Not Rs.EOF Then Rs.MoveFirst
    
    B = True

    cad = ""
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        ImpLinea = Rs!Importe
        
        ' si se trata de una liquidacion hemos de descontar los anticipos por variedad
        ' que no sean anticipo de gastos
        If (Tipo = 2 Or Tipo = 4 Or Tipo = 8 Or Tipo = 10) And TipoFact <> "FTS" Then
            Sql3 = "select sum(baseimpo) from rfactsoc_anticipos, variedades "
            Sql3 = Sql3 & " where " & Replace(cadWHERE, "rfactsoc", "rfactsoc_anticipos")
            Sql3 = Sql3 & " and rfactsoc_anticipos.codvarieanti = variedades.codvarie "
            Sql3 = Sql3 & " and variedades.ctaliquidacion = " & DBSet(Rs!cuenta, "N")
            
            ImpAnticipo = DevuelveValor(Sql3)
            
            ImpLinea = ImpLinea - ImpAnticipo
        End If
        '----
        totimp = totimp + ImpLinea
        
        I = I + 1
        
        cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
        cad = cad & DBSet(I, "N") & "," & DBSet(Rs!cuenta, "T") & "," & DBSet(numdocum, "T") & ","
        
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        '[Monica]16/06/2016: antes RS.Fields(2).Value
        If ImpLinea > 0 Then
            ' importe al debe en positivo                                                       '[Monica]16/06/2016: antes RS.Fields(2).Value
            cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(ImpLinea, "N") & ","
            cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
        
            ImporteD = ImporteD + (CCur(ImpLinea)) '[Monica]16/06/2016: antes RS.Fields(2).Value
        Else
            ' importe al haber en positivo, cambiamos el signo
            cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
            '[Monica]16/06/2016: antes RS.Fields(2).Value
            cad = cad & DBSet((ImpLinea) * (-1), "N") & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
        
            ImporteH = ImporteH + (CCur(ImpLinea) * (-1)) '[Monica]16/06/2016: antes RS.Fields(2).Value
        End If
        
        cad = "(" & cad & ")"
        
        B = InsertarLinAsientoDia(cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & I

        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)

        If ImpLinea > 0 Then
            If vParamAplic.ContabilidadNueva Then
                Sql = "update hlinapu set timporteD = " & DBSet(totimp, "N")
            Else
                Sql = "update linapu set timporteD = " & DBSet(totimp, "N")
            End If
            Sql = Sql & " where numdiari = " & DBSet(vEmpresa.NumDiarioInt, "N")
            Sql = Sql & " and fechaent = " & DBSet(FecRecep, "F")
            Sql = Sql & " and numasien = " & DBSet(Contador, "N")
            Sql = Sql & " and linliapu = " & DBSet(I, "N")
            
            ConnConta.Execute Sql
        Else
            If vParamAplic.ContabilidadNueva Then
                Sql = "update hlinapu set timporteH = " & DBSet(totimp, "N")
            Else
                Sql = "update linapu set timporteH = " & DBSet(totimp, "N")
            End If
            Sql = Sql & " where numdiari = " & DBSet(vEmpresa.NumDiarioInt, "N")
            Sql = Sql & " and fechaent = " & DBSet(FecRecep, "F")
            Sql = Sql & " and numasien = " & DBSet(Contador, "N")
            Sql = Sql & " and linliapu = " & DBSet(I, "N")
            
            ConnConta.Execute Sql
        End If
    End If

    If B And I > 0 Then
        I = I + 1
        
        ' el Total es sobre la cuenta del socio
        cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
        cad = cad & DBSet(I, "N") & ","
        cad = cad & DBSet(CtaSocio, "T") & "," & DBSet(numdocum, "T") & ","
            
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If ImporteD - ImporteH < 0 Then
            ' importe al debe en positivo
            cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet((ImporteD - ImporteH) * (-1), "N") & ","
            cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
        Else
            ' importe al haber en positivo, cambiamos el signo
            cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
            cad = cad & DBSet((ImporteD - ImporteH), "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
        End If
        
        cad = "(" & cad & ")"
        
        B = InsertarLinAsientoDia(cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & I
        
    End If

    If B Then
        ' las retenciones si las hay
        If ImpReten <> 0 Then
            I = I + 1
            
            cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
            cad = cad & DBSet(I, "N") & "," & DBSet(CtaSocio, "T") & "," & DBSet(numdocum, "T") & ","
        
            If ImpReten > 0 Then
                ' importe al debe en positivo
                cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(ImpReten, "N") & ","
                cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaReten, "T") & "," & ValorNulo & ",0"
            Else
                ' importe al haber en positivo, cambiamos el signo
                cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                cad = cad & DBSet((ImpReten * (-1)), "N") & "," & ValorNulo & "," & DBSet(CtaReten, "T") & "," & ValorNulo & ",0"
            
            End If
            
            cad = "(" & cad & ")"
            
            B = InsertarLinAsientoDia(cad, cadMen)
            cadMen = "Insertando Lin. Asiento: " & I
            
            If B Then
                I = I + 1
                
                cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
                cad = cad & DBSet(I, "N") & "," & DBSet(CtaReten, "T") & "," & DBSet(numdocum, "T") & ","
                If ImpReten > 0 Then
                    ' importe al haber en positivo
                    cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet(ImpReten, "N") & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & DBSet((ImpReten * (-1)), "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
                End If
                
                cad = "(" & cad & ")"
                
                B = InsertarLinAsientoDia(cad, cadMen)
                cadMen = "Insertando Lin. Asiento: " & I
            End If
            
        End If
    End If
    
    
    If B Then
        ' las aportaciones de fondo operativo si las hay
        If ImpAport <> 0 Then
            I = I + 1
            
            cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
            cad = cad & DBSet(I, "N") & "," & DBSet(CtaSocio, "T") & "," & DBSet(numdocum, "T") & ","
        
            If ImpAport > 0 Then
                ' importe al debe en positivo
                cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(ImpAport, "N") & ","
                cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaAport, "T") & "," & ValorNulo & ",0"
            Else
                ' importe al haber en positivo, cambiamos el signo
                cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                cad = cad & DBSet((ImpAport * (-1)), "N") & "," & ValorNulo & "," & DBSet(CtaAport, "T") & "," & ValorNulo & ",0"
            
            End If
            
            cad = "(" & cad & ")"
            
            B = InsertarLinAsientoDia(cad, cadMen)
            cadMen = "Insertando Lin. Asiento: " & I
            
            If B Then
                I = I + 1
                
                cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
                cad = cad & DBSet(I, "N") & "," & DBSet(CtaAport, "T") & "," & DBSet(numdocum, "T") & ","
                If ImpAport > 0 Then
                    ' importe al haber en positivo
                    cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet(ImpAport, "N") & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & DBSet((ImpAport * (-1)), "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
                End If
                
                cad = "(" & cad & ")"
                
                B = InsertarLinAsientoDia(cad, cadMen)
                cadMen = "Insertando Lin. Asiento: " & I
            End If
        End If
    End If
    
    '[Monica]09/03/2015: para el caso de Catadau no hay apuntes de gastos, añadida la condicion de cooperativa
    If B And vParamAplic.Cooperativa <> 0 Then
        ' insertamos todos los gastos a pie de factura rfactsoc_gastos
        Sql = " SELECT rconcepgasto.codmacta as cuenta, sum(rfactsoc_gastos.importe) as importe "
        Sql = Sql & " from rconcepgasto INNER JOIN rfactsoc_gastos ON rconcepgasto.codgasto = rfactsoc_gastos.codgasto "
        Sql = Sql & " where " & Replace(cadWHERE, "rfactsoc", "rfactsoc_gastos")
        Sql = Sql & " group by 1 "
        Sql = Sql & " order by 1 "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        While Not Rs.EOF And B
            I = I + 1
            
            cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
            cad = cad & DBSet(I, "N") & "," & DBSet(CtaSocio, "T") & "," & DBSet(numdocum, "T") & ","
        
            If Rs!Importe > 0 Then
                ' importe al debe en positivo
                cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Rs!Importe, "N") & ","
                cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!cuenta, "T") & "," & ValorNulo & ",0"
            Else
                ' importe al haber en positivo, cambiamos el signo
                cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                cad = cad & DBSet((Rs!Importe * (-1)), "N") & "," & ValorNulo & "," & DBSet(Rs!cuenta, "T") & "," & ValorNulo & ",0"
            End If
            
            cad = "(" & cad & ")"
            
            B = InsertarLinAsientoDia(cad, cadMen)
            cadMen = "Insertando Lin. Asiento: " & I
            
            If B Then
                I = I + 1
                
                cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
                cad = cad & DBSet(I, "N") & "," & DBSet(Rs!cuenta, "T") & "," & DBSet(numdocum, "T") & ","
                If ImpAport > 0 Then
                    ' importe al haber en positivo
                    cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet(Rs!Importe, "N") & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & DBSet((Rs!Importe * (-1)), "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
                End If
            
                cad = "(" & cad & ")"
                
                B = InsertarLinAsientoDia(cad, cadMen)
                cadMen = "Insertando Lin. Asiento: " & I
            End If

        
            Rs.MoveNext
        Wend
        Set Rs = Nothing
        
    End If
'    'Insertar en la contabilidad
'    If cad <> "" Then
'        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
'        Sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
'        Sql = Sql & " VALUES " & cad
'        ConnConta.Execute Sql
'    End If
    
    Tipo = TipoAnt

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinAsientoFactIntProv = False
        cadErr = Err.Description
    Else
        InsertarLinAsientoFactIntProv = True
    End If
    Set Rs = Nothing
    InsertarLinAsientoFactIntProv = B
    Exit Function
End Function





Public Sub FechasEjercicioConta(FIni As String, FFin As String)
'Dim RS As ADODB.Recordset
'
'    On Error GoTo EFechas
'
'    FIni = "Select fechaini,fechafin From parametros"
'    Set RS = New ADODB.Recordset
'    RS.Open FIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not RS.EOF Then
'        FIni = DBLet(RS!FechaIni, "F")
'        FFin = DBLet(RS!FechaFin, "F")
'    End If
'    RS.Close
'    Set RS = Nothing
'
'EFechas:
'    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------------
' FACTURAS TRANSPORTE
'----------------------------------------------------------------------

Public Function PasarFacturaTerc(cadWHERE As String, CodCCost As String, FechaFin As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariagro.tcafpc --> conta.cabfactprov
' ariagro.tlifpc --> conta.linfactprov
'Actualizar la tabla ariges.scafpc.inconta=1 para indicar que ya esta contabilizada
Dim B As Boolean
Dim cadMen As String
Dim Sql As String
Dim Mc As Contadores


    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
        
    
    Set Mc = New Contadores
    
    Set vSoc = New cSocio
    
    Sql = "select codsocio from rcafter where " & cadWHERE
    vSoc.LeerDatos DevuelveValor(Sql)
    
    
    '---- Insertar en la conta Cabecera Factura
    B = InsertarCabFactTerc(cadWHERE, cadMen, Mc, FechaFin, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If B Then
        CCoste = CodCCost
        '---- Insertar lineas de Factura en la Conta
        If vParamAplic.ContabilidadNueva Then
            B = InsertarLinFact_newContaNueva("rcafter", cadWHERE, cadMen, Mc.Contador)
        Else
            B = InsertarLinFact_new("rcafter", cadWHERE, cadMen, Mc.Contador)
        End If
        cadMen = "Insertando Lin. Factura: " & cadMen

        If B Then
            '$$$$
'            If vParamAplic.ContabilidadNueva Then
'                If vParamAplic.Cooperativa = 12 Then
'                    b = InsertarEnTesoreriaTercMontifrut(cadWHERE)
'                Else
'                    b = InsertarEnTesoreriaTerc(cadWHERE)
'                End If
'            End If

            If Not EsFacturaInterna(cadWHERE) Then
                If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaProv(vContaFra.NumeroFactura, vContaFra.Anofac)
            End If

            '---- Poner intconta=1 en ariges.scafac
            B = ActualizarCabFact("rcafter", cadWHERE, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
    Set vSoc = Nothing
    
    
'    If Not b Then
'        SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
'        SQL = SQL & " Select *," & DBSet(Mid(cadMen, 1, 200), "T") & " as error From tmpFactu "
'        SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
'        Conn.Execute SQL
'    End If
    
EContab:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If B Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaTerc = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaTerc = False
        If Not B Then
            InsertarTMPErrFac cadMen, cadWHERE
'            SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
'            SQL = SQL & " Select *," & DBSet(Mid(cadMen, 1, 200), "T") & " as error From tmpFactu "
'            SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
'            Conn.Execute SQL
        End If
    End If
End Function

Private Function InsertarCabFactTerc(cadWHERE As String, cadErr As String, ByRef Mc As Contadores, FechaFin As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Nulo4 As String

Dim TipoOpera As Integer
Dim Aux As String

Dim Sql2 As String
Dim CadenaInsertFaclin2 As String
Dim ImporAux As Currency

Dim ImporAux2 As Currency
    
    
    On Error GoTo EInsertar


    Sql = " SELECT fecfactu,year(fecrecep) as anofacpr,fecrecep,numfactu,rsocios_seccion.codmacpro as codmacta,"
    Sql = Sql & "rcafter.dtoppago,rcafter.dtognral,baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    Sql = Sql & "totalfac,tipoiva1,tipoiva2,tipoiva3, retfacpr, trefacpr, rsocios_seccion.codsocio, rsocios.nomsocio, rcafter.intracom, "
    Sql = Sql & "rsocios.dirsocio,rsocios.pobsocio,rsocios.codpostal,rsocios.prosocio,rsocios.nifsocio, rcafter.codforpa "
    Sql = Sql & " FROM (" & "rcafter "
    Sql = Sql & "INNER JOIN " & "rsocios ON rcafter.codsocio=rsocios.codsocio )"
    Sql = Sql & " INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.Seccionhorto, "N")
    Sql = Sql & " WHERE " & cadWHERE

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    If Not Rs.EOF Then

        If Mc.ConseguirContador("1", (Rs!FecRecep <= CDate(FechaFin) - 365), True) = 0 Then
            vContaFra.NumeroFactura = Mc.Contador
            vContaFra.Anofac = DBLet(Rs!anofacpr)

            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            DtoPPago = Rs!DtoPPago
            DtoGnral = Rs!DtoGnral
            BaseImp = Rs!BaseIVA1 + CCur(DBLet(Rs!BaseIVA2, "N")) + CCur(DBLet(Rs!BaseIVA3, "N"))
            TotalFac = Rs!TotalFac
            AnyoFacPr = Rs!anofacpr

            FecRecep = DBLet(Rs!FecRecep, "F")
            ForPago = DBLet(Rs!Codforpa)

            mCodmacta = DBLet(Rs!Codmacta)

            Nulo2 = "N"
            Nulo3 = "N"
            Nulo4 = "N"
            If DBLet(Rs!BaseIVA2, "N") = "0" Then Nulo2 = "S"
            If DBLet(Rs!BaseIVA3, "N") = "0" Then Nulo3 = "S"
            If DBLet(Rs!trefacpr, "N") = "0" Then Nulo4 = "S"
            Sql = ""
            If vParamAplic.ContabilidadNueva Then Sql = "'" & SerieFraPro & "',"
            Sql = Sql & Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & Rs!anofacpr & "," & DBSet(Rs!FecRecep, "F") & "," & DBSet(Rs!FecRecep, "F") & "," & DBSet(Rs!numfactu, "T") & "," & DBSet(Rs!Codmacta, "T") & "," & ValorNulo & ","
            
            If vParamAplic.ContabilidadNueva Then
                Sql = Sql & DBSet(Rs!nomsocio, "T") & "," & DBSet(Rs!dirsocio, "T", "S") & ","
                Sql = Sql & DBSet(Rs!codpostal, "T", "S") & "," & DBSet(Rs!pobsocio, "T", "S") & "," & DBSet(Rs!prosocio, "T", "S") & ","
                Sql = Sql & DBSet(Rs!nifSocio, "T", "S")
                If DBLet(Rs!Intracom) = 1 Then
                    Dim Pais As String
                    Pais = DevuelveDesdeBDNew(cConta, "cuentas", "codpais", "codmacta", mCodmacta, "T")
                
                    Sql = Sql & "," & DBSet(Pais, "T", "S") & ","
                Else
                    Sql = Sql & ",'ES',"
                End If
                Sql = Sql & DBSet(Rs!Codforpa, "N") & ","
                
                TipoOpera = 0
                
                If DBLet(Rs!Intracom) = 1 Then TipoOpera = 1
                
                Aux = "0"
                Select Case TipoOpera
                Case 0
                    If Rs!TotalFac < 0 Then
                        Aux = "D"
                    Else
                        If Not IsNull(Rs!TipoIVA2) Then Aux = "C"
                    End If
                End Select
                
                If DBLet(Rs!Intracom) = 1 Then Aux = "P"
                
                'codopera,codconce340,codintra
                Sql = Sql & TipoOpera & "," & DBSet(Aux, "T") & ","
                
                If DBLet(Rs!Intracom) = 1 Then
                    Sql = Sql & "'A',"
                Else
                    Sql = Sql & ValorNulo & ","
                End If
                
                
                'para las lineas
                'factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
                'IVA 1, siempre existe
                Aux = "'" & SerieFraPro & "'," & Mc.Contador & "," & DBSet(Rs!FecRecep, "F") & "," & Rs!anofacpr & ","
                
                Sql2 = Aux & "1," & DBSet(Rs!BaseIVA1, "N") & "," & Rs!TipoIVA1 & "," & DBSet(Rs!porciva1, "N") & ","
                Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
                vTipoIva(0) = Rs!TipoIVA1
                vPorcIva(0) = Rs!porciva1
                vPorcRec(0) = 0
                vImpIva(0) = Rs!impoiva1
                vImpRec(0) = 0
                vBaseIva(0) = Rs!BaseIVA1
                
                vTipoIva(1) = 0: vTipoIva(2) = 0
                
                If Not IsNull(Rs!porciva2) Then
                    Sql2 = Aux & "2," & DBSet(Rs!BaseIVA2, "N") & "," & Rs!TipoIVA2 & "," & DBSet(Rs!porciva2, "N") & ","
                    Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva2, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
                    vTipoIva(1) = Rs!TipoIVA2
                    vPorcIva(1) = Rs!porciva2
                    vPorcRec(1) = 0
                    vImpIva(1) = Rs!impoiva2
                    vImpRec(1) = 0
                    vBaseIva(1) = Rs!BaseIVA2
                End If
                
                If Not IsNull(Rs!porciva3) Then
                    Sql2 = Aux & "3," & DBSet(Rs!BaseIVA3, "N") & "," & Rs!TipoIVA3 & "," & DBSet(Rs!porciva3, "N") & ","
                    Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva3, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
                    vTipoIva(2) = Rs!TipoIVA3
                    vPorcIva(2) = Rs!porciva3
                    vPorcRec(2) = 0
                    vImpIva(2) = Rs!impoiva3
                    vImpRec(2) = 0
                    vBaseIva(2) = Rs!BaseIVA3
                End If
                    
                    
                'Los totales
                'totbases,totbasesret,totivas,totrecargo,totfacpr,
                ImporAux = Rs!BaseIVA1 + DBLet(Rs!BaseIVA2, "N") + DBLet(Rs!BaseIVA3, "N")
                Sql = Sql & DBSet(ImporAux, "N") & ","
                
                If DBLet(Rs!retfacpr, "N") <> 0 Then
                    ImporAux2 = Rs!impoiva1 + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
                    Sql = Sql & DBSet(ImporAux + ImporAux2, "N")
                Else
                    Sql = Sql & ValorNulo
                End If
                Sql = Sql & ","

                
                        
                'totivas
                ImporAux = Rs!impoiva1 + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
                Sql = Sql & DBSet(ImporAux, "N") & "," & DBSet(TotalFac, "N") & ","
                
                If DBLet(Rs!retfacpr, "N") <> 0 Then
                    Sql = Sql & DBSet(Rs!retfacpr, "N") & "," & DBSet(Rs!trefacpr, "N") & "," & DBSet(vParamAplic.CtaRetenSoc, "T") & ",2"
                Else
                    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                End If
                
                cad = cad & "(" & Sql & ")"
            
                Sql = "INSERT INTO factpro(numserie,numregis,fecfactu,anofactu,fecharec,fecliqpr,numfactu,codmacta,observa,nommacta,"
                Sql = Sql & "dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,codforpa,codopera,codconce340,codintra,"
                Sql = Sql & "totbases,totbasesret,totivas,totfacpr,retfacpr,trefacpr,cuereten,tiporeten)"
                Sql = Sql & " VALUES " & cad
                ConnConta.Execute Sql
                
                'Las  lineas de IVA
                Sql = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)"
                Sql = Sql & " VALUES " & CadenaInsertFaclin2
                ConnConta.Execute Sql
            Else
                Sql = Sql & DBSet(Rs!BaseIVA1, "N") & "," & DBSet(Rs!BaseIVA2, "N", "S") & "," & DBSet(Rs!BaseIVA3, "N", "S") & ","
                Sql = Sql & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", Nulo2) & "," & DBSet(Rs!porciva3, "N", Nulo3) & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & DBSet(Rs!impoiva2, "N", Nulo2) & "," & DBSet(Rs!impoiva3, "N", Nulo3) & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!TipoIVA1, "N") & "," & DBSet(Rs!TipoIVA2, "N", Nulo2) & "," & DBSet(Rs!TipoIVA3, "N", Nulo3) & ","
                Sql = Sql & DBSet(Rs!Intracom, "N") & ","
                Sql = Sql & DBSet(Rs!retfacpr, "N", Nulo4) & "," & DBSet(Rs!trefacpr, "N", Nulo4) & ","
                If Nulo4 = "S" Then
                    Sql = Sql & ValorNulo & ","
                Else
                    Sql = Sql & DBSet(vParamAplic.CtaTerReten, "T") & ","
                End If
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & "0"
                cad = cad & "(" & Sql & ")"
    
                'Insertar en la contabilidad
                Sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,fecliqpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                Sql = Sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                Sql = Sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,nodeducible) "
                Sql = Sql & " VALUES " & cad
                ConnConta.Execute Sql
            
            
            End If
            'añadido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            Sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            Sql = Sql & ",'" & DevNombreSQL(Rs!numfactu) & " @ " & Format(Rs!fecfactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!nomsocio) & "'," & Rs!Codsocio & ")"
            conn.Execute Sql

        End If
    End If
    Rs.Close
    Set Rs = Nothing

EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactTerc = False
        cadErr = Err.Description
    Else
        InsertarCabFactTerc = True
    End If
End Function


Public Function InsertarEnTesoreriaSoc(cadWHERE As String, MenError As String, numfactu As String, fecfactu As Date) As Boolean
'Guarda datos de Tesoreria en tablas: spagop o scobro dependiendo del signo de la factura
Dim B As Boolean
Dim Sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim CC As String
Dim letraser As String
Dim Text33csb As String
Dim Text41csb As String
Dim Text42csb As String
Dim GastosPie As Currency
Dim GastosVarias As Currency
Dim FactuRec As String
Dim rsVenci As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim FecVenci1 As Date
Dim ImpVenci As Currency

Dim vBancoSoc As String
Dim vSucurSoc As String

Dim PorcCorredor As Currency
Dim TotalTesor1 As Currency

Dim UltimoVto As Integer

Dim CadValuesGastos As String
Dim CadValuesVarias As String
Dim SqlGastos As String
Dim J As Integer

    On Error GoTo EInsertarTesoreriaSoc

    InsertarEnTesoreriaSoc = False
    
    
    '[Monica] 21/01/2010 tenemos que descontar del totaltesor los gastos a pie de factura
    Sql = "select sum(importe) from rfactsoc_gastos where " & Replace(cadWHERE, "rfactsoc", "rfactsoc_gastos")
    GastosPie = DevuelveValor(Sql)
    '[Monica]29/11/2013: si es Montifrut los gastos a pie no se descuentan del importe
    If vParamAplic.Cooperativa = 12 Then GastosPie = 0
    
    
    '[Monica] 13/06/2013 tenemos que descontar las facturas varias que se insertaron
    Sql = "select sum(totalfac) from fvarcabfact where (codsecci, codtipom, numfactu, fecfactu) in (select codsecci, codtipomfvar, numfactufvar, fecfactufvar from rfactsoc_fvarias where " & Replace(cadWHERE, "rfactsoc", "rfactsoc_fvarias") & ")"
    GastosVarias = DevuelveValor(Sql)
    
    
    '[Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay
    FactuRec = DevuelveValor("select numfacrec from rfactsoc where " & cadWHERE)
    If FactuRec = "0" Then
        letraser = ""
        letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(tipoMov, "T"))
    
        FactuRec = letraser & "-" & numfactu
    End If
    
    vBancoSoc = ""
    If BancoSoc <> 0 Then vBancoSoc = BancoSoc
    
    vSucurSoc = ""
    If SucurSoc <> 0 Then vSucurSoc = SucurSoc
    
    
    TotalTesor = TotalTesor - GastosPie - GastosVarias
    
    
    'si hay porcentaje de corredor hemos de descontarlo tb. Este porcentaje lo cargaba Montifrut
    Sql = "select porccorredor from rfactsoc where " & cadWHERE
    PorcCorredor = DevuelveValor(Sql)
    
    TotalTesor1 = Round2(TotalTesor * PorcCorredor / 100, 2)
    TotalTesor = TotalTesor - Round2(TotalTesor * PorcCorredor / 100, 2)
    
    If TotalTesor >= 0 Then ' se insertara en la cartera de pagos (spagop)
        
        '[Monica]09/05/2013: Añadido el nro de vencimientos
        CadValues2 = ""
        CadValuesGastos = ""
        CadValuesVarias = ""
        
        Sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & ForpaPosi
        Set rsVenci = New ADODB.Recordset
        rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not rsVenci.EOF Then
            If rsVenci!numerove > 0 Then
                
                'vamos creando la cadena para insertar en spagosp de la CONTA
                letraser = ""
                letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(tipoMov, "T"))
                
                'Obtener los dias de pago de la tabla de parametros: spara1
                    
                    '[Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
                    '[Monica]03/07/2013: añado trim(codmacta)
                    CadValuesAux2 = "("
                    If vParamAplic.ContabilidadNueva Then CadValuesAux2 = CadValuesAux2 & DBSet(SerieFraPro, "T") & ","
                    CadValuesAux2 = CadValuesAux2 & "'" & Trim(CtaSocio) & "', " & DBSet(FactuRec, "T") & ", '" & Format(fecfactu, FormatoFecha) & "', "
                    
                      'Primer Vencimiento
                      '------------------------------------------------------------
                      I = 1
                      'FECHA VTO
                      FecVenci = CDate(fecfactu)
                      '=== Modificado: Laura 23/01/2007
        '              FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                      FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                      '==================================
                      'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                      
                      FecVenci1 = FecVenci
        
        
                      CadValues2 = CadValuesAux2 & I
                      CadValues2 = CadValues2 & ", " & ForpaPosi & ", '" & Format(FecVenci, FormatoFecha) & "', "
                      
                      '[Monica]28/04/2017: nuevo pago pagado de los gastos y fras varias
                      If vParamAplic.ContabilidadNueva Then
                        If GastosPie <> 0 Then
                            I = I + 1
                            CadValuesGastos = CadValuesAux2 & I
                            CadValuesGastos = CadValuesGastos & ", " & ForpaPosi & ", '" & Format(FecVenci, FormatoFecha) & "', "
                        End If
                        
                        If GastosVarias <> 0 Then
                            I = I + 1
                            CadValuesVarias = CadValuesAux2 & I
                            CadValuesVarias = CadValuesVarias & ", " & ForpaPosi & ", '" & Format(FecVenci, FormatoFecha) & "', "
                        End If
                      End If
                      
                      'IMPORTE del Vencimiento
                      If rsVenci!numerove = 1 Then
                          ImpVenci = TotalTesor
                      Else
                          ImpVenci = Round(TotalTesor / rsVenci!numerove, 2)
                          'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                          If ImpVenci * rsVenci!numerove <> TotalTesor Then
                              ImpVenci = Round(ImpVenci + (TotalTesor - ImpVenci * rsVenci!numerove), 2)
                          End If
                      End If
                      
                      CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBanco, "T") & ","
                      '[Monica]28/04/2017: nuevo pago pagado de los gastos y fras varias
                      If vParamAplic.ContabilidadNueva Then
                        If GastosPie <> 0 Then
                              CadValuesGastos = CadValuesGastos & DBSet(GastosPie, "N") & ", " & DBSet(CtaBanco, "T") & ","
                        End If
                        If GastosVarias <> 0 Then
                              CadValuesVarias = CadValuesVarias & DBSet(GastosVarias, "N") & ", " & DBSet(CtaBanco, "T") & ","
                        End If
                      End If
                
                      If Not vParamAplic.ContabilidadNueva Then
                            'David. Para que ponga la cuenta bancaria (SI LA tiene)
                            CadValues2 = CadValues2 & DBSet(vBancoSoc, "T", "S") & "," & DBSet(vSucurSoc, "T", "S") & ","
                            CadValues2 = CadValues2 & DBSet(DigcoSoc, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & ","
                      End If
                
                      'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
                      Select Case TipoFact
                        Case 1, 3, 7, 9 ' anticipo y anticipo venta campo
            '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
            '                Sql = "Anticipo num.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                            Sql = "Anticipo num.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                        Case 2, 4, 8, 10 ' liquidacion y liquidacion venta campo
            '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
            '                Sql = "Liquidacion num.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                            Sql = "Liquidacion num.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                        Case Else
            '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
            '                Sql = "Fact.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                            Sql = "Fact.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                      End Select
                        
                      CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
                      '[Monica]28/04/2017: nuevo pago pagado de los gastos y fras varias
                      If vParamAplic.ContabilidadNueva Then
                        If GastosPie <> 0 Then
                            CadValuesGastos = CadValuesGastos & "'" & DevNombreSQL(Sql) & "',"
                        End If
                        If GastosVarias <> 0 Then
                            CadValuesVarias = CadValuesVarias & "'" & DevNombreSQL(Sql) & "',"
                        End If
                      End If
                    
                      Sql = "Variedades: " & Variedades
                      CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" '),"
                      
                      '[Monica]28/04/2017: nuevo pago pagado de los gastos y fras varias
                      If vParamAplic.ContabilidadNueva Then
                        If GastosPie <> 0 Then
                            CadValuesGastos = CadValuesGastos & "'" & DevNombreSQL(Sql) & "'"
                        End If
                        If GastosVarias <> 0 Then
                            CadValuesVarias = CadValuesVarias & "'" & DevNombreSQL(Sql) & "'"
                        End If
                      End If
                      
                      If vParamAplic.ContabilidadNueva Then

                            vvIban = MiFormat(IbanSoc, "") & MiFormat(vBancoSoc, "0000") & MiFormat(vSucurSoc, "0000") & MiFormat(DigcoSoc, "00") & MiFormat(CtaBaSoc, "0000000000")
                            
                            CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                            'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                            CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                            CadValues2 = CadValues2 & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES',"
                            
                            If TotalTesor = 0 Then
                                CadValues2 = CadValues2 & DBSet(fecfactu, "F") & "," & DBSet(0, "N") & ",1),"
                            Else
                                CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & ",0),"
                            End If
                            
                          '[Monica]28/04/2017: nuevo pago pagado de los gastos y fras varias
                            If GastosPie <> 0 Then
                                CadValuesGastos = CadValuesGastos & "," & DBSet(vvIban, "T") & ","
                                'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                                CadValuesGastos = CadValuesGastos & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                                CadValuesGastos = CadValuesGastos & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'," & DBSet(fecfactu, "F") & "," & DBSet(GastosPie, "N") & ",1),"
                            End If
                            If GastosVarias <> 0 Then
                                CadValuesVarias = CadValuesVarias & "," & DBSet(vvIban, "T") & ","
                                'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                                CadValuesVarias = CadValuesVarias & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                                CadValuesVarias = CadValuesVarias & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'),"
                            End If
                      Else
                            '[Monica]22/11/2013: Tema iban
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & "),"
                            Else
                                CadValues2 = CadValues2 & "),"
                            End If
                      
                      End If
                      'Resto Vencimientos
                      '--------------------------------------------------------------------
                      UltimoVto = I
                      For J = 2 To rsVenci!numerove
                          UltimoVto = I + J - 1
                         'FECHA Resto Vencimientos
                          '==== Modificado: Laura 23/01/2007
                          'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                          FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                          '==================================================
        
                          CadValues2 = CadValues2 & CadValuesAux2 & UltimoVto 'i
                          CadValues2 = CadValues2 & ", " & ForpaPosi & ", '" & Format(FecVenci, FormatoFecha) & "', "
        
                          'IMPORTE Resto de Vendimientos
                          ImpVenci = Round(TotalTesor / rsVenci!numerove, 2)
        
                          CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & "," & DBSet(CtaBanco, "T") & ","
                          
                          If Not vParamAplic.ContabilidadNueva Then
                                'David. Para que ponga la cuenta bancaria (SI LA tiene)
                                CadValues2 = CadValues2 & DBSet(vBancoSoc, "T", "S") & "," & DBSet(vSucurSoc, "T", "S") & ","
                                CadValues2 = CadValues2 & DBSet(DigcoSoc, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & ","
                          End If
                          'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
                          Select Case TipoFact
                            Case 1, 3, 7, 9 ' anticipo y anticipo venta campo
                '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
                '                Sql = "Anticipo num.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                                Sql = "Anticipo num.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                            Case 2, 4, 8, 10 ' liquidacion y liquidacion venta campo
                '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
                '                Sql = "Liquidacion num.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                                Sql = "Liquidacion num.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                            Case Else
                '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
                '                Sql = "Fact.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                                Sql = "Fact.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                          End Select
                            
                          CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
                        
                          Sql = "Variedades: " & Variedades
                          CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" '),"
                          
                          If vParamAplic.ContabilidadNueva Then
                                
                                vvIban = MiFormat(IbanSoc, "") & MiFormat(vBancoSoc, "0000") & MiFormat(vSucurSoc, "0000") & MiFormat(DigcoSoc, "00") & MiFormat(CtaBaSoc, "0000000000")
                                
                                CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                                'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                                CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                                CadValues2 = CadValues2 & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'," & ValorNulo & "," & ValorNulo & ",0),"
                          Else
                                
                                '[Monica]22/11/2013: Tema iban
                                If vEmpresa.HayNorma19_34Nueva = 1 Then
                                    CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & "),"
                                Else
                                    CadValues2 = CadValues2 & "),"
                                End If
                          End If
                      Next J
                      
                      
                      'Ultimo Vencimiento es si lo hay la parte de descuento
                      If TotalTesor1 <> 0 Then ' For i = 2 To rsVenci!numerove
                          I = UltimoVto + 1
                          
        
                          CadValues2 = CadValues2 & CadValuesAux2 & I & ", " & ForpaPosi & ", '" & Format(FecVenci1, FormatoFecha) & "', "
        
                          'IMPORTE Resto de Vendimientos
                          ImpVenci = TotalTesor1  'Round2(TotalTesor / rsVenci!numerove, 2)
        
                          CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBanco, "T") & ","
                          
                          If Not vParamAplic.ContabilidadNueva Then
                                'David. Para que ponga la cuenta bancaria (SI LA tiene)
                                CadValues2 = CadValues2 & DBSet(vBancoSoc, "T", "S") & "," & DBSet(vSucurSoc, "T", "S") & ","
                                CadValues2 = CadValues2 & DBSet(DigcoSoc, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & ","
                          End If
                          
                          'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
                          Select Case TipoFact
                            Case 1, 3, 7, 9 ' anticipo y anticipo venta campo
                '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
                '                Sql = "Anticipo num.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                                Sql = "Anticipo num.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                            Case 2, 4, 8, 10 ' liquidacion y liquidacion venta campo
                '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
                '                Sql = "Liquidacion num.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                                Sql = "Liquidacion num.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                            Case Else
                '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
                '                Sql = "Fact.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                                Sql = "Fact.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                          End Select
                            
                          CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
                          Sql = "Variedades: " & Variedades
                          CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" '),"
                          
                          If vParamAplic.ContabilidadNueva Then
                                vvIban = MiFormat(IbanSoc, "") & MiFormat(vBancoSoc, "0000") & MiFormat(vSucurSoc, "0000") & MiFormat(DigcoSoc, "00") & MiFormat(CtaBaSoc, "0000000000")
                                
                                CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                                'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                                CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                                CadValues2 = CadValues2 & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'," & ValorNulo & "," & ValorNulo & ",0),"
                          Else
                                '[Monica]22/11/2013: Tema iban
                                If vEmpresa.HayNorma19_34Nueva = 1 Then
                                    CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & "),"
                                Else
                                    CadValues2 = CadValues2 & "),"
                                End If
                          End If
                      
                      End If
'                      Next i
                      
                    If CadValues2 <> "" Then
                        CadValues2 = Mid(CadValues2, 1, Len(CadValues2) - 1)
                    
                        'Insertamos en la tabla spagop de la CONTA
                        'David. Cuenta bancaria y descripcion textos
                        
                        If vParamAplic.ContabilidadNueva Then
                            Sql = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, iban,"
                            Sql = Sql & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais, fecultpa, imppagad, situacion)"
                        
                            SqlGastos = Sql
                        
                        Else
                            Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
                            '[Monica]22/11/2013: Tema iban
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                Sql = Sql & ", iban) "
                            Else
                                Sql = Sql & ") "
                            End If
                        End If
                        
                        Sql = Sql & " VALUES " & CadValues2
                        ConnConta.Execute Sql
                        
                        '[Monica]28/04/2017: nuevo pago pagado de los gastos y fras varias
                        If vParamAplic.ContabilidadNueva Then
                            If GastosPie <> 0 Then
                                Sql = SqlGastos & " VALUES " & Mid(CadValuesGastos, 1, Len(CadValuesGastos) - 1)
                                ConnConta.Execute Sql
                            End If
                        
                            If GastosVarias <> 0 Then
                                Sql = SqlGastos & " VALUES " & Mid(CadValuesVarias, 1, Len(CadValuesVarias) - 1)
                                ConnConta.Execute Sql
                            End If
                        End If
                        
                    End If
                      
            End If
        End If
        
        'hasta aqui de momento
        
' esto es lo que habia antes
'
'        CadValues2 = ""
'
'        'vamos creando la cadena para insertar en spagosp de la CONTA
'        letraser = ""
'        letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(tipoMov, "T"))
'
'        '[Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
'        CadValuesAux2 = "('" & CtaSocio & "', " & DBSet(FactuRec, "T") & ", '" & Format(fecfactu, FormatoFecha) & "', "
'
'        '------------------------------------------------------------
'        I = 1
'        CadValues2 = CadValuesAux2 & I
'
'        CadValues2 = CadValues2 & ", " & ForpaPosi & ", '" & Format(FecVenci, FormatoFecha) & "', "
'        CadValues2 = CadValues2 & DBSet(TotalTesor, "N") & ", " & DBSet(CtaBanco, "T") & ","
'
'        'David. Para que ponga la cuenta bancaria (SI LA tiene)
'        CadValues2 = CadValues2 & DBSet(BancoSoc, "T", "S") & "," & DBSet(SucurSoc, "T", "S") & ","
'        CadValues2 = CadValues2 & DBSet(DigcoSoc, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & ","
'
'        'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
'        Select Case TipoFact
'            Case 1, 3, 7, 9 ' anticipo y anticipo venta campo
''                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
''                Sql = "Anticipo num.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
'                Sql = "Anticipo num.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
'            Case 2, 4, 8, 10 ' liquidacion y liquidacion venta campo
''                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
''                Sql = "Liquidacion num.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
'                Sql = "Liquidacion num.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
'            Case Else
''                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
''                Sql = "Fact.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
'                Sql = "Fact.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
'        End Select
'
'        CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
'
'        Sql = "Variedades: " & Variedades
'        CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "')"
'
'        'Grabar tabla spagop de la CONTABILIDAD
'        '-------------------------------------------------
'        If CadValues2 <> "" Then
'            'Insertamos en la tabla spagop de la CONTA
'            'David. Cuenta bancaria y descripcion textos
'            Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb) "
'            Sql = Sql & " VALUES " & CadValues2
'            ConnConta.Execute Sql
'        End If
        
    Else
        ' si es negativo se inserta en positivo en la cartera de cobros (scobro)
'[Monica]09/05/2013: añadido los vencimientos
        letraser = ""
        letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(tipoMov, "T"))

'                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
'        Text33csb = "'Factura:" & DBLet(letraser, "T") & "-" & DBLet(numfactu, "T") & " " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
        Text33csb = "'Factura:" & DBLet(FactuRec, "T") & " " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
        Text41csb = "de " & DBSet(TotalTesor * (-1), "N")
        Text42csb = "Variedades: " & Variedades

        Sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & DBSet(ForpaNega, "N")
        Set rsVenci = New ADODB.Recordset
        rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        If Not rsVenci.EOF Then
            If DBLet(rsVenci!numerove, "N") > 0 Then
                
                CadValuesAux2 = "(" & DBSet(letraser, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & ","

                '-------- Primer Vencimiento
                I = 1
                'FECHA VTO
                FecVenci = DBLet(fecfactu, "F")
                '=== Laura 23/01/2007
                'FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                '===
                
                CadValues2 = CadValuesAux2 & I & ", "
                '[Monica]03/07/2013: añado trim(codmacta)
                CadValues2 = CadValues2 & DBSet(Trim(CtaSocio), "T") & ", " & DBSet(ForpaNega, "N") & ", " & DBSet(FecVenci, "F") & ", "
                
                
                '[Monica]28/04/2017: nuevo pago pagado de los gastos y fras varias
                If vParamAplic.ContabilidadNueva Then
                  If GastosPie <> 0 Then
                      I = I + 1
                      CadValuesGastos = CadValuesAux2 & I & ", "
                      CadValuesGastos = CadValuesGastos & DBSet(Trim(CtaSocio), "T") & ", " & ForpaNega & ", '" & Format(FecVenci, FormatoFecha) & "', "
                  End If
                  
                  If GastosVarias <> 0 Then
                      I = I + 1
                      CadValuesVarias = CadValuesAux2 & I & ", "
                      CadValuesVarias = CadValuesVarias & DBSet(Trim(CtaSocio), "T") & ", " & ForpaNega & ", '" & Format(FecVenci, FormatoFecha) & "', "
                  End If
                End If
                
                'IMPORTE del Vencimiento
                ImpVenci = TotalTesor * (-1)

                CC = DBLet(DigcoSoc, "T")
                If DBLet(DigcoSoc, "T") = "**" Then CC = "00"
        
                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ","
                CadValues2 = CadValues2 & DBSet(CtaBanco, "T") & ","
                
                '[Monica]28/04/2017: nuevo pago pagado de los gastos y fras varias
                If vParamAplic.ContabilidadNueva Then
                    If GastosPie <> 0 Then
                        CadValuesGastos = CadValuesGastos & DBSet(GastosPie * (-1), "N") & ","
                        CadValuesGastos = CadValuesGastos & DBSet(CtaBanco, "T") & ","
                    End If
                  
                    If GastosVarias <> 0 Then
                        CadValuesVarias = CadValuesVarias & DBSet(GastosVarias * (-1), "N") & ","
                        CadValuesVarias = CadValuesVarias & DBSet(CtaBanco, "T") & ","
                    End If
                End If
                
                
                
                
                
                If vParamAplic.ContabilidadNueva Then
                    CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text42csb, "T") & ",1,"
                    
                    vvIban = MiFormat(IbanSoc, "") & MiFormat(vBancoSoc, "0000") & MiFormat(vSucurSoc, "0000") & MiFormat(DigcoSoc, "00") & MiFormat(CtaBaSoc, "0000000000")
                    
                    CadValues2 = CadValues2 & DBSet(vvIban, "T") & ","
                    'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                    CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                    CadValues2 = CadValues2 & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES',"
                    
                    
                    If TotalTesor <> 0 Then
                        CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & ",0),"
                    Else
                        CadValues2 = CadValues2 & DBSet(fecfactu, "F") & "," & DBSet(TotalTesor * (-1), "N") & ",1),"
                    End If
                    
                    
                    '[Monica]28/04/2017: nuevo pago pagado de los gastos y fras varias
                    If GastosPie <> 0 Then
                        CadValuesGastos = CadValuesGastos & Text33csb & "," & DBSet(Text42csb, "T") & ",1,"
                        
                        CadValuesGastos = CadValuesGastos & DBSet(vvIban, "T") & ","
                        'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                        CadValuesGastos = CadValuesGastos & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                        CadValuesGastos = CadValuesGastos & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'," & DBSet(fecfactu, "F") & "," & DBSet(GastosPie * (-1), "N") & ",1),"
                    End If
                    If GastosVarias <> 0 Then
                        CadValuesVarias = CadValuesVarias & Text33csb & "," & DBSet(Text42csb, "T") & ",1,"
                        
                        CadValuesVarias = CadValuesVarias & DBSet(vvIban, "T") & ","
                        'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                        CadValuesVarias = CadValuesVarias & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                        CadValuesVarias = CadValuesVarias & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'," & DBSet(fecfactu, "F") & "," & DBSet(GastosVarias * (-1), "N") & ",1),"
                    End If
                    
                Else
                    CadValues2 = CadValues2 & DBSet(vBancoSoc, "T", "S") & "," & DBSet(vSucurSoc, "T", "S") & ","
                    CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & ","
                    CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                    CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & "," & DBSet(Text42csb, "T") & ",1" '),"
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & "),"
                    Else
                        CadValues2 = CadValues2 & "),"
                    End If
                End If
                'Resto Vencimientos
                '--------------------------------------------------------------------
                If TotalTesor1 <> 0 Then 'For i = 2 To rsVenci!numerove
                   'FECHA Resto Vencimientos
                    '=== Laura 23/01/2007
                    'FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                    '===
                    I = I + 1
                    
                    CadValues2 = CadValues2 & CadValuesAux2 & I & ", " & DBSet(Trim(CtaSocio), "T") & ", " & DBSet(ForpaNega, "N") & ", '" & Format(FecVenci, FormatoFecha) & "', "
                    
                    'IMPORTE Resto de Vendimientos
                    'ImpVenci = Round2(TotalTesor * (-1) / rsVenci!numerove, 2)
                    ImpVenci = TotalTesor1 * (-1)
                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ","
                    
                    If Not vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & DBSet(CtaBanco, "T") & "," & DBSet(vBancoSoc, "N", "S") & "," & DBSet(vSucurSoc, "N", "S") & ","
                        CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                        CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & "," & DBSet(Text42csb, "T") & ",1" '),"
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & "),"
                        Else
                            CadValues2 = CadValues2 & "),"
                        End If
                    Else
                        CadValues2 = CadValues2 & DBSet(CtaBanco, "T") & ","
                        CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text42csb, "T") & ",1,"
                        
                        vvIban = MiFormat(IbanSoc, "") & MiFormat(vBancoSoc, "0000") & MiFormat(vSucurSoc, "0000") & MiFormat(DigcoSoc, "00") & MiFormat(CtaBaSoc, "0000000000")
                        
                        CadValues2 = CadValues2 & DBSet(vvIban, "T") & ","
                        'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                        CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                        CadValues2 = CadValues2 & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'," & ValorNulo & "," & ValorNulo & ",0),"
                    End If
                    
                End If
                'Next i
                ' quitamos la ultima coma
                CadValues2 = Mid(CadValues2, 1, Len(CadValues2) - 1)

                'Insertamos en la tabla scobro de la CONTA
                If vParamAplic.ContabilidadNueva Then
                    Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                    Sql = Sql & "ctabanc1, "
                    Sql = Sql & " text33csb, text41csb, agente, iban, " ') "
                    Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais, fecultco, impcobro, situacion"
                    Sql = Sql & ") "
                
                    SqlGastos = Sql
                Else
                    Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                    Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                    Sql = Sql & " text33csb, text41csb, text42csb, agente" ') "
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        Sql = Sql & ", iban) "
                    Else
                        Sql = Sql & ") "
                    End If
                End If
                
                Sql = Sql & " VALUES " & CadValues2
                ConnConta.Execute Sql
                
                '[Monica]28/04/2017: nuevo pago pagado de los gastos y fras varias
                If vParamAplic.ContabilidadNueva Then
                    If GastosPie <> 0 Then
                        Sql = SqlGastos & " VALUES " & Mid(CadValuesGastos, 1, Len(CadValuesGastos) - 1)
                        ConnConta.Execute Sql
                    End If
                
                    If GastosVarias <> 0 Then
                        Sql = SqlGastos & " VALUES " & Mid(CadValuesVarias, 1, Len(CadValuesGastos) - 1)
                        ConnConta.Execute Sql
                    End If
                End If
            End If
        End If
'hasta aqui de momento

' lo que habia antes
'        letraser = ""
'        letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(tipoMov, "T"))
'
''                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
''        Text33csb = "'Factura:" & DBLet(letraser, "T") & "-" & DBLet(numfactu, "T") & " " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
'        Text33csb = "'Factura:" & DBLet(FactuRec, "T") & " " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
'        Text41csb = "de " & DBSet(TotalTesor, "N")
'        Text42csb = "Variedades: " & Variedades
'
'        CC = DBLet(DigcoSoc, "T")
'        If DBLet(DigcoSoc, "T") = "**" Then CC = "00"
'
'        CadValuesAux2 = "(" & DBSet(letraser, "T") & "," & DBSet(NumFactu, "N") & "," & DBSet(fecfactu, "F") & ", 1," & DBSet(CtaSocio, "T") & ","
'        CadValues2 = CadValuesAux2 & DBSet(ForpaNega, "N") & "," & DBSet(fecfactu, "F") & "," & DBSet(TotalTesor * (-1), "N") & ","
'        CadValues2 = CadValues2 & DBSet(CtaBanco, "T") & "," & DBSet(BancoSoc, "N") & "," & DBSet(SucurSoc, "N") & ","
'        CadValues2 = CadValues2 & DBSet(CC, "T") & "," & DBSet(CtaBaSoc, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
'        CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & "," & DBSet(Text42csb, "T") & ",1)"
'
'        'Insertamos en la tabla scobro de la CONTA
'        Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
'        Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
'        Sql = Sql & " text33csb, text41csb, text42csb, agente) "
'        Sql = Sql & " VALUES " & CadValues2
'        ConnConta.Execute Sql
'
    End If

    B = True

EInsertarTesoreriaSoc:
    If Err.Number <> 0 Then
        B = False
        MenError = "Error al insertar en Tesoreria: " & Err.Description
    End If
    InsertarEnTesoreriaSoc = B
End Function

' ### [Monica] 16/01/2008
Public Function InsertarEnTesoreriaNewADV(cadWHERE As String, CtaBan As String, FecVen As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
Dim B As Boolean
Dim Sql As String, Text33csb As String, Text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim DigConta As String
Dim CC As String
Dim vrefer As String
Dim letraser As String
Dim Rsx As ADODB.Recordset
Dim vSocio As cSocio

    On Error GoTo EInsertarTesoreriaNew

    B = False
    InsertarEnTesoreriaNewADV = False
    CadValues = ""
    CadValues2 = ""
    
    Set Rsx = New ADODB.Recordset
    Sql = "select * from advfacturas where " & cadWHERE
    Rsx.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    If Not Rsx.EOF Then
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(DBLet(Rsx!Codsocio, "N")) Then
            If vSocio.LeerDatosSeccion(DBLet(Rsx!Codsocio, "N"), vParamAplic.SeccionADV) Then
'[Monica]27/09/2011: tanto si el importe es positivo o negativo se introduce en la scobro
'                If DBLet(Rsx!TotalFac, "N") >= 0 Then
                    ' si el importe de la factura es positiva o cero
                    letraser = ""
                    letraser = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "letraser", "codtipom", Rsx!CodTipom, "T")
        
                    Text33csb = "'Factura:" & DBLet(letraser, "T") & "-" & DBLet(Rsx!numfactu, "T") & " " & Format(DBLet(Rsx!fecfactu, "F"), "dd/mm/yy") & "'"
                    Text41csb = "de " & DBSet(Rsx!TotalFac, "N")
        
                    CC = DBLet(vSocio.Digcontrol, "T")
                    If DBLet(vSocio.Digcontrol, "T") = "**" Then CC = "00"
                    
                    '[Monica]03/07/2013: añado trim(codmacta)
                    CadValuesAux2 = "(" & DBSet(letraser, "T") & "," & DBSet(Rsx!numfactu, "N") & "," & DBSet(Rsx!fecfactu, "F") & ", 1," & DBSet(Trim(vSocio.CtaClien), "T") & ","
                    CadValues2 = CadValuesAux2 & DBSet(Rsx!Codforpa, "N") & "," & DBSet(FecVen, "F") & "," & DBSet(Rsx!TotalFac, "N") & ","
                    CadValues2 = CadValues2 & DBSet(CtaBan, "T") & ","
                    
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & ","
                        CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & ",1," ')"
                    
                        vvIban = MiFormat(IbanSoc, "") & MiFormat(vSocio.Banco, "0000") & MiFormat(vSocio.Sucursal, "0000") & MiFormat(DigcoSoc, "00") & MiFormat(CtaBaSoc, "0000000000")
                        
                        CadValues2 = CadValues2 & DBSet(vvIban, "T") & ","
                        'nomsocio,dirsocio,pobsocio,codpostal,prosocio,nifsocio
                        CadValues2 = CadValues2 & DBSet(Rsx!nomsocio, "T") & "," & DBSet(Rsx!dirsocio, "T") & "," & DBSet(Rsx!pobsocio, "T") & "," & DBSet(Rsx!codpostal, "T") & ","
                        CadValues2 = CadValues2 & DBSet(Rsx!prosocio, "T") & "," & DBSet(Rsx!nifSocio, "T") & ",'ES') "
                        
                        Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                        Sql = Sql & "ctabanc1, fecultco, impcobro, "
                        Sql = Sql & " text33csb, text41csb, agente, iban, " ') "
                        Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
                        Sql = Sql & ") "
                        
                    Else
                        CadValues2 = CadValues2 & DBSet(vSocio.Banco, "N", "S") & "," & DBSet(vSocio.Sucursal, "N", "S") & ","
                        CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(vSocio.CuentaBan, "T", "S") & ","
                        CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                        CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & ",1" ')"
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & "," & DBSet(vSocio.Iban, "T", "S") & ") "
                        Else
                            CadValues2 = CadValues2 & ") "
                        End If
                        'Insertamos en la tabla scobro de la CONTA
                        Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                        Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                        Sql = Sql & " text33csb, text41csb, agente" ') "
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            Sql = Sql & ", iban) "
                        Else
                            Sql = Sql & ") "
                        End If
                    
                    End If
                    
        
                    
                    Sql = Sql & " VALUES " & CadValues2
                    ConnConta.Execute Sql
                    
'[Monica]27/09/2011: quitamos todo el else
'                Else
'                    '********** si la factura es negativa se inserta en la spago con valor poositivo
'                    CadValues2 = ""
'
'                    CadValuesAux2 = "('" & vSocio.CtaClien & "', " & DBSet(Rsx!NumFactu, "N") & ", '" & Format(Rsx!fecfactu, FormatoFecha) & "', "
'
'                    '------------------------------------------------------------
'
'                    CC = DBLet(vSocio.Digcontrol, "T")
'                    If DBLet(vSocio.Digcontrol, "T") = "**" Then CC = "00"
'
'                    i = 1
'                    CadValues2 = CadValuesAux2 & i
'                    CadValues2 = CadValues2 & ", " & DBSet(Rsx!Codforpa, "N") & ", " & DBSet(FecVen, "F") & ", "
'                    CadValues2 = CadValues2 & DBSet(DBLet(Rsx!TotalFac, "N") * (-1), "N") & ", " & DBSet(CtaBan, "T") & ","
'
'                    'David. Para que ponga la cuenta bancaria (SI LA tiene)
'                    CadValues2 = CadValues2 & DBSet(vSocio.Banco, "T", "S") & "," & DBSet(vSocio.Sucursal, "T", "S") & ","
'                    CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(vSocio.CuentaBan, "T", "S") & ","
'
'                    'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
'                    SQL = "Factura ADV.Nro.:" & DBLet(Rsx!NumFactu, "N")
'
'                    CadValues2 = CadValues2 & "'" & DevNombreSQL(SQL) & "',"
'
'                    SQL = " de " & Format(DBLet(Rsx!fecfactu, "F"), "dd/mm/yyyy")
'                    CadValues2 = CadValues2 & "'" & DevNombreSQL(SQL) & "')"
'
'                    'Grabar tabla spagop de la CONTABILIDAD
'                    '-------------------------------------------------
'                    If CadValues2 <> "" Then
'                        'Insertamos en la tabla spagop de la CONTA
'                        'David. Cuenta bancaria y descripcion textos
'                        SQL = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb) "
'                        SQL = SQL & " VALUES " & CadValues2
'                        ConnConta.Execute SQL
'                    End If
'                    '*******
'                End If
            End If
        End If
    
        B = True
    End If
    
EInsertarTesoreriaNew:
    If Err.Number <> 0 Then
        B = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaNewADV = B
End Function



' ### [Monica] 16/01/2008
Public Function InsertarEnTesoreriaNewBOD(cadWHERE As String, CtaBan As String, FecVen As String, MenError As String, Tipo As Byte) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
' Tipo: 0 = almazara
'       1 = bodega

Dim B As Boolean
Dim Sql As String, Text33csb As String, Text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim DigConta As String
Dim CC As String
Dim vrefer As String
Dim letraser As String
Dim Rsx As ADODB.Recordset
Dim vSocio As cSocio
Dim Seccion As Integer
    On Error GoTo EInsertarTesoreriaNew

    B = False
    InsertarEnTesoreriaNewBOD = False
    CadValues = ""
    CadValues2 = ""
    
    Set Rsx = New ADODB.Recordset
    Sql = "select * from rbodfacturas where " & cadWHERE
    Rsx.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    If Not Rsx.EOF Then
    
        Select Case Tipo
            Case 0 ' almazara
                Seccion = vParamAplic.SeccionAlmaz
            Case 1 ' bodega
                Seccion = vParamAplic.SeccionBodega
        End Select
    
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(DBLet(Rsx!Codsocio, "N")) Then
            If vSocio.LeerDatosSeccion(DBLet(Rsx!Codsocio, "N"), CStr(Seccion)) Then
'[Monica]27/09/2011: tanto si es positivo como si es negativo se inserta en la cartera de cobros
'                If DBLet(Rsx!TotalFac, "N") >= 0 Then
                    ' si el importe de la factura es positiva o cero
                    letraser = ""
                    letraser = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "letraser", "codtipom", Rsx!CodTipom, "T")
        
                    Text33csb = "'Factura:" & DBLet(letraser, "T") & "-" & DBLet(Rsx!numfactu, "T") & " " & Format(DBLet(Rsx!fecfactu, "F"), "dd/mm/yy") & "'"
                    Text41csb = "de " & DBSet(Rsx!TotalFac, "N")
        
                    CC = DBLet(vSocio.Digcontrol, "T")
                    If DBLet(vSocio.Digcontrol, "T") = "**" Then CC = "00"
        
                    '[Monica]03/07/2013: añado trim(codmacta)
                    CadValuesAux2 = "(" & DBSet(letraser, "T") & "," & DBSet(Rsx!numfactu, "N") & "," & DBSet(Rsx!fecfactu, "F") & ", 1," & DBSet(Trim(vSocio.CtaClien), "T") & ","
                    CadValues2 = CadValuesAux2 & DBSet(Rsx!Codforpa, "N") & "," & DBSet(FecVen, "F") & "," & DBSet(Rsx!TotalFac, "N") & ","
                    CadValues2 = CadValues2 & DBSet(CtaBan, "T") & ","
                    
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & ","
                        CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & ",1" ')"
                    
                        vvIban = MiFormat(vSocio.Iban, "") & MiFormat(vSocio.Banco, "0000") & MiFormat(vSocio.Sucursal, "0000") & MiFormat(CC, "00") & MiFormat(vSocio.CuentaBan, "0000000000")
                        
                        CadValues2 = CadValues2 & DBSet(vvIban, "T") & ","
                        'nomsocio,dirsocio,pobsocio,codpostal,prosocio,nifsocio
                        CadValues2 = CadValues2 & DBSet(Rsx!nomsocio, "T") & "," & DBSet(Rsx!dirsocio, "T") & "," & DBSet(Rsx!pobsocio, "T") & "," & DBSet(Rsx!codpostal, "T") & ","
                        CadValues2 = CadValues2 & DBSet(Rsx!prosocio, "T") & "," & DBSet(Rsx!nifSocio, "T") & ",'ES') "
                        
                        Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                        Sql = Sql & "ctabanc1, fecultco, impcobro, "
                        Sql = Sql & " text33csb, text41csb, text42csb, agente, iban, " ') "
                        Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
                        Sql = Sql & ") "
                    
                    Else
                        CadValues2 = CadValues2 & DBSet(vSocio.Banco, "N", "S") & "," & DBSet(vSocio.Sucursal, "N", "S") & ","
                        CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(vSocio.CuentaBan, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                        CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & ",1" ')"
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & ", " & DBSet(vSocio.Iban, "T", "S") & ") "
                        Else
                            CadValues2 = CadValues2 & ") "
                        End If
                    
        
                        'Insertamos en la tabla scobro de la CONTA
                        Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                        Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                        Sql = Sql & " text33csb, text41csb, agente" ') "
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            Sql = Sql & ", iban) "
                        Else
                            Sql = Sql & ") "
                        End If
                    End If
                        
                    
                    Sql = Sql & " VALUES " & CadValues2
                    ConnConta.Execute Sql
'[Monica]27/09/2011: quitamos toda la parte del else
'                Else
'                    '********** si la factura es negativa se inserta en la spago con valor poositivo
'                    CadValues2 = ""
'
'                    CadValuesAux2 = "('" & vSocio.CtaClien & "', " & DBSet(Rsx!NumFactu, "N") & ", '" & Format(Rsx!fecfactu, FormatoFecha) & "', "
'
'                    '------------------------------------------------------------
'
'                    CC = DBLet(vSocio.Digcontrol, "T")
'                    If DBLet(vSocio.Digcontrol, "T") = "**" Then CC = "00"
'
'                    i = 1
'                    CadValues2 = CadValuesAux2 & i
'                    CadValues2 = CadValues2 & ", " & DBSet(Rsx!Codforpa, "N") & ", " & DBSet(FecVen, "F") & ", "
'                    CadValues2 = CadValues2 & DBSet(DBLet(Rsx!TotalFac, "N") * (-1), "N") & ", " & DBSet(CtaBan, "T") & ","
'
'                    'David. Para que ponga la cuenta bancaria (SI LA tiene)
'                    CadValues2 = CadValues2 & DBSet(vSocio.Banco, "T", "S") & "," & DBSet(vSocio.Sucursal, "T", "S") & ","
'                    CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(vSocio.CuentaBan, "T", "S") & ","
'
'                    'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
'                    SQL = "Factura No.:" & DBLet(Rsx!NumFactu, "N")
'
'                    CadValues2 = CadValues2 & "'" & DevNombreSQL(SQL) & "',"
'
'                    SQL = " de " & Format(DBLet(Rsx!fecfactu, "F"), "dd/mm/yyyy")
'                    CadValues2 = CadValues2 & "'" & DevNombreSQL(SQL) & "')"
'
'                    'Grabar tabla spagop de la CONTABILIDAD
'                    '-------------------------------------------------
'                    If CadValues2 <> "" Then
'                        'Insertamos en la tabla spagop de la CONTA
'                        'David. Cuenta bancaria y descripcion textos
'                        SQL = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb) "
'                        SQL = SQL & " VALUES " & CadValues2
'                        ConnConta.Execute SQL
'                    End If
'                    '*******
'                End If
            End If
        End If
    
        B = True
    End If
    
EInsertarTesoreriaNew:
    If Err.Number <> 0 Then
        B = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaNewBOD = B
End Function





Private Function VariedadesFactura(cadenawhere As String) As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String

    On Error Resume Next
    

    Sql = "select distinct  nomvarie from rfactsoc_variedad INNER JOIN variedades ON rfactsoc_variedad.codvarie = variedades.codvarie "
    Sql = Sql & " where (rfactsoc_variedad.codtipom, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu) "
    Sql = Sql & " in (select codtipom, numfactu, fecfactu from rfactsoc where " & cadenawhere & ")"
     
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    cad = ""
    While Not Rs.EOF
        cad = cad & DBLet(Rs.Fields(0).Value, "T") & ","
    
        Rs.MoveNext
    Wend
    
    If cad <> "" Then ' quitamos la ultima coma
        cad = Mid(cad, 1, Len(cad) - 1)
    End If
    
    Set Rs = Nothing
    
    VariedadesFactura = cad
    
End Function


'----------------------------------------------------------------------
' FACTURAS ALMAZARA SOCIOS
'----------------------------------------------------------------------

Public Function PasarFacturaAlmzSoc(cadWHERE As String, FechaFin As String, FecRecep As Date, CtaRete As String, TotalFactura As Currency, FP As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura Socio
' ariagro.rcabfactalmz --> conta.cabfactprov
' ariagro.rlinfactalmz --> conta.linfactprov
'Actualizar la tabla ariagro.rcabfactalmz.contabilizada=1 para indicar que ya esta contabilizada
Dim B As Boolean
Dim cadMen As String
Dim Sql As String
Dim Mc As Contadores


    On Error GoTo EContab

    
    Set Mc = New Contadores
    
    
    CtaReten = CtaRete
    
    
    '---- Insertar en la conta Cabecera Factura
    B = InsertarCabFactAlmzSoc(cadWHERE, cadMen, Mc, CDate(FechaFin), FecRecep, TotalFactura, FP, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If B Then
        
        If B Then
            '---- Insertar lineas de Factura en la Conta
            If vParamAplic.ContabilidadNueva Then
                B = InsertarLinFactAlmzSocContaNueva("rcabfactalmz", cadWHERE, cadMen, Mc.Contador)
            Else
                B = InsertarLinFactAlmzSoc("rcabfactalmz", cadWHERE, cadMen, Mc.Contador)
            End If
            cadMen = "Insertando Lin. Factura Almazara Socio: " & cadMen
    
            If B Then
                If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaProv(vContaFra.NumeroFactura, vContaFra.Anofac)
            End If
    
            If B Then
                '---- Poner intconta=1 en ariges.scafac
                B = ActualizarCabFactAlmz("rcabfactalmz", cadWHERE, cadMen)
                cadMen = "Actualizando Factura Almazara Socio: " & cadMen
            End If
        End If
    End If
    
    
EContab:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Contabilizando Factura Socio", Err.Description
    End If
    If B Then
        PasarFacturaAlmzSoc = True
    Else
        PasarFacturaAlmzSoc = False
        If Not B Then
            Sql = "Insert into tmpErrFac(tipofichero,numfactu,fecfactu,codsocio,error) "
            Sql = Sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
            Sql = Sql & " WHERE " & Replace(cadWHERE, "rcabfactalmz", "tmpFactu")
            conn.Execute Sql
        End If
    End If
End Function


Private Function InsertarCabFactAlmzSoc(cadWHERE As String, cadErr As String, ByRef Mc As Contadores, FechaFin As Date, FecRec As Date, TotalFactura As Currency, FP As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Concepto As String

Dim TipoOpera As Integer

Dim Aux As String

Dim Sql2 As String
Dim CadenaInsertFaclin2 As String

    On Error GoTo EInsertar
       
    Sql = " SELECT fecfactu,year(" & DBSet(FecRec, "F") & ") as anofacpr," & DBSet(FecRec, "F") & ",numfactu,rsocios_seccion.codmacpro,"
    Sql = Sql & "baseimpo, tipoiva, porc_iva,imporiva,basereten,porc_ret,impreten,totalfac, rcabfactalmz.codsocio, rsocios.nomsocio, "
    Sql = Sql & " rsocios.dirsocio, rsocios.pobsocio, rsocios.codpostal, rsocios.prosocio, rsocios.nifsocio "
    Sql = Sql & " FROM (" & "rcabfactalmz "
    Sql = Sql & "INNER JOIN rsocios ON rcabfactalmz.codsocio=rsocios.codsocio) "
    Sql = Sql & " INNER JOIN rsocios_seccion ON rcabfactalmz.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N")
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
    
        If Mc.ConseguirContador("1", (FecRec <= CDate(FechaFin) - 365), True) = 0 Then
            
            vContaFra.NumeroFactura = Mc.Contador
            vContaFra.Anofac = DBLet(Rs!anofacpr)
            
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            BaseImp = DBLet(Rs!baseimpo, "N")
            TotalFac = BaseImp + DBLet(Rs!ImporIva, "N")
            AnyoFacPr = Rs!anofacpr
            
            ImpReten = DBLet(Rs!ImpReten, "N")
            
            TotalFactura = TotalFac - ImpReten
            
            FacturaSoc = DBLet(Rs!numfactu, "N")
            FecFactuSoc = DBLet(Rs!fecfactu, "F")
            
            CtaSocio = Rs!codmacpro
            
            '[Monica]29/07/2015: si es un asociado hay que seleccionar raiz de asociado + codigo de asociado
            If vParamAplic.Cooperativa = 0 Then
               Sql = "select rsocios.tiporelacion from rsocios where codsocio in (select codsocio from rcabfactalmz where " & cadWHERE & ")"
               If DevuelveValor(Sql) = 1 Then
                   
                   Sql = "select nroasociado from rsocios where codsocio in (select codsocio from rcabfactalmz where " & cadWHERE & ")"
                   Socio = DevuelveValor(Sql)
                   
                   Sql = "select raiz_cliente_asociado from rseccion where codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N")
                   CtaSocio = DevuelveValor(Sql) & Format(Socio, "00000")
               End If
            End If
            
            FecRecep = FecRec
            
            Concepto = "ALMAZARA ACEITE"
            
            Sql = ""
            Sql = Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & AnyoFacPr & "," & DBSet(FecRecep, "F") & "," & DBSet(FecRecep, "F") & ","
            Sql = Sql & DBSet(FacturaSoc, "T") & "," & DBSet(CtaSocio, "T") & "," & DBSet(Concepto, "T") & ","
            
            If vParamAplic.ContabilidadNueva Then
                Sql = Sql & DBSet(Rs!nomsocio, "T") & "," & DBSet(Rs!dirsocio, "T", "S") & ","
                Sql = Sql & DBSet(Rs!codpostal, "T", "S") & "," & DBSet(Rs!pobsocio, "T", "S") & "," & DBSet(Rs!prosocio, "T", "S") & ","
                Sql = Sql & DBSet(Rs!nifSocio, "F", "S") & ",'ES',"
                Sql = Sql & DBSet(FP, "N") & ","
            
                '$$$
                TipoOpera = 5 ' REA
                
                '[Monica]21/04/2017: antes tenia un 0 en Aux
                Aux = "X"
                If Rs!TotalFac < 0 Then Aux = "D"
                'codopera,codconce340,codintra
                Sql = Sql & TipoOpera & "," & DBSet(Aux, "T") & "," & ValorNulo & ","
                
                '[Monica]10/11/2016: en totalfac llevabamos base + impiva pq antes retencion estaba en lineas
                '                    en la nueva conta está en la cabecera
                TotalFac = TotalFac - ImpReten
                
                'para las lineas
                'factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
                'IVA 1, siempre existe
                Aux = "'" & SerieFraPro & "'," & Mc.Contador & "," & DBSet(Rs!FecRecep, "F") & "," & Rs!anofacpr & ","
                
                Sql2 = Aux & "1," & DBSet(BaseImp, "N") & "," & DBSet(Rs!TipoIVA, "N") & "," & DBSet(Rs!porc_iva, "N") & ","
                Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!ImporIva, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
                    
                'Los totales
                'totbases,totbasesret,totivas,totrecargo,totfacpr,
                Sql = Sql & DBSet(BaseImp, "N") & "," & DBSet(Rs!BaseReten, "N", "S") & ","
                'totivas
                Sql = Sql & DBSet(Rs!ImporIva, "N") & "," & DBSet(TotalFac, "N") & ","
                If DBLet(Rs!porc_ret, "N") <> 0 Then
                    Sql = Sql & DBSet(Rs!porc_ret, "N") & "," & DBSet(Rs!ImpReten, "N") & "," & DBSet(CtaReten, "T") & ",2" ' 2=retenciones agrícolas
                Else
                    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                End If
            
                vTipoIva(0) = Rs!TipoIVA
                vPorcIva(0) = Rs!porc_iva
                vPorcRec(0) = 0
                vImpIva(0) = Rs!ImporIva
                vImpRec(0) = 0
                vBaseIva(0) = Rs!BaseImp
                
                vTipoIva(1) = 0: vTipoIva(2) = 0
            
            
            
            
                'Insertar en la contabilidad
                Sql = "INSERT INTO factpro(numserie,numregis,fecfactu,anofactu,fecharec,fecliqpr,numfactu,codmacta,observa,nommacta,"
                Sql = Sql & "dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,codforpa,codopera,codconce340,codintra,"
                Sql = Sql & "totbases,totbasesret,totivas,totfacpr,retfacpr,trefacpr,cuereten,tiporeten)"
                Sql = Sql & " VALUES " & cad
                ConnConta.Execute Sql
            
                'Las  lineas de IVA
                Sql = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)"
                Sql = Sql & " VALUES " & CadenaInsertFaclin2
                ConnConta.Execute Sql
            
            
            Else
                Sql = Sql & DBSet(BaseImp, "N") & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & DBSet(Rs!porc_iva, "N") & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!ImporIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & DBSet(TotalFac, "N") & "," & DBSet(Rs!TipoIVA, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                cad = cad & "(" & Sql & ")"
                
                'Insertar en la contabilidad
                Sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,fecliqpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                Sql = Sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                Sql = Sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,nodeducible) "
                Sql = Sql & " VALUES " & cad
                ConnConta.Execute Sql
            End If
            
            
            'añadido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            Sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            Sql = Sql & ",'" & DevNombreSQL(FacturaSoc) & " @ " & Format(Rs!fecfactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!nomsocio) & "'," & Rs!Codsocio & ")"
            conn.Execute Sql
            
            FacturaSoc = DBLet(Rs!numfactu, "N")
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactAlmzSoc = False
        cadErr = Err.Description
    Else
        InsertarCabFactAlmzSoc = True
    End If
End Function

Public Function InsertarEnTesoreriaAlmz(MenError As String, Socio As Long, numfactu As String, fecfactu As Date, TotalTesor As Currency, FecVenci As Date, FecRecep As Date, ForpaPosi As Integer, ForpaNega As Integer, CtaBanco As String, LetraSerie As String) As Boolean
'Guarda datos de Tesoreria en tablas: spagop o scobro dependiendo del signo de la factura
Dim B As Boolean
Dim Sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim CC As String
Dim letraser As String
Dim Text33csb As String
Dim Text41csb As String
Dim Text42csb As String

Dim Rs As ADODB.Recordset

Dim BancoSoc As Integer
Dim SucurSoc As Integer
Dim DigcoSoc As String
Dim CtaBaSoc As String
Dim UltimaFactura As String
Dim Socio2 As Long

    On Error GoTo EInsertarTesoreriaAlmz

    InsertarEnTesoreriaAlmz = False
    B = False
    
    Sql = "select rsocios.nomsocio, rsocios.codbanco, rsocios.codsucur, rsocios.digcontr, rsocios.cuentaba, rsocios_seccion.codmacpro, rsocios.iban, "
    Sql = Sql & " rsocios.dirsocio,rsocios.pobsocio,rsocios.codpostal,rsocios.prosocio,rsocios.nifsocio "
    Sql = Sql & " from rsocios inner join rsocios_seccion on rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & vParamAplic.SeccionAlmaz
    Sql = Sql & " where rsocios.codsocio = " & DBSet(Socio, "N")

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    BancoSoc = 0
    SucurSoc = 0
    DigcoSoc = ""
    CtaBaSoc = ""
    CtaSocio = ""
    If Not Rs.EOF Then
        BancoSoc = DBLet(Rs!CodBanco, "N")
        SucurSoc = DBLet(Rs!CodSucur, "N")
        DigcoSoc = DBLet(Rs!digcontr, "T")
        CtaBaSoc = DBLet(Rs!CuentaBa, "T")
        IbanSoc = DBLet(Rs!Iban, "T")
       '[Monica]03/07/2013: añado trim(codmacta)
        CtaSocio = DBLet(Trim(Rs!codmacpro), "T")
            
        '[Monica]29/07/2015: si es un asociado hay que seleccionar raiz de asociado + codigo de asociado
        If vParamAplic.Cooperativa = 0 Then
           Sql = "select rsocios.tiporelacion from rsocios where codsocio = " & DBSet(Socio, "N")
           If DevuelveValor(Sql) = 1 Then
               
               Sql = "select nroasociado from rsocios where codsocio = " & DBSet(Socio, "N")
               Socio2 = DevuelveValor(Sql)
               
               Sql = "select raiz_cliente_asociado from rseccion where codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N")
               CtaSocio = DevuelveValor(Sql) & Format(Socio2, "00000")
           End If
        End If

'lo dejamos como estaba
'[Monica]27/09/2011: tanto si es positivo como si no se almacena en la cartera de cobros
        If TotalTesor > 0 Then ' se insertara en la cartera de pagos (spagop)
            CadValues2 = ""
        
            UltimaFactura = Mid(numfactu, Len(numfactu) - 6, 8)
            
            CadValuesAux2 = "("
            If vParamAplic.ContabilidadNueva Then CadValuesAux2 = CadValuesAux2 & DBSet(SerieFraPro, "T") & ","
            CadValuesAux2 = CadValuesAux2 & "'" & CtaSocio & "', " & DBSet(UltimaFactura, "T") & ", '" & Format(fecfactu, FormatoFecha) & "', "
        
            '------------------------------------------------------------
            I = 1
            CadValues2 = CadValuesAux2 & I
            CadValues2 = CadValues2 & ", " & ForpaPosi & ", '" & Format(FecVenci, FormatoFecha) & "', "
            CadValues2 = CadValues2 & DBSet(TotalTesor, "N") & ", " & DBSet(CtaBanco, "T") & ","
        
        
            If Not vParamAplic.ContabilidadNueva Then
                'David. Para que ponga la cuenta bancaria (SI LA tiene)
                CadValues2 = CadValues2 & DBSet(BancoSoc, "T", "S") & "," & DBSet(SucurSoc, "T", "S") & ","
                CadValues2 = CadValues2 & DBSet(DigcoSoc, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & ","
            End If
            
            'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
            Sql = "Almz.Nros:" & numfactu
                
            CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
            
            Sql = " de " & Format(fecfactu, "dd/mm/yyyy")
            CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" ')"
            
            If Not vParamAplic.ContabilidadNueva Then
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & ") "
                Else
                    CadValues2 = CadValues2 & ") "
                End If
            End If
        
            'Grabar tabla spagop de la CONTABILIDAD
            '-------------------------------------------------
            If CadValues2 <> "" Then
                'Insertamos en la tabla spagop de la CONTA
                'David. Cuenta bancaria y descripcion textos
                If vParamAplic.ContabilidadNueva Then
                
                    vvIban = MiFormat(IbanSoc, "") & MiFormat(CStr(BancoSoc), "0000") & MiFormat(CStr(SucurSoc), "0000") & MiFormat(DigcoSoc, "00") & MiFormat(CtaBaSoc, "0000000000")
                    
                    CadValues2 = CadValues2 & DBSet(vvIban, "T") & ","
                    'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                    CadValues2 = CadValues2 & DBSet(Rs!nomsocio, "T") & "," & DBSet(Rs!dirsocio, "T") & "," & DBSet(Rs!pobsocio, "T") & "," & DBSet(Rs!codposta, "T") & ","
                    CadValues2 = CadValues2 & DBSet(Rs!prosocio, "T") & "," & DBSet(Rs!nifSocio, "T") & ",'ES'),"
                    
                    Sql = "INSERT INTO pagos (codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, iban,"
                    Sql = Sql & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais)"
                    
                Else
                    Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        Sql = Sql & ", iban) "
                    Else
                        Sql = Sql & ") "
                    End If
                End If
                Sql = Sql & " VALUES " & CadValues2
                ConnConta.Execute Sql
            End If
'lo dejamos como estaba
'[Monica]27/09/2011: quitamos toda la parte del else
        Else
            ' si es negativo se inserta en positivo en la cartera de cobros (scobro)
            Text33csb = "'Almazara Nros:" & numfactu & "'"
            Text41csb = "de fecha " & Format(DBLet(fecfactu, "F"), "dd/mm/yyyy")
            Text42csb = "de " & DBSet((TotalTesor) * (-1), "N")

            CC = DBLet(DigcoSoc, "T")
            If DBLet(DigcoSoc, "T") = "**" Then CC = "00"

            UltimaFactura = Mid(numfactu, Len(numfactu) - 6, 8)

            '[Monica]03/07/2013: añado trim(codmacta)
            CadValuesAux2 = "(" & DBSet(LetraSerie, "T") & "," & DBSet(UltimaFactura, "N") & "," & DBSet(fecfactu, "F") & ", 1," & DBSet(Trim(CtaSocio), "T") & ","
            CadValues2 = CadValuesAux2 & DBSet(ForpaNega, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet((TotalTesor) * (-1), "N") & ","
            CadValues2 = CadValues2 & DBSet(CtaBanco, "T") & ","
            
            If vParamAplic.ContabilidadNueva Then
                CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & ","
                CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & "," & DBSet(Text42csb, "T") & ",1,"
                
                vvIban = MiFormat(IbanSoc, "") & MiFormat(CStr(BancoSoc), "0000") & MiFormat(CStr(SucurSoc), "0000") & MiFormat(CC, "00") & MiFormat(CtaBaSoc, "0000000000")
                
                CadValues2 = CadValues2 & DBSet(vvIban, "T") & ","
                'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                CadValues2 = CadValues2 & DBSet(Rs!nomsocio, "T") & "," & DBSet(Rs!dirsocio, "T") & "," & DBSet(Rs!pobsocio, "T") & "," & DBSet(Rs!codpostal, "T") & ","
                CadValues2 = CadValues2 & DBSet(Rs!prosocio, "T") & "," & DBSet(Rs!nifSocio, "T") & ",'ES'),"
            
            
                Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                Sql = Sql & "ctabanc1, fecultco, impcobro, "
                Sql = Sql & " text33csb, text41csb, text42csb, agente, iban, "
                Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
                Sql = Sql & ") "
            Else
                CadValues2 = CadValues2 & DBSet(BancoSoc, "N", "S") & "," & DBSet(SucurSoc, "N", "S") & ","
                CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & "," & DBSet(Text42csb, "T") & ",1" ')"
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & ") "
                Else
                    CadValues2 = CadValues2 & ") "
                End If
                
    
                'Insertamos en la tabla scobro de la CONTA
                Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                Sql = Sql & " text33csb, text41csb, text42csb, agente" ') "
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    Sql = Sql & ", iban) "
                Else
                    Sql = Sql & ") "
                End If
            
            End If
            Sql = Sql & " VALUES " & CadValues2
            ConnConta.Execute Sql
        End If

        B = True
    End If
    
    
EInsertarTesoreriaAlmz:
    
    If Err.Number <> 0 Then
        B = False
        MenError = "Error al insertar en Tesoreria de Almazara: " & Err.Description
    End If
    InsertarEnTesoreriaAlmz = B
End Function



Private Function InsertarLinFactAlmzSoc(cadTabla As String, cadWHERE As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String
Dim LineaVariedad As Integer

    On Error GoTo EInLinea
    

    Sql = " SELECT sum(rlinfactalmz.importel) as importe "
    Sql = Sql & " FROM rlinfactalmz "
    Sql = Sql & " WHERE " & Replace(cadWHERE, "rcabfactalmz", "rlinfactalmz")

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    If Not Rs.EOF Then
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        ImpLinea = Rs!Importe
        '----
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(vParamAplic.CtaGastosAlmz, "T")
        
        Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        Sql = Sql & ValorNulo ' centro de coste
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
    End If
    
    Rs.Close
    Set Rs = Nothing
    
    ' las retenciones si las hay
    If ImpReten <> 0 Then
        Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(CtaSocio, "T")
        Sql = Sql & "," & DBSet(ImpReten, "N") & ","
        Sql = Sql & ValorNulo ' no llevan centro de coste
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
        
        Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(CtaReten, "T")
        Sql = Sql & "," & DBSet(ImpReten * (-1), "N") & ","
        Sql = Sql & ValorNulo ' no llevan centro de coste
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
    End If
    
    
    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactAlmzSoc = False
        cadErr = Err.Description
    Else
        InsertarLinFactAlmzSoc = True
    End If
End Function

Private Function InsertarLinFactAlmzSocContaNueva(cadTabla As String, cadWHERE As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String
Dim LineaVariedad As Integer

    On Error GoTo EInLinea
    
    Sql = " SELECT sum(rlinfactalmz.importel) as importe "
    Sql = Sql & " FROM rlinfactalmz "
    Sql = Sql & " WHERE " & Replace(cadWHERE, "rcabfactalmz", "rlinfactalmz")

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    If Not Rs.EOF Then
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        ImpLinea = Rs!Importe
        '----
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        Sql = DBSet(SerieFraPro, "T") & "," & NumRegis & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(vParamAplic.CtaGastosAlmz, "T") & ","
        Sql = Sql & ValorNulo ' centro de coste
        Sql = Sql & "," & vTipoIva(0)
        Sql = Sql & "," & vPorcIva(0)
        Sql = Sql & "," & vPorcRec(0)
        Sql = Sql & "," & ImpLinea
        Sql = Sql & "," & vImpIva(0)
        Sql = Sql & "," & vImpRec(0)
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
    End If
    
    Rs.Close
    Set Rs = Nothing
    
    
    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO factpro_lineas (numserie,numregis,fecharec,anofactu,numlinea,codmacta,codccost,codigiva,porciva,porcrec,baseimpo,impoiva,imporec,aplicret) "
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactAlmzSocContaNueva = False
        cadErr = Err.Description
    Else
        InsertarLinFactAlmzSocContaNueva = True
    End If
End Function



Private Function ActualizarCabFactAlmz(cadTabla As String, cadWHERE As String, cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim Sql As String

    On Error GoTo EActualizar
    
    Sql = "UPDATE " & cadTabla & " SET contabilizado=1 "
    Sql = Sql & " WHERE " & cadWHERE

    conn.Execute Sql
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCabFactAlmz = False
        cadErr = Err.Description
    Else
        ActualizarCabFactAlmz = True
    End If
End Function


Public Function PasarFacturaAlmzCli(cadWHERE As String, CodCCost As String, LetraSerie As String, TotalFactura As Currency, FP As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariagro.rcabfactalmz --> conta.cabfact
' ariagro.rlinfactalmz --> conta.linfact
'Actualizar la tabla ariagro.rcabfactalmz.inconta=1 para indicar que ya esta contabilizada
Dim B As Boolean
Dim cadMen As String
Dim Sql As String

    On Error GoTo EContab

    
    'Insertar en la conta Cabecera Factura
    B = InsertarCabFactAlmzCli(cadWHERE, cadMen, LetraSerie, TotalFactura, FP, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If B Then
        CCoste = CodCCost
        'Insertar lineas de Factura en la Conta
        If vParamAplic.ContabilidadNueva Then
            B = InsertarLinFactAlmzCliContaNueva("rcabfactalmz", cadWHERE, cadMen, LetraSerie)
        Else
            B = InsertarLinFactAlmzCli("rcabfactalmz", cadWHERE, cadMen, LetraSerie)
        End If
        cadMen = "Insertando Lin. Factura Almazara Cliente: " & cadMen

        If B Then
            'Poner intconta=1 en ariagro.facturas
            B = ActualizarCabFactAlmz("rcabfactalmz", cadWHERE, cadMen)
            cadMen = "Actualizando Factura Almazara: " & cadMen
        End If
        
        If B Then
            If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
        End If
        
        
    End If
    
'    If Not b Then
'        Sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
'        Sql = Sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
'        Sql = Sql & " WHERE " & Replace(cadWhere, "facturas", "tmpFactu")
'        Conn.Execute Sql
'    End If
    
EContab:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If B Then
        PasarFacturaAlmzCli = True
    Else
        PasarFacturaAlmzCli = False
        
        Sql = "Insert into tmpErrFac(tipofichero,numfactu,fecfactu,codsocio,error) "
        Sql = Sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "rcabfactalmz", "tmpFactu")
        conn.Execute Sql
    End If
End Function


Private Function InsertarCabFactAlmzCli(cadWHERE As String, cadErr As String, LetraSerie As String, TotalFactura As Currency, FP As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Concepto As String
Dim cad As String

Dim Sql2 As String
Dim CadenaInsertFaclin2 As String



    On Error GoTo EInsertar
    
    Sql = Sql & " SELECT " & DBSet(LetraSerie, "T") & ",tipofichero,numfactu,fecfactu,rsocios_seccion.codmacpro,year(fecfactu) as anofaccl,"
    Sql = Sql & "baseimpo,tipoiva,porc_iva,imporiva,basereten, porc_ret, impreten, totalfac,  "
    Sql = Sql & "rsocios.nomsocio, rsocios.dirsocio,rsocios.pobsocio,rsocios.codpostal,rsocios.prosocio,rsocios.nifsocio"
    Sql = Sql & " FROM (" & "rcabfactalmz inner join " & "rsocios_seccion on rcabfactalmz.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & vParamAplic.SeccionAlmaz & ") "
    Sql = Sql & "INNER JOIN " & "rsocios ON rsocios_seccion.codsocio=rsocios.codsocio "
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
        vContaFra.NumeroFactura = DBLet(Rs!numfactu)
        vContaFra.Serie = LetraSerie
        vContaFra.Anofac = DBLet(Rs!anofaccl)
    
        BaseImp = Rs!baseimpo
        '---- Laura 10/10/2006:  añadir el totalfac para utilizarlo en insertar lineas
        TotalFac = Rs!TotalFac
        '----
        
        TotalFactura = TotalFac ' sacamos el importe total fuera para tesoreria
        
        Concepto = "ALMAZARA "
        If DBLet(Rs!tipofichero, "N") = 0 Then
            Concepto = Concepto & "ACEITE"
        Else
            Concepto = Concepto & "STOCK"
        End If
        
        CtaSocio = Rs!codmacpro
        
        '[Monica]29/07/2015: si es un asociado hay que seleccionar raiz de asociado + codigo de asociado
        If vParamAplic.Cooperativa = 0 Then
           Sql = "select rsocios.tiporelacion from rsocios where codsocio in (select codsocio from rcabfactalmz where " & cadWHERE & ")"
           If DevuelveValor(Sql) = 1 Then
               
               Sql = "select nroasociado from rsocios where codsocio in (select codsocio from rcabfactalmz where " & cadWHERE & ")"
               Socio = DevuelveValor(Sql)
               
               Sql = "select raiz_cliente_asociado from rseccion where codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N")
               CtaSocio = DevuelveValor(Sql) & Format(Socio, "00000")
           End If
        End If
        
        Sql = ""
        Sql = "'" & LetraSerie & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(CtaSocio, "T") & "," & Year(Rs!fecfactu) & "," & DBSet(Concepto, "T") & ","
        
        
        If vParamAplic.ContabilidadNueva Then
            ' para el caso de las rectificativas
            Dim vTipM As String
            vTipM = DevuelveValor("select codtipom from stipom where letraser = " & DBSet(Rs!letraser, "T"))
            If vTipM = "FAR" Then
                Sql = Sql & "'D',"
            Else
                Sql = Sql & "'0',"
            End If
            
            Sql = Sql & "0," & DBSet(FP, "N") & "," & DBSet(Rs!baseimpo, "N") & "," & ValorNulo & "," & DBSet(Rs!ImporIva, "N") & ","
            Sql = Sql & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,"
            Sql = Sql & DBSet(Rs!nomsocio, "T") & "," & DBSet(Rs!dirsocio, "T") & "," & DBSet(Rs!pobsocio, "T") & "," & DBSet(Rs!codpostal, "T") & ","
            Sql = Sql & DBSet(Rs!prosocio, "T") & "," & DBSet(Rs!nifSocio, "T") & ",'ES',1"
        
            cad = cad & "(" & Sql & ")"
        
        
            Sql = "INSERT INTO factcli (numserie,numfactu,fecfactu,fecliqcl,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
            Sql = Sql & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
            Sql = Sql & "codpais,codagente)"
            Sql = Sql & " VALUES " & cad
            ConnConta.Execute Sql
    '***
            CadenaInsertFaclin2 = ""
                
            
            'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
            'IVA 1, siempre existe
            Sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
            Sql2 = Sql2 & "1," & DBSet(Rs!baseimpo, "N") & "," & Rs!TipoIVA & "," & DBSet(Rs!porc_iva, "N") & ","
            Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!ImporIva, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
        
            Sql = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
            Sql = Sql & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
            ConnConta.Execute Sql
            
            'para las lineas
            vTipoIva(0) = Rs!TipoIVA
            vPorcIva(0) = Rs!porc_iva
            vPorcRec(0) = 0
            vImpIva(0) = Rs!ImporIva
            vImpRec(0) = 0
            vBaseIva(0) = Rs!baseimpo
            
            vTipoIva(1) = 0: vTipoIva(2) = 0
            
            
            
            
        Else
            Sql = Sql & DBSet(Rs!baseimpo, "N") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!porc_iva, "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!ImporIva, "N", "N") & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!TipoIVA, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
        
            cad = cad & "(" & Sql & ")"
        
            'Insertar en la contabilidad
            Sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,fecliqcl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
            Sql = Sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
            Sql = Sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien) "
            Sql = Sql & " VALUES " & cad
            ConnConta.Execute Sql
        
        End If
'        RS.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactAlmzCli = False
        cadErr = Err.Description
    Else
        InsertarCabFactAlmzCli = True
    End If
End Function


Private Function InsertarLinFactAlmzCli(cadTabla As String, cadWHERE As String, cadErr As String, LetraSerie As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String

    On Error GoTo EInLinea
    

    Sql = " SELECT " & DBSet(LetraSerie, "T") & ",rlinfactalmz.numfactu,rlinfactalmz.fecfactu,rlinfactalmz.codsocio," & vParamAplic.CtaVentasAlmz & ",sum(importel) as importe "
    Sql = Sql & " FROM rlinfactalmz "
    Sql = Sql & " WHERE " & Replace(cadWHERE, "rcabfactalmz", "rlinfactalmz")
    Sql = Sql & " GROUP BY 1,2,3,4,5 "
    Sql = Sql & " order by 1,2,3,4,5 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        totimp = totimp + DBLet(Rs!Importe, "N")
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        Sql = "'" & LetraSerie & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
        Sql = Sql & DBSet(vParamAplic.CtaVentasAlmz, "T")
        
        Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        Sql = Sql & "," & DBSet(Rs!Importe, "N") & ","
        
        Sql = Sql & ValorNulo ' centro de coste
        
        
        cad = cad & "(" & Sql & ")" & ","
        
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
' siempre cuadrará
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)
        Sql2 = Sql2 & DBSet(totimp, "N") & ","
        If CCoste = "" Then
            Sql2 = Sql2 & ValorNulo
        Else
            Sql2 = Sql2 & DBSet(CCoste, "T")
        End If
        If SQLaux <> "" Then 'hay mas de una linea
            cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If


    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactAlmzCli = False
        cadErr = Err.Description
    Else
        InsertarLinFactAlmzCli = True
    End If
End Function


Private Function InsertarLinFactAlmzCliContaNueva(cadTabla As String, cadWHERE As String, cadErr As String, LetraSerie As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String

    On Error GoTo EInLinea
    

    Sql = " SELECT " & DBSet(LetraSerie, "T") & ",rlinfactalmz.numfactu,rlinfactalmz.fecfactu,rlinfactalmz.codsocio," & vParamAplic.CtaVentasAlmz & ",sum(importel) as importe "
    Sql = Sql & " FROM rlinfactalmz "
    Sql = Sql & " WHERE " & Replace(cadWHERE, "rcabfactalmz", "rlinfactalmz")
    Sql = Sql & " GROUP BY 1,2,3,4,5 "
    Sql = Sql & " order by 1,2,3,4,5 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        totimp = totimp + DBLet(Rs!Importe, "N")
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        Sql = "'" & LetraSerie & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
        Sql = Sql & DBSet(vParamAplic.CtaVentasAlmz, "T")
        Sql = Sql & ValorNulo ' centro de coste
        Sql = Sql & "," & DBSet(Rs!fecfactu, "F")
        Sql = Sql & "," & DBSet(vTipoIva(0), "N")
        Sql = Sql & "," & DBSet(vPorcIva(0), "N")
        Sql = Sql & "," & DBSet(vPorcRec(0), "N")
        Sql = Sql & "," & DBSet(Rs!Importe, "N")
        Sql = Sql & "," & DBSet(vImpIva(0), "N")
        Sql = Sql & "," & DBSet(vImpRec(0), "N")
        
        
        cad = cad & "(" & Sql & ")" & ","
        
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
' siempre cuadrará
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
'    If totimp <> BaseImp Then
''        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
'        'en SQL esta la ult linea introducida
'        totimp = BaseImp - totimp
'        totimp = ImpLinea + totimp '(+- diferencia)
'        Sql2 = Sql2 & DBSet(totimp, "N") & ","
'        If CCoste = "" Then
'            Sql2 = Sql2 & ValorNulo
'        Else
'            Sql2 = Sql2 & DBSet(CCoste, "T")
'        End If
'        If SQLaux <> "" Then 'hay mas de una linea
'            Cad = SQLaux & "(" & Sql2 & ")" & ","
'        Else 'solo una linea
'            Cad = "(" & Sql2 & ")" & ","
'        End If
'
''        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
''        cad = Replace(cad, SQL, Aux)
'    End If


    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO factcli_lineas (numserie,numfactu,anofactu,numlinea,codmacta,codccost,fecfactu,codigiva,porciva,porcrec,baseimpo,impoiva,imporec) "
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactAlmzCliContaNueva = False
        cadErr = Err.Description
    Else
        InsertarLinFactAlmzCliContaNueva = True
    End If
End Function





'??????????????
'?????????????? POZOS
'??????????????

Public Function InsertarEnTesoreriaPOZOS(MenError As String, ByRef RS1 As ADODB.Recordset, FecVenci As Date, Forpa As String, CtaBanco As String) As Boolean
'Guarda datos de Tesoreria en tablas: spagop o scobro dependiendo del signo de la factura
Dim B As Boolean
Dim Sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim CC As String
Dim letraser As String

Dim Rs4 As ADODB.Recordset
Dim Sql4 As String

Dim Rs6 As ADODB.Recordset
Dim Sql6 As String
Dim Sql5 As String

Dim Text33csb As String
Dim Text41csb As String
Dim Text42csb As String
Dim Text43csb As String
Dim Text51csb As String
Dim Text52csb As String
Dim Text53csb As String
Dim Text61csb As String
Dim Text62csb As String
Dim Text63csb As String
Dim Text71csb As String
Dim Text72csb As String
Dim Text73csb As String
Dim Text81csb As String
Dim Text82csb As String
Dim Text83csb As String

Dim Partida As String
Dim hanegada As Currency

Dim Total_1 As Currency
Dim Total_2 As Currency
Dim ImpIva_1 As Currency
Dim ImpIva_2 As Currency
Dim TTotal_1 As Currency
Dim TTotal_2 As Currency

Dim Rs As ADODB.Recordset

Dim BancoSoc As Integer
Dim SucurSoc As Integer
Dim DigcoSoc As String
Dim CtaBaSoc As String
Dim UltimaFactura As String
Dim LetraSerie As String

Dim Accion1 As Currency
Dim Accion2 As Currency
Dim Accion3 As Currency

Dim TotalFact As Currency
Dim Hidrantes As String
Dim Hidrantes2 As String
Dim Hidrantes3 As String

Dim Brazas As Long
Dim v_hanegada As Long
Dim v_brazas As Currency

Dim J As Integer
Dim k As Integer
Dim vPorcen As String

Dim Referencia As String

Dim BaseImp As Currency
Dim PorcIva As Currency

Dim CadValues As String
Dim Base As Currency

Dim Cad1 As String
Dim cad As String
            
    On Error GoTo EInsertarTesoreriaPOZ

    InsertarEnTesoreriaPOZOS = False
    B = False
    
    Text71csb = ""
    Text72csb = ""
    Text82csb = ""
    
    Sql = "select rsocios.nomsocio, rsocios.codbanco, rsocios.codsucur, rsocios.digcontr, rsocios.cuentaba, rsocios_seccion.codmaccli, rsocios.nifsocio, "
    '[Monica]03/08/2012: añadimos los datos fiscales a la scobro
    Sql = Sql & " rsocios.dirsocio, rsocios.pobsocio, rsocios.prosocio, rsocios.codpostal, rsocios.iban "
    Sql = Sql & " from rsocios inner join rsocios_seccion on rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & vParamAplic.SeccionPOZOS
    Sql = Sql & " where rsocios.codsocio = " & DBSet(RS1!Codsocio, "N")

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    BancoSoc = 0
    SucurSoc = 0
    DigcoSoc = ""
    CtaBaSoc = ""
    CtaSocio = ""
    If Not Rs.EOF Then
        BancoSoc = DBLet(Rs!CodBanco, "N")
        SucurSoc = DBLet(Rs!CodSucur, "N")
        DigcoSoc = DBLet(Rs!digcontr, "T")
        CtaBaSoc = DBLet(Rs!CuentaBa, "T")
        IbanSoc = DBLet(Rs!Iban, "T")
        
        '[Monica]03/07/2013: añado trim(codmacta)
        CtaSocio = Trim(DBLet(Rs!codmaccli, "T"))
        
        LetraSerie = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(RS1!CodTipom, "T"))
        
        '09/09/2010: el total factura ahora es la suma de todos los recibos cuando son de consumo
        Sql5 = "select sum(totalfact) from rrecibpozos where codtipom = " & DBSet(RS1!CodTipom, "T")
        Sql5 = Sql5 & " and numfactu = " & DBSet(RS1!numfactu, "N")
        Sql5 = Sql5 & " and fecfactu = " & DBSet(RS1!fecfactu, "F")
        
        TotalFact = DevuelveValor(Sql5)
        
        Select Case DBLet(RS1!CodTipom, "T")
            '[Monica]02/02/2016: contabilizacion de las facturas de quatretonda
            Case "FIN" ' factura interna
                Hidrantes = ""
                Sql6 = "select hidrante from rrecibpozos where codtipom = " & DBSet(RS1!CodTipom, "T")
                Sql6 = Sql6 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                Sql6 = Sql6 & " and fecfactu = " & DBSet(RS1!fecfactu, "F")
            
                Set Rs6 = New ADODB.Recordset
                Rs6.Open Sql6, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not Rs6.EOF
                    Hidrantes = Hidrantes & Trim(DBLet(Rs6!Hidrante, "T")) & " "
                    Rs6.MoveNext
                Wend
                Set Rs6 = Nothing
            
                If vParamAplic.Cooperativa = 7 Then
                    BaseImp = DevuelveValor("select sum(baseimpo) from rrecibpozos where codtipom = " & DBSet(RS1!CodTipom, "T") & _
                                                   " and numfactu = " & DBSet(RS1!numfactu, "N") & _
                                                   " and fecfactu = " & DBSet(RS1!fecfactu, "F"))
                    PorcIva = DevuelveValor("select porc_iva from rrecibpozos where codtipom = " & DBSet(RS1!CodTipom, "T") & _
                                                   " and numfactu = " & DBSet(RS1!numfactu, "N") & _
                                                   " and fecfactu = " & DBSet(RS1!fecfactu, "F"))
                    TotalFact = Round2(BaseImp * (1 + (PorcIva / 100)), 2)
                
                End If
            
                Text33csb = "** Recibo de Consumo POZOS **" 'POZOS Nros:" & DBLet(RS1!numfactu, "N") & "'"
                Text41csb = "Factura Interna: " & Format(DBLet(RS1!numfactu, "N"), "0000000") & " de fecha " & Format(DBLet(RS1!fecfactu, "F"), "dd/mm/yyyy")
    
                Referencia = ""
        
            
            Case "RCP" ' recibos de consumo
                '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
                If vParamAplic.Cooperativa <> 8 And vParamAplic.Cooperativa <> 10 Then ' Mallaes y Quatretonda
                    Hidrantes = ""
                    Sql6 = "select hidrante from rrecibpozos where codtipom = " & DBSet(RS1!CodTipom, "T")
                    Sql6 = Sql6 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                    Sql6 = Sql6 & " and fecfactu = " & DBSet(RS1!fecfactu, "F")
                
                    Set Rs6 = New ADODB.Recordset
                    Rs6.Open Sql6, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    While Not Rs6.EOF
                        Hidrantes = Hidrantes & Trim(DBLet(Rs6!Hidrante, "T")) & " "
                        Rs6.MoveNext
                    Wend
                    Set Rs6 = Nothing
                
                    If vParamAplic.Cooperativa = 7 Then
                        BaseImp = DevuelveValor("select sum(baseimpo) from rrecibpozos where codtipom = " & DBSet(RS1!CodTipom, "T") & _
                                                       " and numfactu = " & DBSet(RS1!numfactu, "N") & _
                                                       " and fecfactu = " & DBSet(RS1!fecfactu, "F"))
                        PorcIva = DevuelveValor("select porc_iva from rrecibpozos where codtipom = " & DBSet(RS1!CodTipom, "T") & _
                                                       " and numfactu = " & DBSet(RS1!numfactu, "N") & _
                                                       " and fecfactu = " & DBSet(RS1!fecfactu, "F"))
                        TotalFact = Round2(BaseImp * (1 + (PorcIva / 100)), 2)
                    
                    End If
                
                
                    Text33csb = "** Recibo de Consumo POZOS **" 'POZOS Nros:" & DBLet(RS1!numfactu, "N") & "'"
                    Text41csb = "FACTURA : " & Format(DBLet(RS1!numfactu, "N"), "0000000") & " de fecha " & Format(DBLet(RS1!fecfactu, "F"), "dd/mm/yyyy")
                    
                    If vParamAplic.Cooperativa <> 7 Then
                    
                        Text42csb = "CONTADORES : "
                        
                        If Len(Hidrantes) > 27 Then
                            J = InStr(1, Hidrantes, " ")
                            Hidrantes2 = Mid(Hidrantes, 1, J - 1)
                            k = InStr(J + 1, Hidrantes, " ")
                            While Len(Hidrantes2) < 27 And k <> 0
                                Hidrantes3 = Hidrantes2
                                Hidrantes2 = Hidrantes2 & " " & Mid(Hidrantes, J + 1, k - J - 1)
                                J = k
                                k = InStr(J + 1, Hidrantes, " ")
                            Wend
                            If Len(Hidrantes2) > 27 Then
                                Text42csb = "CONTADORES : " & Hidrantes3
                                ' el resto de cadena lo meto en la linea de abajo
                                Text51csb = Mid(Hidrantes, Len(Hidrantes3) + 2, Len(Hidrantes))
                            Else
                                Text42csb = "CONTADORES : " & Hidrantes2
                                ' el resto de cadena lo meto en la linea de abajo
                                Text51csb = Mid(Hidrantes, Len(Hidrantes2) + 2, Len(Hidrantes))
                            End If
                        Else
                            Text42csb = "CONTADORES : " & Hidrantes
                        End If
                    End If
                    Referencia = ""
                    
                Else
                    ' rellenamos un recibo por consumo UTXERA y Escalona
                    Sql6 = "select * from rrecibpozos where codtipom = " & DBSet(RS1!CodTipom, "T")
                    Sql6 = Sql6 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                    Sql6 = Sql6 & " and fecfactu = " & DBSet(RS1!fecfactu, "F")
                    
                    Set Rs6 = New ADODB.Recordset
                    Rs6.Open Sql6, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    If vParamAplic.Cooperativa = 8 Then 'Utxera
                       '[Monica]29/01/2014: cambios text33csb(80) text41csb(60)
                        Text33csb = ""
                        Text41csb = ""
                        cad = Mid(Year(DBLet(RS1!fecfactu, "F")), 3, 2) & "RCP" & Format(DBLet(RS1!numfactu, "N"), "0000000") & " Cont:" & Format(CLng(DBLet(Rs6!Hidrante, "T")), "00000")
                        cad = cad & " Soc:" & Format(DBLet(RS1!Codsocio, "N"), "000000") & " " & Mid(DBLet(Rs6!Concepto), 1, 15) & " Pol/Par:" & Mid(Trim(DBLet(Rs6!Poligono, "T")), 1, 2) & "/" & DBLet(Rs6!parcelas)

                        If Len(cad) > 80 Then cad = Mid(cad, 1, 78) & ".."
                        Text33csb = cad

                        cad = "Lec:" & Format(DBLet(Rs6!fech_act, "F"), "dd-mm-yy") & " " & Format(DBLet(Rs6!Consumo1, "N"), "000000") & " m³ Pr:" & Format(DBLet(Rs6!Precio1, "N"), "0.00") & " €/m³ Total: " & Format(DBLet(Rs6!TotalFact, "N"), "###,##0.00")
                        Text41csb = cad

                        '[Monica]20/02/2014: en utxera tb grabamos el codigo de socio
                        'Referencia = DBLet(Rs6!Hidrante, "T")
                        Referencia = Format(DBLet(RS1!Codsocio, "N"), "000000")
                    Else ' Escalona
                       '[Monica]29/01/2014: cambios text33csb(80) text41csb(60)
                       
                       
                       '[Monica]20/06/2014: cambiamos lo que imprimimos en los textos (quitamos socio y añadimos fecha de lectura anterior
                       '                    los mismos cambios para utxera
                       
                        Text33csb = ""
                        Text41csb = ""
                        cad = Mid(Year(DBLet(RS1!fecfactu, "F")), 3, 2) & "RCP" & Format(DBLet(RS1!numfactu, "N"), "0000000") & " Cont:" & Format(CLng(DBLet(Rs6!Hidrante, "T")), "00000")
                        cad = cad & " Pol/Par:" & Mid(Trim(DBLet(Rs6!Poligono, "T")), 1, 2) & "/" & Mid(Trim(DBLet(Rs6!parcelas)), 1, 20) & " Lec.ant:" & Format(DBLet(Rs6!lect_ant, "N"), "000000000")
                        
'                        If Len(Cad) > 80 Then Cad = Mid(Cad, 1, 78) & ".."

                        Text33csb = cad
                        
                        Dim longitud As Integer
                        longitud = Len(cad)
                        
                        Text33csb = Text33csb & Space(80 - longitud)
                        
                        cad = "Le.ac:" & Format(DBLet(Rs6!lect_act, "N"), "000000000") & " Con:" & Format(DBLet(Rs6!Consumo1, "N"), "000000") & " Pr:" & Format(DBLet(Rs6!Precio1, "N"), "#0.00") & "€/m³ Tot:" & Format(DBLet(Rs6!TotalFact, "N") - DBLet(Rs6!imprecargo, "N"), "#####0.00")
                        '[Monica]15/01/2016: si hay recargo lo especifico
                        If DBLet(Rs6!imprecargo, "N") <> 0 Then
                            cad = cad & "+" & Format(DBLet(Rs6!imprecargo, "N"), "##0.00")
                        End If
                        Text41csb = cad
                        
                        longitud = Len(cad)
                        Text41csb = Text41csb & Space(60 - longitud)
                        
                        Referencia = Format(DBLet(RS1!Codsocio, "N"), "000000")
                    End If
                End If

        
            Case "RMP" ' recibos de mantenimiento
                Text33csb = "** Recibo de Mantenimiento **"
                
                Sql6 = "select * from rrecibpozos where codtipom = " & DBSet(RS1!CodTipom, "T")
                Sql6 = Sql6 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                Sql6 = Sql6 & " and fecfactu = " & DBSet(RS1!fecfactu, "F")
                
                Set Rs6 = New ADODB.Recordset
                Rs6.Open Sql6, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not Rs6.EOF Then
                
                    If vParamAplic.Cooperativa = 8 Then
'[Monica]29/01/2014: quitamos esto, sutituimos por lo de abajo
'                        Text41csb = "FACTURA: " & Format(DBLet(RS1!numfactu, "N"), "#######") & " DE FECHA " & Format(DBLet(RS1!fecfactu, "N"), "dd/mm/yyyy")
'
'                        Text42csb = "CONCEPTO: "
'
'                        Text51csb = RecuperaValor(ParteCadena(DBLet(Rs6!Concepto, "T"), 3, 40), 1)
'                        Text53csb = RecuperaValor(ParteCadena(DBLet(Rs6!Concepto, "T"), 3, 40), 2)
'                        Text62csb = RecuperaValor(ParteCadena(DBLet(Rs6!Concepto, "T"), 3, 40), 3)
'
'                        Text43csb = "CONTADOR: " & DBLet(Rs6!Hidrante, "T")
'
                        Sql4 = "select hanegada from rpozos where hidrante= " & DBSet(Rs6!Hidrante, "T")
                        Sql4 = Sql4 & " and fechabaja is null"

                        hanegada = DevuelveValor(Sql4)
                        'Brazas = (Int(Hanegada) * 200) + (Hanegada - Int(Hanegada)) * 1000
                        v_hanegada = Int(hanegada)
                        v_brazas = (hanegada - Int(hanegada)) * 200
'
'                        Text52csb = "Importe  : " & Round2(DBLet(Rs6!TotalFact, "N"), 2)
'                        Text61csb = "Hanegadas: " & Format(v_hanegada, "#####0") & "    Brazas: " & Format(v_brazas, "#####0.00")
'
'                        Text63csb = ""
'                        vPorcen = DevuelveDesdeBDNew(cAgro, "rpozos_cooprop", "porcentaje", "hidrante", Rs6!Hidrante, "T", , "codsocio", RS1!Codsocio, "N")
'                        If vPorcen <> "" Then
'                            Text63csb = "Porcentaje Participacion " & vPorcen & "%"
'                        End If
                        
                        '[Monica]29/01/2014: cambios text33csb(80) text41csb(60)
                        Text33csb = ""
                        Text41csb = ""
                        cad = Mid(Year(DBLet(RS1!fecfactu, "F")), 3, 2) & "RMP" & Format(DBLet(RS1!numfactu, "N"), "0000000")
                        
                        '[Monica]29/04/2014: grabamos las hanegadas y las brazas en lugar de "Precios según información enviada"
                        cad = cad & " Soc:" & Format(DBLet(RS1!Codsocio, "N"), "000000") & " " & Mid(DBLet(Rs6!Concepto), 1, 20) & " " & Format(v_hanegada, "#####0") & "hg " & Format(v_brazas, "#####0") & "br a " & DBSet(Rs6!Precio, "N") & "Eur" ' " Precios según información enviada"
                         
                        Text33csb = cad
                         
                        cad = "Tot:" & Format(DBLet(Rs6!TotalFact, "N"), "####0.00") & " "
                        cad = cad & DBLet(Rs6!Hidrante, "T")
                        
                        Text41csb = cad
                        
                        '[Monica]20/02/2014: grabamos el codigo de socio en lugar del hidrante
                        'Referencia = DBLet(Rs6!Hidrante, "T")
                        Referencia = Format(DBLet(RS1!Codsocio, "N"), "000000")
                    Else
                        '[Monica]10/05/2012: añadida Escalona que funciona como Utxera
                        If vParamAplic.Cooperativa = 10 Then
                            Text42csb = ""
                            Text51csb = ""
                            Text53csb = ""
                            Text62csb = ""
                            Text71csb = ""
                            Text73csb = ""
                            Text82csb = ""
                            Text41csb = ""
                            Text43csb = ""
                            Text52csb = ""
                            Text61csb = ""
                            Text63csb = ""
                            Text72csb = ""
                            Text81csb = ""
                            
'                            Text33csb = RecuperaValor(ParteCadena(DBLet(Rs6!Concepto, "T"), 3, 40), 1)
                            
                            '[Monica]29/01/2014: cambios text33csb(80) text41csb(60)
                            Text33csb = ""
                            Text41csb = ""
                            cad = Mid(Year(DBLet(RS1!fecfactu, "F")), 3, 2) & "RMP" & Format(DBLet(RS1!numfactu, "N"), "0000000")
                            cad = cad & " Soc:" & Format(DBLet(RS1!Codsocio, "N"), "000000") & " " & Mid(DBLet(Rs6!Concepto), 1, 20) & " Precios según información enviada"
                             
                            Text33csb = cad
                             
                            cad = "Tot:" & Format(DBLet(Rs6!TotalFact, "N") - DBLet(Rs6!imprecargo, "N"), "####0.00")
                            '[Monica]15/01/2016: metemos el recargo
                            If DBLet(Rs6!imprecargo, "N") <> 0 Then
                                cad = cad & " Rec:" & Format(DBLet(Rs6!imprecargo, "N"), "###0.00")
                            End If
                            cad = cad & " "
                            
                            
                            Sql4 = "select rpartida.nomparti, rpozos.poligono, rpozos.parcelas, rrecibpozos_hid.hanegada, rrecibpozos.precio, "
                            Sql4 = Sql4 & " rrecibpozos.porcdto, rrecibpozos.impdto, rrecibpozos.totalfact "
                            Sql4 = Sql4 & " from rpozos, rpartida, rrecibpozos_hid, rrecibpozos "
                            Sql4 = Sql4 & " where rpozos.codparti = rpartida.codparti "
                            Sql4 = Sql4 & " and rpozos.hidrante = rrecibpozos_hid.hidrante "
                            Sql4 = Sql4 & " and rrecibpozos_hid.codtipom = " & DBSet(Rs6!CodTipom, "T")
                            Sql4 = Sql4 & " and rrecibpozos_hid.numfactu = " & DBSet(Rs6!numfactu, "N")
                            Sql4 = Sql4 & " and rrecibpozos_hid.fecfactu = " & DBSet(Rs6!fecfactu, "F")
                            Sql4 = Sql4 & " and rrecibpozos_hid.codtipom = rrecibpozos.codtipom"
                            Sql4 = Sql4 & " and rrecibpozos_hid.numfactu = rrecibpozos.numfactu"
                            Sql4 = Sql4 & " and rrecibpozos_hid.fecfactu = rrecibpozos.fecfactu"
                            
                            Set Rs4 = New ADODB.Recordset
                            Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                            
                            I = 0
                            While Not Rs4.EOF And I <= 6 '15
                                I = I + 1
'[Monica]29/01/2014: sustituido por
'                                hanegada = DBLet(DBLet(Rs4!hanegada, "N"))
'                                'Brazas = (Int(Hanegada) * 200) + (Hanegada - Int(Hanegada)) * 1000
'                                v_hanegada = Int(hanegada)
'                                v_brazas = (hanegada - Int(hanegada)) * 200
'
'                                CadValues = Mid(Rs4!nomparti, 1, 15) & " " & Format(DBLet(Rs4!Poligono, "N"), "##0") & " " & Format(DBLet(Rs4!parcelas, "N"), "####0") & " " & Format(v_hanegada, "##0") & " " & Format(v_brazas, "###0") & " " & Format(DBLet(Rs4!Precio, "N"), "##0.0000")
'
'                                Select Case i
'                                    Case 1
'                                        Text42csb = CadValues
'                                    Case 2
'                                        Text51csb = CadValues
'                                    Case 3
'                                        Text53csb = CadValues
'                                    Case 4
'                                        Text62csb = CadValues
'                                    Case 5
'                                        Text71csb = CadValues
'                                    Case 6
'                                        Text73csb = CadValues
'                                    Case 7
'                                        Text82csb = CadValues
'                                    Case 8
'                                        Text41csb = CadValues
'                                    Case 9
'                                        Text43csb = CadValues
'                                    Case 10
'                                        Text52csb = CadValues
'                                    Case 11
'                                        Text61csb = CadValues
'                                    Case 12
'                                        Text63csb = CadValues
'                                    Case 13
'                                        Text72csb = CadValues
'                                    Case 14
'                                        Text81csb = CadValues
'                                End Select

                                If I > 1 Then cad = cad & "/"

                                cad = cad & Format(CLng(DBLet(Rs6!Hidrante, "T")), "00000")
                                
                                Rs4.MoveNext
                            Wend
                            Text41csb = cad
'[Monica]29/01/2014: quitado
'                            If i > 14 Then Text81csb = "y otros"
'
'                            Text83csb = ""
'                            If DBLet(Rs6!PorcDto) <> 0 Then
'                                If DBLet(Rs6!PorcDto) < 0 Then
'                                    Base = Rs6!TotalFact + Rs6!ImpDto
'                                    CadValues = "Recargo " & Format(Base, "###,##0.00") & " " & Format(Abs(Rs6!PorcDto), "##0.00") & " " & Format(Rs6!TotalFact, "###,##0.00")
'                                Else
'                                    Base = Rs6!TotalFact + Rs6!ImpDto
'                                    CadValues = "Bonificacion " & Format(Base, "###,##0.00") & " " & Format(Rs6!PorcDto, "##0.00") & " " & Format(Rs6!TotalFact, "###,##0.00")
'                                End If
'                                Text83csb = CadValues
'                            End If
                            
                            '[Monica]08/01/2014: para el caso de escalona lo cambiamos para que imprima en referencia
                            '                    el codigo de socio con formato
                            Referencia = Format(DBLet(RS1!Codsocio, "N"), "000000")

                        
                        Else
                            Text41csb = "FACTURA: " & Format(DBLet(RS1!numfactu, "N"), "#######") & " DE FECHA " & Format(DBLet(RS1!fecfactu, "N"), "dd/mm/yyyy")
                            Text42csb = "CONCEPTO: " & DBLet(Rs6!Concepto, "T")
                            Text43csb = ""
                            
                            Sql4 = "select rsocios_pozos.numfases, rsocios_pozos.acciones from rsocios_pozos  "
                            Sql4 = Sql4 & " where rsocios_pozos.codsocio = " & DBSet(Rs6!Codsocio, "N")
                            Sql4 = Sql4 & " order by 1 "
                            
                            Set Rs4 = New ADODB.Recordset
                            Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                            
                            Accion1 = 0
                            Accion2 = 0
                            Accion3 = 0
                            
                            While Not Rs4.EOF
                                Select Case DBLet(Rs4!numfases, "N")
                                    Case 1
                                        Accion1 = DBLet(Rs4!Acciones, "N")
                                    Case 2
                                        Accion2 = DBLet(Rs4!Acciones, "N")
                                    Case 3
                                        Accion3 = DBLet(Rs4!Acciones, "N")
                                End Select
                                Rs4.MoveNext
                            Wend
                            
                            Set Rs4 = Nothing
                            
                            Text51csb = "Acc.Fase 1 : " & Format(Accion1, "##0.00") & " Acc.Fase 2 : " & Format(Accion2, "##0.00")
                                        '123456789012345                                     67890123      4567                                     8901234567
                            Text52csb = "Acc.Fase 3 : " & Format(Accion3, "##0.00")
                            Text53csb = ""
                            Text61csb = "SOCIO : " & DBLet(Rs!nomsocio, "T")
                            Text62csb = ""
                            Text63csb = "N.I.F.: " & DBLet(Rs!nifSocio, "N")
                            Text71csb = ""
                            Text72csb = ""
                            Text73csb = ""
                            Text81csb = ""
                            Text82csb = ""
                            Text83csb = ""
                            
                            Referencia = ""
                        End If
                    End If
                End If
                
                Set Rs6 = Nothing
                
            Case "TAL" ' recibos de talla de escalona
                Text33csb = "** Recibo Talla **"
                
                Sql6 = "select * from rrecibpozos where codtipom = " & DBSet(RS1!CodTipom, "T")
                Sql6 = Sql6 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                Sql6 = Sql6 & " and fecfactu = " & DBSet(RS1!fecfactu, "F")
                
                Set Rs6 = New ADODB.Recordset
                Rs6.Open Sql6, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not Rs6.EOF Then
                    Text42csb = ""
                    Text51csb = ""
                    Text53csb = ""
                    Text62csb = ""
                    Text71csb = ""
                    Text73csb = ""
                    Text82csb = ""
                    Text41csb = ""
                    Text43csb = ""
                    Text52csb = ""
                    Text61csb = ""
                    Text63csb = ""
                    Text72csb = ""
                    Text81csb = ""
'[Monica]29/01/2014:
'                    Text33csb = RecuperaValor(ParteCadena(DBLet(Rs6!Concepto, "T"), 3, 40), 1)
                    '[Monica]29/01/2014: cambios text33csb(80) text41csb(60)
                     Text33csb = ""
                     Text41csb = ""
                     cad = Mid(Year(DBLet(RS1!fecfactu, "F")), 3, 2) & "TAL" & Format(DBLet(RS1!numfactu, "N"), "0000000")
                     cad = cad & " Soc:" & Format(DBLet(RS1!Codsocio, "N"), "000000") & " " & Mid(DBLet(Rs6!Concepto), 1, 15) & " Precios según información enviada"
                     
                     Text33csb = cad
                     
                     cad = "Tot:" & Format(DBLet(Rs6!TotalFact, "N") - DBLet(Rs6!imprecargo, "N"), "####0.00") & " "
                    '[Monica]15/01/2016: metemos el recargo
                    If DBLet(Rs6!imprecargo, "N") <> 0 Then
                        cad = cad & " Rec:" & Format(DBLet(Rs6!imprecargo, "N"), "###0.00")
                    End If
                     
                    
                    Sql4 = "select rpartida.nomparti, rrecibpozos_cam.poligono, rrecibpozos_cam.parcela, rrecibpozos_cam.hanegada, (if(rrecibpozos_cam.precio1 is null, 0, rrecibpozos_cam.precio1) + if(rrecibpozos_cam.precio2 is null, 0, rrecibpozos_cam.precio2)) precio, "
                    Sql4 = Sql4 & " rrecibpozos.porcdto, rrecibpozos.impdto, rrecibpozos.totalfact, rrecibpozos_cam.subparce "
                    Sql4 = Sql4 & " from rcampos, rpartida, rrecibpozos_cam, rrecibpozos "
                    Sql4 = Sql4 & " where rcampos.codparti = rpartida.codparti "
                    Sql4 = Sql4 & " and rcampos.codcampo = rrecibpozos_cam.codcampo "
                    Sql4 = Sql4 & " and rrecibpozos_cam.codtipom = " & DBSet(Rs6!CodTipom, "T")
                    Sql4 = Sql4 & " and rrecibpozos_cam.numfactu = " & DBSet(Rs6!numfactu, "N")
                    Sql4 = Sql4 & " and rrecibpozos_cam.fecfactu = " & DBSet(Rs6!fecfactu, "F")
                    Sql4 = Sql4 & " and rrecibpozos_cam.codtipom = rrecibpozos.codtipom"
                    Sql4 = Sql4 & " and rrecibpozos_cam.numfactu = rrecibpozos.numfactu"
                    Sql4 = Sql4 & " and rrecibpozos_cam.fecfactu = rrecibpozos.fecfactu"
                    
                    Set Rs4 = New ADODB.Recordset
                    Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    I = 0
                    While Not Rs4.EOF And I < 6 '15
                        I = I + 1

'                        hanegada = DBLet(DBLet(Rs4!hanegada, "N"))
'                        'Brazas = (Int(Hanegada) * 200) + (Hanegada - Int(Hanegada)) * 1000
'                        v_hanegada = Int(hanegada)
'                        v_brazas = (hanegada - Int(hanegada)) * 200
                    
                        'CadValues = Mid(Rs4!nomparti, 1, 15) & " " & Format(DBLet(Rs4!Poligono, "N"), "##0") & " " & Format(DBLet(Rs4!Parcela, "N"), "####0") & " " & DBLet(Rs4!SubParce, "T") & " " & Format(v_hanegada, "##0") & " " & Format(v_brazas, "###0") & " " & Format(DBLet(Rs4!Precio, "N"), "##0.0000")
                        cad = cad & Format(DBLet(Rs4!Poligono, "N"), "00") & "/" & Format(DBLet(Rs4!Parcela, "N"), "000")
                        If DBLet(Rs4!SubParce, "T") = "" Then
                            cad = cad & "  "
                        Else
                            cad = cad & Mid(DBLet(Rs4!SubParce, "T"), 1, 1) & " "
                        End If
                        
'                        Select Case i
'                            Case 1
'                                Text42csb = CadValues
'                            Case 2
'                                Text51csb = CadValues
'                            Case 3
'                                Text53csb = CadValues
'                            Case 4
'                                Text62csb = CadValues
'                            Case 5
'                                Text71csb = CadValues
'                            Case 6
'                                Text73csb = CadValues
'                            Case 7
'                                Text82csb = CadValues
'                            Case 8
'                                Text41csb = CadValues
'                            Case 9
'                                Text43csb = CadValues
'                            Case 10
'                                Text52csb = CadValues
'                            Case 11
'                                Text61csb = CadValues
'                            Case 12
'                                Text63csb = CadValues
'                            Case 13
'                                Text72csb = CadValues
'                            Case 14
'                                Text81csb = CadValues
'                        End Select
'
                        Rs4.MoveNext
                    Wend
                    Text41csb = cad
'                    If i > 14 Then Text81csb = "y otros"
'
'                    Text83csb = ""
'                    If DBLet(Rs6!PorcDto) <> 0 Then
'                        If DBLet(Rs6!PorcDto) < 0 Then
'                            Base = Rs6!TotalFact + Rs6!ImpDto
'                            CadValues = "Recargo " & Format(Base, "###,##0.00") & " " & Format(Abs(Rs6!PorcDto), "##0.00") & " " & Format(Rs6!TotalFact, "###,##0.00")
'                        Else
'                            Base = Rs6!TotalFact + Rs6!ImpDto
'                            CadValues = "Bonificacion " & Format(Base, "###,##0.00") & " " & Format(Rs6!PorcDto, "##0.00") & " " & Format(Rs6!TotalFact, "###,##0.00")
'                        End If
'                        Text83csb = CadValues
'                    End If
                 
'                    '[Monica]03/08/2012: la referencia en Escalona es el codigo de socio
'                    Referencia = DBLet(RS1!Codsocio, "T")

                    '[Monica]08/01/2014: para el caso de escalona lo cambiamos para que imprima en referencia
                    '                    el codigo de socio con formato, quito el if de arriba
                    Referencia = Format(DBLet(RS1!Codsocio, "T"), "000000")
                 
                End If
                
                Set Rs6 = Nothing
                
            '[Monica]27/06/2013 añadimos los recibos de contadores funcionan como los de mto
            Case "RVP" ' recibos de Contadores
                Text33csb = "** Recibo de Contadores **"
                
                Sql6 = "select * from rrecibpozos where codtipom = " & DBSet(RS1!CodTipom, "T")
                Sql6 = Sql6 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                Sql6 = Sql6 & " and fecfactu = " & DBSet(RS1!fecfactu, "F")
                
                Set Rs6 = New ADODB.Recordset
                Rs6.Open Sql6, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not Rs6.EOF Then
                
                    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
'[Monica]29/01/2014: sustituido por lo de abajo
'                        Text41csb = "FACTURA: " & Format(DBLet(RS1!numfactu, "N"), "#######") & " DE FECHA " & Format(DBLet(RS1!fecfactu, "N"), "dd/mm/yyyy")
'
'                        Text42csb = "CONCEPTO: "
'
'                        Text51csb = ""
'                        Text52csb = ""
'                        Text53csb = ""
'                        Text61csb = ""
'                        Text62csb = ""
'                        Text63csb = ""
'                        Text71csb = ""
'                        Text72csb = ""
'                        Text73csb = ""
'                        Text81csb = ""
'                        Text82csb = ""
'                        Text83csb = ""
'
'                        Text51csb = DBLet(Rs6!Conceptomo, "T")
'                        If Not IsNull(Rs6!importemo) Then
'                            Text52csb = Right("          " & Format(DBLet(Rs6!importemo, "N"), "###,##0.00"), 10)
'                        End If
'
'                        Text53csb = DBLet(Rs6!Conceptoar1, "T")
'                        If Not IsNull(Rs6!importear1) Then
'                            Text61csb = Right("          " & Format(DBLet(Rs6!importear1, "N"), "###,##0.00"), 10)
'                        End If
'                        Text62csb = DBLet(Rs6!Conceptoar2, "T")
'                        If Not IsNull(Rs6!importear2) Then
'                            Text63csb = Right("          " & Format(DBLet(Rs6!importear2, "N"), "###,##0.00"), 10)
'                        End If
'                        Text71csb = DBLet(Rs6!Conceptoar3, "T")
'                        If Not IsNull(Rs6!importear3) Then
'                            Text72csb = Right("          " & Format(DBLet(Rs6!importear3, "N"), "###,##0.00"), 10)
'                        End If
'                        Text73csb = DBLet(Rs6!Conceptoar4, "T")
'                        If Not IsNull(Rs6!importear4) Then
'                            Text81csb = Right("          " & Format(DBLet(Rs6!importear4, "N"), "###,##0.00"), 10)
'                        End If
'                        Text82csb = "TOTAL FACTURA"
'                        Text83csb = Right("          " & Format(DBLet(Rs6!TotalFact, "N"), "###,##0.00"), 10)
'
                        '[Monica]29/01/2014: cambios text33csb(80) text41csb(60)
                        Text33csb = ""
                        Text41csb = ""
                        cad = Mid(Year(DBLet(RS1!fecfactu, "F")), 3, 2) & "RVP" & Format(DBLet(RS1!numfactu, "N"), "0000000")
                        cad = cad & " Soc:" & Format(DBLet(RS1!Codsocio, "N"), "000000") & " " & Mid(DBLet(Rs6!importemo), 1, 30) & " Tot:" & Format(DBLet(Rs6!TotalFact, "N") - DBLet(Rs6!imprecargo, "N"), "####0.00")
                        
                        '[Monica]15/01/2016: metemos el recargo
                        If DBLet(Rs6!imprecargo, "N") <> 0 Then
                            cad = cad & " Rec:" & Format(DBLet(Rs6!imprecargo, "N"), "###0.00")
                        End If
                        
                         
                        Text33csb = cad
                         
                        cad = DBLet(Rs6!Conceptoar1, "T") & "/" & DBLet(Rs6!Conceptoar2, "T")
                        
                        Text41csb = cad
                        
'                        Referencia = DBLet(Rs6!Codsocio, "N")
                    '[Monica]08/01/2014: para el caso de escalona lo cambiamos para que imprima en referencia
                    '                    el codigo de socio con formato, quito el if de arriba
                        Referencia = Format(DBLet(Rs6!Codsocio, "N"), "000000")
                    End If
                    
                End If
                
                Set Rs6 = Nothing
            
            Case "RMT" ' recibos a manta (solo para Escalona)
                Text33csb = "** Recibo de Consumo a Manta **"
                
                Sql6 = "select * from rrecibpozos where codtipom = " & DBSet(RS1!CodTipom, "T")
                Sql6 = Sql6 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                Sql6 = Sql6 & " and fecfactu = " & DBSet(RS1!fecfactu, "F")
                
                Set Rs6 = New ADODB.Recordset
                Rs6.Open Sql6, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not Rs6.EOF Then
                
                        Text42csb = ""
                        Text51csb = ""
                        Text53csb = ""
                        Text62csb = ""
                        Text71csb = ""
                        Text73csb = ""
                        Text82csb = ""
                        Text41csb = ""
                        Text43csb = ""
                        Text52csb = ""
                        Text61csb = ""
                        Text63csb = ""
                        Text72csb = ""
                        Text81csb = ""
                        
                        '[Monica]29/01/2014: cambios text33csb(80) text41csb(60)
                        Text33csb = ""
                        Text41csb = ""
                        cad = Mid(Year(DBLet(RS1!fecfactu, "F")), 3, 2) & "RMT" & Format(DBLet(RS1!numfactu, "N"), "0000000")
                        cad = cad & " " & DBLet(Rs6!Concepto)
                         
                        Text33csb = cad
                         
                        cad = "Tot:" & Format(DBLet(Rs6!TotalFact, "N") - DBLet(Rs6!imprecargo, "N"), "####0.00") & " "
                        '[Monica]15/01/2016: metemos el recargo
                        If DBLet(Rs6!imprecargo, "N") <> 0 Then
                            cad = cad & " Rec:" & Format(DBLet(Rs6!imprecargo, "N"), "###0.00")
                        End If
                        
                        Sql4 = "select rpartida.nomparti, rrecibpozos_cam.poligono, rrecibpozos_cam.parcela, rrecibpozos_cam.hanegada, rrecibpozos_cam.precio1, "
                        Sql4 = Sql4 & " rrecibpozos.totalfact "
                        Sql4 = Sql4 & " from rpartida, rrecibpozos_cam, rrecibpozos, rcampos "
                        Sql4 = Sql4 & " where rcampos.codparti = rpartida.codparti "
                        Sql4 = Sql4 & " and rrecibpozos_cam.codtipom = " & DBSet(Rs6!CodTipom, "T")
                        Sql4 = Sql4 & " and rrecibpozos_cam.numfactu = " & DBSet(Rs6!numfactu, "N")
                        Sql4 = Sql4 & " and rrecibpozos_cam.fecfactu = " & DBSet(Rs6!fecfactu, "F")
                        Sql4 = Sql4 & " and rrecibpozos_cam.codcampo = rcampos.codcampo "
                        Sql4 = Sql4 & " and rrecibpozos_cam.codtipom = rrecibpozos.codtipom"
                        Sql4 = Sql4 & " and rrecibpozos_cam.numfactu = rrecibpozos.numfactu"
                        Sql4 = Sql4 & " and rrecibpozos_cam.fecfactu = rrecibpozos.fecfactu"
                        
                        Set Rs4 = New ADODB.Recordset
                        Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        
                        If Not Rs4.EOF Then
                            hanegada = DBLet(Rs4!hanegada, "N")
                            'Brazas = (Int(Hanegada) * 200) + (Hanegada - Int(Hanegada)) * 1000
                            v_hanegada = Int(hanegada)
                            v_brazas = (hanegada - Int(hanegada)) * 200
                        
                            cad = cad & " " & Mid(DBLet(Rs4!nomparti, "T"), 1, 15) & " " & DBLet(Rs4!Poligono, "N") & " " & DBLet(Rs4!Parcela, "N") & " " & Format(v_hanegada, "###0") & "H " & Format(v_brazas, "###0") & "B " & Format(DBLet(Rs4!Precio1, "N"), "#,##0.0000")
                        End If
                        Text41csb = cad
                        
                        Set Rs4 = Nothing
                        
                        Referencia = Format(DBLet(RS1!Codsocio, "N"), "000000")
                    
                
                End If
                
                Set Rs6 = Nothing
            
            '[Monica]15/01/2016: todas las facturas rectificativas de escalona
            Case "RRC", "RRM", "RRT", "RRV" ', "RTA"
                 Text33csb = ""
                 Text41csb = ""
                 
                 cad = Mid(Year(DBLet(RS1!fecfactu, "F")), 3, 2) & DBLet(RS1!CodTipom, "T") & Format(DBLet(RS1!numfactu, "N"), "0000000")
                 cad = cad & " Rectifica la factura: " & DBLet(RS1!CodTipomrec, "T") & "-" & Format(DBLet(RS1!numfacturec, "N"), "0000000") & " de fecha " & Format(DBLet(RS1!fecfacturec, "F"), "dd/mm/yyyy")
                 
                 Text33csb = cad
                 
                 cad = "Tot:" & Format(DBLet(RS1!TotalFact, "N") - DBLet(RS1!imprecargo, "N"), "####0.00") & " "
                '[Monica]15/01/2016: metemos el recargo
                If DBLet(RS1!imprecargo, "N") <> 0 Then
                    cad = cad & " Rec:" & Format(DBLet(RS1!imprecargo, "N"), "###0.00")
                End If
                 
                
                Text41csb = cad
                Referencia = Format(DBLet(RS1!Codsocio, "T"), "000000")
                 
                
                '[Monica]15/01/2016: para el caso de Escalona cuando la factura es rectificativa actualizamos su cobro y el de la factura que rectifica
                '                    con el importe de vencimiento + gastos
                If vParamAplic.Cooperativa = 10 Then
                    If DBLet(RS1!CodTipom, "T") = "RRC" Or DBLet(RS1!CodTipom, "T") = "RRM" Or DBLet(RS1!CodTipom, "T") = "RRT" Or _
                       DBLet(RS1!CodTipom, "T") = "RRV" Or DBLet(RS1!CodTipom, "T") = "RTA" Then
                                 
                         Dim LSer As String
'                         LSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(RS1!CodTipom, "T"))
'
'                         SQL = "update scobro set impcobro = coalesce(impvenci,0) + coalesce(gastos,0), fecultco = " & DBSet(FecVenci, "F")
'                         SQL = SQL & " where numserie = " & DBSet(LSer, "T") & " and codfaccl = " & DBSet(RS1!numfactu, "N")
'                         SQL = SQL & " and fecfaccl = " & DBSet(RS1!fecfactu, "F")
'
'                         ConnConta.Execute SQL
'
                         LSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(RS1!CodTipomrec, "T"))
                         
                         Sql = "update scobro set impcobro = coalesce(impvenci,0) + coalesce(gastos,0), fecultco = " & DBSet(FecVenci, "F")
                         Sql = Sql & " where numserie = " & DBSet(LSer, "T") & " and codfaccl = " & DBSet(RS1!numfacturec, "N")
                         Sql = Sql & " and fecfaccl = " & DBSet(RS1!fecfacturec, "F")
                         
                         ConnConta.Execute Sql
                    End If
                End If
                
                
                InsertarEnTesoreriaPOZOS = True
                Exit Function
                
        End Select
        
        If Referencia <> "" Then Referencia = RellenaABlancos(Referencia, True, 12)
        
        CC = DBLet(DigcoSoc, "T")
        If DBLet(DigcoSoc, "T") = "**" Then CC = "00"
    
        UltimaFactura = DBLet(RS1!numfactu, "N")
    
    
        '[Monica]28/07/2014: si es Escalona y es recibo a manta la forma de pago es la que me ponen en el frame (para el caso de los recibos que NO CONTADOS)
        '                    para el caso de los recibos a manta que sean contado se hace un asiento
        If vParamAplic.Cooperativa = 10 And DBLet(RS1!CodTipom, "T") = "RMT" Then
        
        Else
            '[Monica]15/06/2012: si es escalona y la cuenta de banco son 10 8's, la forma de pago será la de contado de parametros
            '                    en lugar de la pasada en el frame
                '[Monica]19/08/2013: añado la condicion de que utxera tambien tiene contados
            If ((vParamAplic.Cooperativa = 10 Or vParamAplic.Cooperativa = 8) And Trim(CtaBaSoc) = String(10, "8")) Then 'Or (vParamAplic.Cooperativa = 10 And DBLet(RS1!CodTipom, "T") = "RMT") Then
                Forpa = vParamAplic.ForpaConPOZ
            End If
    
        End If
    
        CadValuesAux2 = "(" & DBSet(LetraSerie, "T") & "," & DBSet(UltimaFactura, "N") & "," & DBSet(RS1!fecfactu, "F") & ", 1," & DBSet(CtaSocio, "T") & ","
        CadValues2 = CadValuesAux2 & DBSet(Forpa, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet((TotalFact), "N") & "," & DBSet(CtaBanco, "T") & ","
        If vParamAplic.ContabilidadNueva Then
            vvIban = MiFormat(IbanSoc, "") & MiFormat(CStr(BancoSoc), "0000") & MiFormat(CStr(SucurSoc), "0000") & MiFormat(CC, "00") & MiFormat(CtaBaSoc, "0000000000")
        
            CadValues2 = CadValues2 & DBSet(vvIban, "T") & "," & ValorNulo & "," & ValorNulo & ","
            CadValues2 = CadValues2 & DBSet(Text33csb, "T") & "," & DBSet(Text41csb, "T") & ",1,"
            CadValues2 = CadValues2 & DBSet(Referencia, "T", "S") & ","
            CadValues2 = CadValues2 & DBSet(Rs!nomsocio, "T") & "," & DBSet(Rs!dirsocio, "T") & "," & DBSet(Rs!pobsocio, "T")
            CadValues2 = CadValues2 & "," & DBSet(Rs!codpostal, "T") & "," & DBSet(Rs!prosocio, "T") & "," & DBSet(Rs!nifSocio, "T") & ",'ES')"
            
            Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
            Sql = Sql & "ctabanc1, iban, fecultco, impcobro, "
            Sql = Sql & " text33csb, text41csb,  agente, referencia, "
            Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
            Sql = Sql & ") "
        Else
            CadValues2 = CadValues2 & DBSet(BancoSoc, "N", "S") & "," & DBSet(SucurSoc, "N", "S") & ","
            CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            CadValues2 = CadValues2 & DBSet(Text33csb, "T") & "," & DBSet(Text41csb, "T") & "," & DBSet(Text42csb, "T") & ",1,"
            CadValues2 = CadValues2 & DBSet(Text43csb, "T") & "," & DBSet(Text51csb, "T") & "," & DBSet(Text52csb, "T") & ","
            CadValues2 = CadValues2 & DBSet(Text53csb, "T") & "," & DBSet(Text61csb, "T") & "," & DBSet(Text62csb, "T") & ","
            CadValues2 = CadValues2 & DBSet(Text63csb, "T") & "," & DBSet(Text71csb, "T") & "," & DBSet(Text72csb, "T") & "," & DBSet(Text73csb, "T") & "," & DBSet(Text81csb, "T") & "," & DBSet(Text82csb, "T") & ","
            CadValues2 = CadValues2 & DBSet(Text83csb, "T") & ","
            CadValues2 = CadValues2 & DBSet(Referencia, "T", "S") & "," '& ")"
            
            '[Monica]03/08/2012: Metemos en todas las cooperativas los datos fiscales del socio
            CadValues2 = CadValues2 & DBSet(Rs!nomsocio, "T") & "," & DBSet(Rs!dirsocio, "T") & "," & DBSet(Rs!pobsocio, "T")
            CadValues2 = CadValues2 & "," & DBSet(Rs!codpostal, "T") & "," & DBSet(Rs!prosocio, "T") ' & ")"
            
            '[Monica]22/11/2013: Tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & ")"
            Else
                CadValues2 = CadValues2 & ")"
            End If
            
        
            'Insertamos en la tabla scobro de la CONTA
            Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
            Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
            '[Monica] 16/07/2010: hemos añadido todo lo que debe llevar impreso el recibo de banco ( desde agente )
            Sql = Sql & " text33csb, text41csb, text42csb, agente, text43csb, text51csb, text52csb, text53csb,"
            Sql = Sql & " text61csb, text62csb, text63csb, text71csb,text72csb,text73csb, text81csb, text82csb, text83csb, referencia, "
            Sql = Sql & " nomclien, domclien, pobclien, cpclien, proclien" ') "
            '[Monica]22/11/2013: Tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                Sql = Sql & ", iban) "
            Else
                Sql = Sql & ") "
            End If
        End If
        Sql = Sql & " VALUES " & CadValues2
        ConnConta.Execute Sql



        B = True
        
    Else
        MenError = "No se ha encontrado socio " & DBLet(RS1!Codsocio, "N") & " o no tiene seccion asignada. Revise. "
    End If
    
    
EInsertarTesoreriaPOZ:
    
    If Err.Number <> 0 Then
        B = False
        MenError = "Error al insertar en Tesoreria de POZOS: " & Err.Description
    End If
    InsertarEnTesoreriaPOZOS = B
End Function


Private Function ParteCadena(Origen As String, NroSubcadenas As Integer, longitud) As String
Dim J As Integer
Dim I As Integer
Dim k As Integer
Dim cad As String

    On Error Resume Next

    ParteCadena = ""

    J = 1
    cad = ""
    For k = 1 To NroSubcadenas
        I = J + longitud - 1
        If Len(Origen) - J > 0 Then
            If Len(Mid(Origen, J + 1, Len(Origen) - J)) > longitud Then
                While Mid(Origen, I + 1, 1) <> " "
                    I = I - 1
                Wend
            End If
            If J > 0 And I - J + 1 > 0 Then
                cad = cad & Mid(Origen, J, I - J + 1) & "|"
            End If
            J = I + 1
        End If
    Next k
    
    ParteCadena = cad
    

End Function


'----------------------------------------------------------------------
' FACTURAS TRANSPORTISTAS
'----------------------------------------------------------------------

Public Function PasarFacturaTra(cadWHERE As String, CodCCost As String, FechaFin As String, Seccion As String, TipoFact As Byte, FecRecep As Date, FecVto As Date, ForpaPos As String, ForpaNeg As String, CtaBanc As String, CtaRete As String, CtaApor As String, TipoM As String, ByRef vContaFra As cContabilizarFacturas, IvaRea As Integer) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariagro.rfactsoc --> conta.cabfactprov
' ariagro.rfactsoc_variedad --> conta.linfactprov
'Actualizar la tabla ariagro.rfactsoc.contabilizada=1 para indicar que ya esta contabilizada
Dim B As Boolean
Dim cadMen As String
Dim Sql As String
Dim Mc As Contadores


    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
        
        
    '[Monica]09/11/2016: nueva clase de socio
    Set vTra = New CTransportista
    
    Sql = "select codtrans from rfacttra where " & cadWHERE
    vTra.LeerDatos DevuelveValor(Sql)
        
    
    Set Mc = New Contadores
    
    '[Monica]17/10/2011: FACTURAS INTERNAS
    If EsFacturaInternaTrans(cadWHERE) Then
        CtaReten = CtaRete
        ' Insertamos en el diario
        B = InsertarAsientoDiarioTRANS(cadWHERE, cadMen, Mc, CDate(FechaFin), Seccion, TipoFact, FecRecep, TipoM)
        cadMen = "Insertando Factura en Diario: " & cadMen
    Else
        CtaReten = CtaRete
        '---- Insertar en la conta Cabecera Factura
        B = InsertarCabFactTra(cadWHERE, cadMen, Mc, CDate(FechaFin), Seccion, TipoFact, FecRecep, TipoM, vContaFra, IvaRea)
        cadMen = "Insertando Cab. Factura: " & cadMen
        
    End If
    
    If B Then
        FecVenci = FecVto
        ForpaPosi = ForpaPos
        ForpaNega = ForpaNeg
        CtaBanco = CtaBanc
        CtaReten = CtaRete
        CtaAport = CtaApor
        tipoMov = TipoM    ' codtipom de la factura de socio
        
'01-06-2009
        B = InsertarEnTesoreriaTra(cadWHERE, cadMen, FacturaTRA, FecFactuTRA)
        cadMen = "Insertando en Tesoreria: " & cadMen

        If B Then
            CCoste = CodCCost
            '[Monica]17/10/2011: INTERNAS
            If Not EsFacturaInternaTrans(cadWHERE) Then
                '---- Insertar lineas de Factura en la Conta
                If vParamAplic.ContabilidadNueva Then
                    B = InsertarLinFactTraContaNueva("rfacttra", cadWHERE, cadMen, TipoFact, FecRecep, Mc.Contador)
                Else
                    B = InsertarLinFactTra("rfacttra", cadWHERE, cadMen, TipoFact, Mc.Contador)
                End If
                cadMen = "Insertando Lin. Factura: " & cadMen
            End If
    
            If B Then
                '---- Poner intconta=1 en ariagro.rfacttra
                If Not EsFacturaInternaTrans(cadWHERE) Then
                    If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaProv(vContaFra.NumeroFactura, vContaFra.Anofac)
                End If
                B = ActualizarCabFactSoc("rfacttra", cadWHERE, cadMen)
                cadMen = "Actualizando Factura Transporte: " & cadMen
            End If
        End If
    End If
    
    
EContab:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Contabilizando Factura Transporte", Err.Description
    End If
    If B Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaTra = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaTra = False
        If Not B Then
            InsertarTMPErrFacSoc cadMen, cadWHERE
        End If
    End If
End Function


Private Function InsertarCabFactTra(cadWHERE As String, cadErr As String, ByRef Mc As Contadores, FechaFin As Date, Secci As String, Tipo As Byte, FecRec As Date, TipoM As String, ByRef vContaFra As cContabilizarFacturas, IvaRea As Integer) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Concepto As String
Dim letraser As String
Dim TipoOpera As Integer
Dim Aux As String
Dim Sql2 As String
Dim CadenaInsertFaclin2 As String


    On Error GoTo EInsertar
       
    
    Sql = " SELECT fecfactu,year(" & DBSet(FecRec, "F") & ") as anofacpr," & DBSet(FecRec, "F") & ",numfactu,rtransporte.codmacpro,"
    Sql = Sql & "baseimpo, tipoiva, porc_iva,imporiva,basereten,porc_ret,impreten,baseaport,porc_apo,impapor,totalfac,"
    Sql = Sql & "rectif_codtipom, rectif_numfactu, rectif_fecfactu,"
    Sql = Sql & "rtransporte.codtrans, rtransporte.nomtrans, rtransporte.codbanco, rtransporte.codsucur, rtransporte.digcontr, rtransporte.cuentaba "
    Sql = Sql & ",rtransporte.iban "
    Sql = Sql & ",rtransporte.dirtrans,rtransporte.pobtrans,rtransporte.codpostal,rtransporte.protrans,rtransporte.niftrans,rtransporte.codforpa  "
    '[Monica]03/05/2017: como en socios
    Sql = Sql & ",rfacttra.tipoirpf "
    Sql = Sql & " FROM (" & "rfacttra "
    Sql = Sql & "INNER JOIN rtransporte ON rfacttra.codtrans=rtransporte.codtrans) "
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
    
        If Mc.ConseguirContador("1", (FecRec <= CDate(FechaFin) - 365), True) = 0 Then
            vContaFra.NumeroFactura = Mc.Contador
            vContaFra.Anofac = DBLet(Rs!anofacpr)
            
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            BaseImp = DBLet(Rs!baseimpo, "N")
            TotalFac = BaseImp + DBLet(Rs!ImporIva, "N")
            AnyoFacPr = Rs!anofacpr
            
            ImpReten = DBLet(Rs!ImpReten, "N")
            ImpAport = DBLet(Rs!impapor, "N")
            
            letraser = ""
            letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(TipoM, "T"))
            
            FacturaTRA = letraser & "-" & DBLet(Rs!numfactu, "N")
            FecFactuTRA = DBLet(Rs!fecfactu, "F")
            
            CodTipomRECT = DBLet(Rs!rectif_codtipom, "T")
            NumfactuRECT = DBLet(Rs!rectif_numfactu, "T")
            FecfactuRECT = DBLet(Rs!rectif_fecfactu, "T")
            
            CtaTransporte = Rs!codmacpro
            Seccion = Secci
            TipoFact = 0 'tipo
            FecRecep = FecRec
            BancoTRA = DBLet(Rs!CodBanco, "N")
            SucurTRA = DBLet(Rs!CodSucur, "N")
            DigcoTRA = DBLet(Rs!digcontr, "T")
            CtaBaTRA = DBLet(Rs!CuentaBa, "T")
            IbanTRA = DBLet(Rs!Iban, "T")
            TotalTesor = DBLet(Rs!TotalFac, "N")
            
'            Variedades = VariedadesFactura(cadwhere)
            Variedades = ""
            
            Select Case TipoFact
                Case 0 ' anticipo
                    Concepto = "FACTURA TRANSPORTE"
                Case 11
                    Concepto = "Rectificativa"
                Case Else
                    Concepto = ""
            End Select
            
            Sql = ""
            If vParamAplic.ContabilidadNueva Then Sql = "'" & SerieFraPro & "',"
            Sql = Sql & Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & AnyoFacPr & "," & DBSet(FecRecep, "F") & "," & DBSet(FecRecep, "F") & "," & DBSet(FacturaTRA, "T") & "," & DBSet(CtaTransporte, "T") & "," & DBSet(Concepto, "T") & ","
            
            If Not vParamAplic.ContabilidadNueva Then
                Sql = Sql & DBSet(BaseImp, "N") & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & DBSet(Rs!porc_iva, "N") & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!ImporIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & DBSet(TotalFac, "N") & "," & DBSet(Rs!TipoIVA, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                cad = cad & "(" & Sql & ")"
                
                'Insertar en la contabilidad
                Sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,fecliqpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                Sql = Sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                Sql = Sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,nodeducible) "
                Sql = Sql & " VALUES " & cad
                ConnConta.Execute Sql
            Else
                
                Sql = Sql & DBSet(Rs!nomtrans, "T") & "," & DBSet(Rs!dirtrans, "T", "S") & ","
                Sql = Sql & DBSet(Rs!codpostal, "T", "S") & "," & DBSet(Rs!pobtrans, "T", "S") & "," & DBSet(Rs!protrans, "T", "S") & ","
                Sql = Sql & DBSet(Rs!NIFTrans, "F", "S") & ",'ES',"
                Sql = Sql & DBSet(Rs!Codforpa, "N") & ","
            
                '$$$
                '[Monica]02/05/2017: Solo en el caso de iva rea
                If DBLet(Rs!TipoIVA, "N") = IvaRea Then
               
                    TipoOpera = 5 ' REA
                    
                    '[Monica]21/04/2017: antes tenia un 0 en Aux
                    Aux = "X"
'                    If Rs!TotalFac < 0 Then Aux = "D"
                    'codopera,codconce340,codintra
                    Sql = Sql & TipoOpera & "," & DBSet(Aux, "T") & "," & ValorNulo & ","
                    
                Else
                
                    TipoOpera = 0 ' general
                    
                    '[Monica]21/04/2017: antes tenia un 0 en Aux
                    Aux = "0" ' estaba X
'                    If Rs!TotalFac < 0 Then Aux = "D"
                    'codopera,codconce340,codintra
                    Sql = Sql & TipoOpera & "," & DBSet(Aux, "T") & "," & ValorNulo & ","
                    
                End If
                
                '[Monica]10/11/2016: en totalfac llevabamos base + impiva pq antes retencion estaba en lineas
                '                    en la nueva conta está en la cabecera
                TotalFac = TotalFac - ImpReten
                
                'para las lineas
                'factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
                'IVA 1, siempre existe
                Aux = "'" & SerieFraPro & "'," & Mc.Contador & "," & DBSet(FecRecep, "F") & "," & Rs!anofacpr & ","
                
                Sql2 = Aux & "1," & DBSet(BaseImp, "N") & "," & DBSet(Rs!TipoIVA, "N") & "," & DBSet(Rs!porc_iva, "N") & ","
                Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!ImporIva, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
                    
                'Los totales
                'totbases,totbasesret,totivas,totrecargo,totfacpr,
                Sql = Sql & DBSet(BaseImp, "N") & "," & DBSet(Rs!BaseReten, "N", "S") & ","
                'totivas
                Sql = Sql & DBSet(Rs!ImporIva, "N") & "," & DBSet(TotalFac, "N") & ","
                If DBLet(Rs!porc_ret, "N") <> 0 Then
                    Sql = Sql & DBSet(Rs!porc_ret, "N") & "," & DBSet(Rs!ImpReten, "N") & "," & DBSet(CtaReten, "T") & ","
                                        
                    '[Monica]03/05/2017: tipo de retencion
'               si retencion : Si REA + modulos ----> tipo retencion = 2 (act.agricola)
'                              Si no REA + modulos--> tipo retencion = 1 (act.profesional)
'                              si E.D.  ------------> tipo retencion = 4 (act.empresarial)
                    If Rs!TipoIVA = IvaRea And Rs!TipoIRPF = 0 Then Sql = Sql & "2"
                    If Rs!TipoIVA <> IvaRea And Rs!TipoIRPF = 0 Then Sql = Sql & "1"
                    If Rs!TipoIRPF = 1 Or Rs!TipoIRPF = 3 Then Sql = Sql & "4"
                    
                Else
                    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                End If
                cad = cad & "(" & Sql & ")"
            
            
                'Insertar en la contabilidad
                Sql = "INSERT INTO factpro(numserie,numregis,fecfactu,anofactu,fecharec,fecliqpr,numfactu,codmacta,observa,nommacta,"
                Sql = Sql & "dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,codforpa,codopera,codconce340,codintra,"
                Sql = Sql & "totbases,totbasesret,totivas,totfacpr,retfacpr,trefacpr,cuereten,tiporeten)"
                Sql = Sql & " VALUES " & cad
                ConnConta.Execute Sql
            
                'Las  lineas de IVA
                Sql = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)"
                Sql = Sql & " VALUES " & CadenaInsertFaclin2
                ConnConta.Execute Sql
        
            End If
            'añadido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            Sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            Sql = Sql & ",'" & DevNombreSQL(FacturaTRA) & " @ " & Format(Rs!fecfactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!codTrans) & "')"
            conn.Execute Sql
            
            FacturaTRA = DBLet(Rs!numfactu, "N")
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactTra = False
        cadErr = Err.Description
    Else
        InsertarCabFactTra = True
    End If
End Function


Public Function InsertarEnTesoreriaTra(cadWHERE As String, MenError As String, numfactu As String, fecfactu As Date) As Boolean
'Guarda datos de Tesoreria en tablas: spagop o scobro dependiendo del signo de la factura
Dim B As Boolean
Dim Sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim CC As String
Dim letraser As String
Dim Text33csb As String
Dim Text41csb As String
Dim Text42csb As String
Dim GastosPie As Currency
Dim vvIban As String

    On Error GoTo EInsertarTesoreriaTra

    InsertarEnTesoreriaTra = False
    
    
    If TotalTesor > 0 Then ' se insertara en la cartera de pagos (spagop)
        CadValues2 = ""
    
        'vamos creando la cadena para insertar en spagosp de la CONTA
        letraser = ""
        letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(tipoMov, "T"))
        
        '[Monica]03/07/2013: añado trim(codmacta)
        CadValuesAux2 = "("
        If vParamAplic.ContabilidadNueva Then CadValuesAux2 = CadValuesAux2 & DBSet(SerieFraPro, "T") & ","
        CadValuesAux2 = CadValuesAux2 & "'" & Trim(CtaTransporte) & "', " & DBSet(letraser & "-" & numfactu, "T") & ", '" & Format(fecfactu, FormatoFecha) & "', "
    
        '------------------------------------------------------------
        I = 1
        CadValues2 = CadValuesAux2 & I
        CadValues2 = CadValues2 & ", " & ForpaPosi & ", '" & Format(FecVenci, FormatoFecha) & "', "
        CadValues2 = CadValues2 & DBSet(TotalTesor, "N") & ", " & DBSet(CtaBanco, "T") & ","
    
    
        If Not vParamAplic.ContabilidadNueva Then
            'David. Para que ponga la cuenta bancaria (SI LA tiene)
            CadValues2 = CadValues2 & DBSet(BancoTRA, "T", "S") & "," & DBSet(SucurTRA, "T", "S") & ","
            CadValues2 = CadValues2 & DBSet(DigcoTRA, "T", "S") & "," & DBSet(CtaBaTRA, "T", "S") & ","
        End If
        'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
        Sql = "Factura num.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
            
        CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
        
        'SQL = "Variedades: " & Variedades
        Sql = ""
        CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" ')"
        
        If vParamAplic.ContabilidadNueva Then
            vvIban = MiFormat(IbanTRA, "") & MiFormat(CStr(BancoTRA), "0000") & MiFormat(CStr(SucurTRA), "0000") & MiFormat(DigcoTRA, "00") & MiFormat(CtaBaTRA, "0000000000")
            
            CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
            'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
            CadValues2 = CadValues2 & DBSet(vTra.Nombre, "T") & "," & DBSet(vTra.Direccion, "T") & "," & DBSet(vTra.Poblacion, "T") & "," & DBSet(vTra.CPostal, "T") & ","
            CadValues2 = CadValues2 & DBSet(vTra.Provincia, "T") & "," & DBSet(vTra.nif, "T") & ",'ES') "
        
        Else
            '[Monica]22/11/2013: Tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                CadValues2 = CadValues2 & ", " & DBSet(IbanTRA, "T", "S") & ") "
            Else
                CadValues2 = CadValues2 & ") "
            End If
        End If
    
        'Grabar tabla spagop de la CONTABILIDAD
        '-------------------------------------------------
        If CadValues2 <> "" Then
            'Insertamos en la tabla spagop de la CONTA
            'David. Cuenta bancaria y descripcion textos
            If vParamAplic.ContabilidadNueva Then
                Sql = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, iban,"
                Sql = Sql & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais)"
            Else
                Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    Sql = Sql & ", iban) "
                Else
                    Sql = Sql & ") "
                End If
            End If
            Sql = Sql & " VALUES " & CadValues2
            ConnConta.Execute Sql
        End If
    Else
        ' si es negativo se inserta en positivo en la cartera de cobros (scobro)
        
        letraser = ""
        letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(tipoMov, "T"))
        
        Text33csb = "'Factura:" & DBLet(letraser, "T") & "-" & DBLet(numfactu, "T") & " " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
        Text41csb = "de " & DBSet(TotalTesor, "N")
'        text42csb = "Variedades: " & Variedades
        
        CC = DBLet(DigcoTRA, "T")
        If DBLet(DigcoTRA, "T") = "**" Then CC = "00"
    
        '[Monica]03/07/2013: añado trim(codmacta)
        CadValuesAux2 = "(" & DBSet(letraser, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & ", 1," & DBSet(Trim(CtaTransporte), "T") & ","
        CadValues2 = CadValuesAux2 & DBSet(ForpaNega, "N") & "," & DBSet(fecfactu, "F") & "," & DBSet(TotalTesor * (-1), "N") & ","
        CadValues2 = CadValues2 & DBSet(CtaBanco, "T") & ","
        
        If Not vParamAplic.ContabilidadNueva Then
        
                CadValues2 = CadValues2 & DBSet(BancoTRA, "N", "S") & "," & DBSet(SucurTRA, "N", "S") & ","
                CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(CtaBaTRA, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & "," & DBSet(Text42csb, "T") & ",1" ')"
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    CadValues2 = CadValues2 & ", " & DBSet(IbanTRA, "T", "S") & ") "
                Else
                    CadValues2 = CadValues2 & ") "
                End If
                
                
                'Insertamos en la tabla scobro de la CONTA
                Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                Sql = Sql & " text33csb, text41csb, text42csb, agente" ') "
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    Sql = Sql & ", iban) "
                Else
                    Sql = Sql & ") "
                End If
                
        Else
                CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & ","
                CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & "," & DBSet(Text42csb, "T") & ",1,"
                
                vvIban = MiFormat(IbanTRA, "") & MiFormat(CStr(BancoTRA), "0000") & MiFormat(CStr(SucurTRA), "0000") & MiFormat(CC, "00") & MiFormat(CtaBaTRA, "0000000000")
                
                CadValues2 = CadValues2 & DBSet(vvIban, "T") & ","
                'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                CadValues2 = CadValues2 & DBSet(vTra.Nombre, "T") & "," & DBSet(vTra.Direccion, "T") & "," & DBSet(vTra.Poblacion, "T") & "," & DBSet(vTra.CPostal, "T") & ","
                CadValues2 = CadValues2 & DBSet(vTra.Provincia, "T") & "," & DBSet(vTra.nif, "T") & ",'ES'),"
        
                Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                Sql = Sql & "ctabanc1, fecultco, impcobro, "
                Sql = Sql & " text33csb, text41csb, text42csb, agente, iban, "
                Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
                Sql = Sql & ") "
        End If
        
        Sql = Sql & " VALUES " & CadValues2
        ConnConta.Execute Sql
    
    End If

    B = True

EInsertarTesoreriaTra:
    If Err.Number <> 0 Then
        B = False
        MenError = "Error al insertar en Tesoreria: " & Err.Description
    End If
    InsertarEnTesoreriaTra = B
End Function


Private Function InsertarLinFactTra(cadTabla As String, cadWHERE As String, cadErr As String, Tipo As Byte, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String
Dim LineaVariedad As Integer

Dim vSocio As cSocio
Dim Socio As String
Dim TipoAnt As Byte
Dim TipoFact As String



Dim ImpAnticipo As Currency
    On Error GoTo EInLinea
    
    TipoAnt = Tipo
'    TipoFactAnt = TipoFact
    
    If Tipo = 11 Then ' si es una factura rectificativa cojo el tipo de movimiento de la factura que rectifico
        Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(CodTipomRECT, "T"))
        
        TipoFact = CodTipomRECT

    Else
' Estoy aqui: en liquidacion de industria

'select if(rsocios.tipoprod = 1, variedades.ctacomtercero, variedades.ctaliquidacion) as cuenta
'From rsocios, Variedades, rfactsoc, rfactsoc_variedad
'where rsocios.codsocio= rfactsoc.codsocio and mid(rfactsoc.codtipom,1,3) = "FLI" and
'rfactsoc.codtipom= rfactsoc_variedad.codtipom and
'rfactsoc.numfactu = rfactsoc_variedad.codtipom and
'rfactsoc.fecfactu = rfactsoc_variedad.fecfactu and
'rfactsoc_variedad.codvarie = Variedades.codvarie

        TipoFact = "FTR"
    
    End If
    
    
    If vEmpresa.TieneAnalitica Then
        Sql = " SELECT 1, variedades.ctatransporte as cuenta, sum(rfacttra_albaran.importe) as importe, variedades.codccost "
    Else
        Sql = " SELECT 1, variedades.ctatransporte as cuenta, sum(rfacttra_albaran.importe) as importe "
    End If
    Sql = Sql & " FROM rfacttra_albaran, variedades "
    Sql = Sql & " WHERE " & Replace(cadWHERE, "rfacttra", "rfacttra_albaran") & " and"
    Sql = Sql & " rfacttra_albaran.codvarie = variedades.codvarie "
    Sql = Sql & " group by 1,2 "
    Sql = Sql & " order by 1,2 "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        ImpLinea = Rs!Importe
        
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(Rs!cuenta, "T")
        
        Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
                If DBLet(Rs!CodCCost, "T") = "----" Then
                    Sql = Sql & DBSet(CCoste, "T")
                Else
                    Sql = Sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBLet(Rs!CodCCost, "T")
                End If
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)
        Sql2 = Sql2 & DBSet(totimp, "N") & ","
        If CCoste = "" Or CCoste = ValorNulo Then
            Sql2 = Sql2 & ValorNulo
        Else
            Sql2 = Sql2 & DBSet(CCoste, "T")
        End If
        If SQLaux <> "" Then 'hay mas de una linea
            cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If

    ' las retenciones si las hay
    If ImpReten <> 0 Then
        Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(CtaTransporte, "T")
        Sql = Sql & "," & DBSet(ImpReten, "N") & ","
        Sql = Sql & ValorNulo ' no llevan centro de coste
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
        
        Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(CtaReten, "T")
        Sql = Sql & "," & DBSet(ImpReten * (-1), "N") & ","
        Sql = Sql & ValorNulo ' no llevan centro de coste
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
    End If
    
    ' las aportaciones de fondo operativo si las hay
    If ImpAport <> 0 Then
        Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(CtaTransporte, "T")
        Sql = Sql & "," & DBSet(ImpAport, "N") & ","
        Sql = Sql & ValorNulo ' no llevan centro de coste
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
    
        Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(CtaAport, "T")
        Sql = Sql & "," & DBSet(ImpAport * (-1), "N") & ","
        Sql = Sql & ValorNulo ' no llevan centro de coste
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
    End If
    
    
    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If
    
    Tipo = TipoAnt

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactTra = False
        cadErr = Err.Description
    Else
        InsertarLinFactTra = True
    End If
End Function


Private Function InsertarLinFactTraContaNueva(cadTabla As String, cadWHERE As String, cadErr As String, Tipo As Byte, FecRecep As Date, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String
Dim LineaVariedad As Integer

Dim vSocio As cSocio
Dim Socio As String
Dim TipoAnt As Byte
Dim TipoFact As String

Dim ImpAnticipo As Currency

Dim vTipoIvaAux As Currency
Dim vImpIvaAux As Currency
Dim vPorIvaAux As Currency
Dim impiva As Currency
Dim TotImpIVA As Currency
Dim SqlAux2 As String




    On Error GoTo EInLinea
    
    TipoAnt = Tipo
'    TipoFactAnt = TipoFact
    
    If Tipo = 11 Then ' si es una factura rectificativa cojo el tipo de movimiento de la factura que rectifico
        Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(CodTipomRECT, "T"))
        
        TipoFact = CodTipomRECT

    Else
' Estoy aqui: en liquidacion de industria

'select if(rsocios.tipoprod = 1, variedades.ctacomtercero, variedades.ctaliquidacion) as cuenta
'From rsocios, Variedades, rfactsoc, rfactsoc_variedad
'where rsocios.codsocio= rfactsoc.codsocio and mid(rfactsoc.codtipom,1,3) = "FLI" and
'rfactsoc.codtipom= rfactsoc_variedad.codtipom and
'rfactsoc.numfactu = rfactsoc_variedad.codtipom and
'rfactsoc.fecfactu = rfactsoc_variedad.fecfactu and
'rfactsoc_variedad.codvarie = Variedades.codvarie

        TipoFact = "FTR"
    
    End If
    
    
    If vEmpresa.TieneAnalitica Then
        Sql = " SELECT 1, variedades.ctatransporte as cuenta, sum(rfacttra_albaran.importe) as importe, variedades.codccost "
    Else
        Sql = " SELECT 1, variedades.ctatransporte as cuenta, sum(rfacttra_albaran.importe) as importe "
    End If
    Sql = Sql & " FROM rfacttra_albaran, variedades "
    Sql = Sql & " WHERE " & Replace(cadWHERE, "rfacttra", "rfacttra_albaran") & " and"
    Sql = Sql & " rfacttra_albaran.codvarie = variedades.codvarie "
    Sql = Sql & " group by 1,2 "
    Sql = Sql & " order by 1,2 "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText




    SqlAux2 = "select rfacttra.tipoiva from " & cadTabla & " where " & cadWHERE
    vTipoIvaAux = DevuelveValor(SqlAux2)
    
    SqlAux2 = "select rfacttra.porc_iva from " & cadTabla & " where " & cadWHERE
    vPorIvaAux = DevuelveValor(SqlAux2)
    
    SqlAux2 = "select rfacttra.imporiva from " & cadTabla & " where " & cadWHERE
    vImpIvaAux = DevuelveValor(SqlAux2)



    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        ImpLinea = Rs!Importe
        
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        Sql = DBSet(SerieFraPro, "T") & "," & NumRegis & "," & DBSet(FecRecep, "F") & "," & AnyoFacPr & "," & I & ","
        Sql = Sql & DBSet(Rs!cuenta, "T") & ","
        
        If vEmpresa.TieneAnalitica Then
            If DBLet(Rs!CodCCost, "T") = "----" Then
                Sql = Sql & DBSet(CCoste, "T")
            Else
                Sql = Sql & DBSet(Rs!CodCCost, "T")
                CCoste = DBLet(Rs!CodCCost, "T")
            End If
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        'tipo de iva, porcentaje iva y porcentaje recargo
        Sql = Sql & "," & vTipoIvaAux
        Sql = Sql & "," & vPorIvaAux
        Sql = Sql & "," & ValorNulo
        Sql = Sql & "," & DBSet(ImpLinea, "N")
        
        Sql2 = Sql & "," 'nos guardamos la linea sin el importe iva por si a la última hay q descontarle para q coincida con total factura
        
        impiva = Round(ImpLinea * vPorIvaAux / 100, 2)
        
        TotImpIVA = TotImpIVA + impiva
        
        Sql = Sql & "," & DBSet(impiva, "N") & ","
        
        ' llevan retencion
        Sql = Sql & ValorNulo & ",1"
        
        
        cad = cad & "(" & Sql & ")" & ","
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    If TotImpIVA <> vImpIvaAux Then
'        MsgBox "FALTA cuadrar importes de iva!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = vImpIvaAux - TotImpIVA
        totimp = impiva + totimp '(+- diferencia)
        Sql2 = Sql2 & DBSet(totimp, "N") & ","
        Sql2 = Sql2 & ValorNulo & ",1"
        If SQLaux <> "" Then 'hay mas de una linea
            cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If

    
    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO factpro_lineas (numserie,numregis,fecharec,anofactu,numlinea,codmacta,codccost,codigiva,porciva,porcrec,baseimpo,impoiva,imporec,aplicret) "
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If
    
    Tipo = TipoAnt

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactTraContaNueva = False
        cadErr = Err.Description
    Else
        InsertarLinFactTraContaNueva = True
    End If
End Function










Public Function EsFacturaInterna(cWhere As String) As Boolean
Dim Sql As String


    On Error GoTo eEsFacturaInterna
    
    Sql = "select rsocios.esfactadvinterna from rfactsoc inner join rsocios on rfactsoc.codsocio = rsocios.codsocio "
    Sql = Sql & " where " & cWhere
    
    EsFacturaInterna = (DevuelveValor(Sql) = 1)
    Exit Function
    
eEsFacturaInterna:
    MuestraError Err.Number, "Es Factura Interna", Err.Description
End Function

Public Function EsFacturaInternaTrans(cWhere As String) As Boolean
Dim Sql As String


    On Error GoTo eEsFacturaInternaTrans
    
    Sql = "select rtransporte.esfacttrainterna from rfacttra inner join rtransporte on rfacttra.codtrans = rtransporte.codtrans "
    Sql = Sql & " where " & cWhere
    
    EsFacturaInternaTrans = (DevuelveValor(Sql) = 1)
    Exit Function
    
eEsFacturaInternaTrans:
    MuestraError Err.Number, "Es Factura Interna de Transporte", Err.Description
End Function


Private Function EsContado(vCadena As String) As Boolean
Dim Sql As String

    Sql = "select escontado from rrecibpozos where " & vCadena
    EsContado = (DevuelveValor(Sql) = "1")

End Function

Public Function PasarFacturaPOZOS(cadWHERE As String, CodCCost As String, CtaBan As String, FecVen As String, TipoM As String, FecFac As Date, Observac As String, Forpa As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
Dim B As Boolean
Dim cadMen As String
Dim Sql As String
Dim Mc As Contadores
Dim Obs As String
Dim RS1 As ADODB.Recordset


    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    
    If TipoM <> "FIN" Then
    
        
        'Insertar en la conta Cabecera Factura
        B = InsertarCabFactPOZ(cadWHERE, Observac, cadMen, Forpa, vContaFra)
        cadMen = "Insertando Cab. Factura: " & cadMen
        
        If B Then
            CCoste = CodCCost
            'Insertar lineas de Factura en la Conta
            If vParamAplic.ContabilidadNueva Then
                B = InsertarLinFactPOZContaNueva("rrecibpozos", cadWHERE, cadMen, TipoM)
            Else
                B = InsertarLinFactPOZ("rrecibpozos", cadWHERE, cadMen, TipoM)
            End If
            cadMen = "Insertando Lin. Factura: " & cadMen
            
            '++monica:añadida la parte de insertar en tesoreria
            If B Then
                '[Monica]30/09/2011: como tenia hecha la funcion de insertar en tesoreria para todos,
                '                    la aprovecho y le paso como parametro el recordset Rs1
                Sql = "select * from rrecibpozos where " & cadWHERE
                Set RS1 = New ADODB.Recordset
                RS1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                '[Monica]18/07/2014: añadida la funcion de si es contado
                If TipoM = "RMT" And EsContado(cadWHERE) Then
                    B = InsertarAsientoCobroPOZOS(cadMen, RS1, CDate(FecVen), CtaBan)
                Else
                    B = InsertarEnTesoreriaPOZOS(cadMen, RS1, CDate(FecVen), Forpa, CtaBan)
                End If
                cadMen = "Insertando en Tesoreria: " & cadMen
                
                If B Then
                    If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
                End If
                
                
                Set RS1 = Nothing
            End If
        End If
            '++

    Else
        ' No insertamos la factura sino un asiento en el diario
        Set Mc = New Contadores
        
        If Mc.ConseguirContador("0", (CDate(FecFac) <= CDate(vEmpresa.FechaFin)), True) = 0 Then
        
            Obs = "Contabilización Factura Interna de Fecha " & Format(FecFac, "dd/mm/yyyy")
        
            'Insertar en la conta Cabecera Asiento
            B = InsertarCabAsientoDia(vEmpresa.NumDiarioInt, Mc.Contador, CStr(Format(FecFac, "dd/mm/yyyy")), Obs, cadMen)
            cadMen = "Insertando Cab. Asiento: " & cadMen
        Else
            B = False
        End If
    
        If B Then
            Socio = DevuelveValor("select codsocio from rrecibpozos where " & cadWHERE)
            CtaSocio = DevuelveValor("select codmaccli from rsocios_seccion where codsocio = " & Socio & " and codsecci = " & vParamAplic.SeccionPOZOS)
        
        
            B = InsertarLinAsientoFactIntPOZ("rrecibpozos", cadWHERE, cadMen, CtaSocio, Mc.Contador)
            cadMen = "Insertando Lin. Factura Interna: " & cadMen
        
            Set Mc = Nothing
        End If
    
        '++monica:añadida la parte de insertar en tesoreria
        If B Then
            CCoste = CodCCost
            Sql = "select * from rrecibpozos where " & cadWHERE
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            B = InsertarEnTesoreriaPOZOS(cadMen, RS1, CDate(FecVen), Forpa, CtaBan)
            
            cadMen = "Insertando en Tesoreria: " & cadMen
            Set RS1 = Nothing
        End If
    End If
    
    If B Then
        'Poner intconta=1 en ariagro.facturas
        B = ActualizarCabFact("rrecibpozos", cadWHERE, cadMen)
        cadMen = "Actualizando Factura: " & cadMen
    End If
    
    
EContab:
    
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Contabilizando Recibos Pozos", Err.Description
    End If
    If B Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaPOZOS = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaPOZOS = False
        
        Sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
        Sql = Sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "rrecibpozos", "tmpFactu")
        conn.Execute Sql
    End If
End Function

Private Function InsertarCabFactPOZ(cadWHERE As String, Observac As String, cadErr As String, FP As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim ImporIva As Currency

Dim Sql2 As String
Dim CadenaInsertFaclin2 As String



    On Error GoTo EInsertar
    
    Sql = "SELECT stipom.letraser,numfactu,fecfactu, rsocios_seccion.codmaccli as codmacta,year(fecfactu) as anofaccl,tipoiva,porc_iva,rrecibpozos.codsocio,"
    Sql = Sql & "sum(baseimpo) baseimpo, sum(imporiva) imporiva, sum(totalfact) totalfact "
    Sql = Sql & " FROM ((" & "rrecibpozos inner join " & "usuarios.stipom on rrecibpozos.codtipom=usuarios.stipom.codtipom) "
    Sql = Sql & "INNER JOIN rsocios ON rrecibpozos.codsocio=rsocios.codsocio) "
    Sql = Sql & "INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionPOZOS, "N")
    Sql = Sql & " WHERE " & cadWHERE
    Sql = Sql & " group by 1,2,3,4,5,6,7,8 "
    Sql = Sql & " order by 1,2,3,4,5,6,7,8 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
        Dim vSoc As cSocio
        Set vSoc = New cSocio
        
        
        If vSoc.LeerDatos(DBLet(Rs!Codsocio, "N")) Then
            vContaFra.NumeroFactura = DBLet(Rs!numfactu)
            vContaFra.Anofac = DBLet(Rs!anofaccl)
            vContaFra.Serie = DBLet(Rs!letraser)
        
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            DtoPPago = 0
            DtoGnral = 0
            BaseImp = Rs!baseimpo
            '[Monica]08/06/2016: para el caso de utxera y escalona lo saco de lo que ya tenia calculado en la factura
            '                    porque el totalfac lleva incluido el iva
            If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                ImporIva = DBLet(Rs!ImporIva, "N")
                TotalFac = DBLet(Rs!TotalFact, "N")
            Else
            ' en otro caso se queda como estaba
                '---- Laura 10/10/2006:  añadir el totalfac para utilizarlo en insertar lineas
                ImporIva = Round2((DBLet(BaseImp, "N") * DBLet(Rs!porc_iva, "N") / 100), 2)
                TotalFac = BaseImp + ImporIva
                '----
            End If
            
            Sql = ""
            Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!fecfactu) & "," & DBSet(Observac, "T") & ","
            
            If vParamAplic.ContabilidadNueva Then
                ' para el caso de las rectificativas
                Dim vTipM As String
                vTipM = DevuelveValor("select codtipom from stipom where letraser = " & DBSet(Rs!letraser, "T"))
                If vTipM = "FAR" Then
                    Sql = Sql & "'D',"
                Else
                    Sql = Sql & "'0',"
                End If
                
                Sql = Sql & "0," & DBSet(FP, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(ImporIva, "N") & ","
                Sql = Sql & ValorNulo & "," & DBSet(TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,"
                Sql = Sql & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                Sql = Sql & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES',1"
            
                cad = cad & "(" & Sql & ")"
            
            
                Sql = "INSERT INTO factcli (numserie,numfactu,fecfactu,fecliqcl,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
                Sql = Sql & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
                Sql = Sql & "codpais,codagente)"
                Sql = Sql & " VALUES " & cad
                ConnConta.Execute Sql
        '***
                CadenaInsertFaclin2 = ""
                    
                
                'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
                'IVA 1, siempre existe
                Sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
                Sql2 = Sql2 & "1," & DBSet(Rs!baseimpo, "N") & "," & Rs!TipoIVA & "," & DBSet(Rs!porc_iva, "N") & ","
                Sql2 = Sql2 & ValorNulo & "," & DBSet(ImporIva, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
            
                Sql = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
                Sql = Sql & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
                ConnConta.Execute Sql
            
                'para las lineas
                vTipoIva(0) = Rs!TipoIVA
                vPorcIva(0) = Rs!porc_iva
                vPorcRec(0) = 0
                vImpIva(0) = Rs!ImporIva
                vImpRec(0) = 0
                vBaseIva(0) = Rs!baseimpo
                
                vTipoIva(1) = 0: vTipoIva(2) = 0
        
            Else
            
                Sql = Sql & DBSet(Rs!baseimpo, "N") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!porc_iva, "N") & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(ImporIva, "N", "N") & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & DBSet(TotalFac, "N") & "," & DBSet(Rs!TipoIVA, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
                cad = cad & "(" & Sql & ")"
            
                'Insertar en la contabilidad
                Sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,fecliqcl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
                Sql = Sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
                Sql = Sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien) "
                Sql = Sql & " VALUES " & cad
                ConnConta.Execute Sql
            
            End If
        End If
'        RS.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
    
    Set vSoc = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactPOZ = False
        cadErr = Err.Description
    Else
        InsertarCabFactPOZ = True
    End If
End Function



Private Function InsertarLinFactPOZ(cadTabla As String, cadWHERE As String, cadErr As String, tipoMov As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpConsumo As Currency, ImpCuota As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String

    On Error GoTo EInLinea
    

    If vParamAplic.Cooperativa = 7 Then ' si la cooperativa es quatretonda
        If vEmpresa.TieneAnalitica Then
            Sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu,sum(round(precio1*consumo1,2)) as importeconsumo,sum(round(precio2*consumo2,2) + impcuota) as importecuota, " & DBSet(vParamAplic.CodCCostPOZ, "T") & " as codccost "
        Else
            Sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu,sum(round(precio1*consumo1,2)) as importeconsumo,sum(round(precio2*consumo2,2) + impcuota) as importecuota "
        End If
        
        Sql = Sql & " FROM rrecibpozos inner join usuarios.stipom on rrecibpozos.codtipom=usuarios.stipom.codtipom "
        Sql = Sql & " WHERE " & cadWHERE
        
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & " GROUP BY 1,2,3,4,7 " '& cadCampo, codccost
        Else
            Sql = Sql & " GROUP BY 1,2,3,4" '& cadCampo
        End If
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        cad = ""
        I = 1
        totimp = 0
        SQLaux = ""
        If Not Rs.EOF Then
            SQLaux = cad
            
            ImpConsumo = DBLet(Rs!Importeconsumo, "N")
            ImpCuota = DBLet(Rs!importecuota, "N")
            totimp = totimp + ImpConsumo + ImpCuota
    
            'concatenamos linea para insertar en la tabla de conta.linfact
            Sql = ""
            Sql2 = ""
            
            
            If ImpConsumo <> 0 Then
                Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & DBSet(I, "N") & "," & DBSet(vParamAplic.CtaVentasConsPOZ, "T") & ","
                
                Sql2 = cad & Sql  'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
                Sql = Sql & DBSet(ImpConsumo, "N") & ","
                
                If vEmpresa.TieneAnalitica Then
                    Sql = Sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBSet(Rs!CodCCost, "T")
                Else
                    Sql = Sql & ValorNulo
                    CCoste = ValorNulo
                End If
                
                cad = "(" & Sql & ")" & ","
                
                SQLaux = SQLaux & cad
                
                I = I + 1
            End If
            
            
            If ImpCuota <> 0 Then
                Sql = "('" & Rs!letraser & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & DBSet(I, "N") & "," & DBSet(vParamAplic.CtaVentasCuoPOZ, "T") & ","
                
                Sql2 = cad & Sql   'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
                Sql = Sql & DBSet(ImpCuota, "N") & ","
                
                If vEmpresa.TieneAnalitica Then
                    Sql = Sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBSet(Rs!CodCCost, "T")
                Else
                    Sql = Sql & ValorNulo
                    CCoste = ValorNulo
                End If
                
                cad = cad & Sql & ")" & ","
                
                SQLaux = SQLaux & cad
            End If
            
            Rs.MoveNext
        End If
        
        Rs.Close
        Set Rs = Nothing
        
        BaseImp = DevuelveValor("select sum(baseimpo) from rrecibpozos where " & cadWHERE)
        'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
        'de la factura
        If totimp <> BaseImp Then
    '        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
            'en SQL esta la ult linea introducida
            totimp = BaseImp - totimp
            totimp = ImpCuota + totimp '(+- diferencia)
            Sql2 = Sql2 & DBSet(totimp, "N") & ","
            If CCoste = "" Or CCoste = ValorNulo Then
                Sql2 = Sql2 & ValorNulo
            Else
                Sql2 = Sql2 & DBSet(CCoste, "T")
            End If
'            If SQLaux <> "" Then 'hay mas de una linea
'                cad = SQLaux & Sql2 & ")" & ","
'            Else 'solo una linea
'                cad = SQLaux & ")" & ","
'            End If
            cad = Sql2 & "),"
        End If
    
    
        'Insertar en la contabilidad
        If cad <> "" Then
            cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
            Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
            Sql = Sql & " VALUES " & cad
            ConnConta.Execute Sql
        End If
    Else
        ' la cooperativa es utxera o escalona
        ' Dependiendo del tipo de movimiento cambiamos la cta de venta
        '[Monica]15/01/2016: añadimos los tipos de movimientos rectificativos RTA,RRM,RRV,RRT,RRC (DE CONSUMO)
        Select Case tipoMov
            Case "TAL", "RTA"
                ' Recibos de talla
                cadCampo = vParamAplic.CtaVentasTalPOZ
            Case "RMP", "RVP", "RRM", "RRV" '[Monica]28/06/2013: añadido el tipo de movimiento de contadores
                ' Recibos de mantenimiento o de contadores
                cadCampo = vParamAplic.CtaVentasMtoPOZ
            Case "RMT", "RRT"
                ' Recibos de consumo a manta
                cadCampo = vParamAplic.CtaVentasMantaPOZ
            Case Else
                ' Recibos de consumo y de contadores de momento
                cadCampo = vParamAplic.CtaVentasConsPOZ
        End Select
        
        If vEmpresa.TieneAnalitica Then
            Sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(baseimpo-coalesce(imprecargo,0)) as importe, " & DBSet(vParamAplic.CodCCostPOZ, "T")
        Else
            Sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(baseimpo-coalesce(imprecargo,0)) as importe "
        End If
        
        Sql = Sql & " FROM rrecibpozos inner join usuarios.stipom on rrecibpozos.codtipom=usuarios.stipom.codtipom "
        Sql = Sql & " WHERE " & cadWHERE
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & " GROUP BY 1,2,3,4,6 " '& cadCampo, codccost
        Else
            Sql = Sql & " GROUP BY 1,2,3,4 " '& cadCampo
        End If
        
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        cad = ""
        I = 1
        totimp = 0
        SQLaux = ""
        While Not Rs.EOF
            SQLaux = cad
            'calculamos la Base Imp del total del importe para cada cta cble ventas
            ImpLinea = DBLet(Rs!Importe, "N")
            totimp = totimp + DBLet(Rs!Importe, "N")
    
            'concatenamos linea para insertar en la tabla de conta.linfact
            Sql = ""
            Sql2 = ""
            
            Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
            Sql = Sql & DBSet(Rs!cuenta, "T")
            
            Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
            Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
            
            If vEmpresa.TieneAnalitica Then
                Sql = Sql & DBSet(Rs!CodCCost, "T")
                CCoste = DBSet(Rs!CodCCost, "T")
            Else
                Sql = Sql & ValorNulo
                CCoste = ValorNulo
            End If
            
            cad = cad & "(" & Sql & ")" & ","
            
            I = I + 1
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        '[Monica]21/01/2016: faltaria añadir el recargo
        cadCampo = vParamAplic.CtaRecargosPOZ
        
        If vEmpresa.TieneAnalitica Then
            Sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(coalesce(imprecargo,0)) as importe, " & DBSet(vParamAplic.CodCCostPOZ, "T")
        Else
            Sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(coalesce(imprecargo,0)) as importe "
        End If
        
        Sql = Sql & " FROM rrecibpozos inner join usuarios.stipom on rrecibpozos.codtipom=usuarios.stipom.codtipom "
        Sql = Sql & " WHERE " & cadWHERE
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & " GROUP BY 1,2,3,4,6 " '& cadCampo, codccost
        Else
            Sql = Sql & " GROUP BY 1,2,3,4 " '& cadCampo
        End If
        
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        While Not Rs.EOF
            If DBLet(Rs!Importe, "N") <> 0 Then
                SQLaux = cad
                'calculamos la Base Imp del total del importe para cada cta cble ventas
                ImpLinea = DBLet(Rs!Importe, "N")
                totimp = totimp + DBLet(Rs!Importe, "N")
        
                'concatenamos linea para insertar en la tabla de conta.linfact
                Sql = ""
                Sql2 = ""
                
                Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
                Sql = Sql & DBSet(Rs!cuenta, "T")
                
                Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
                Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
                
                If vEmpresa.TieneAnalitica Then
                    Sql = Sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBSet(Rs!CodCCost, "T")
                Else
                    Sql = Sql & ValorNulo
                    CCoste = ValorNulo
                End If
                
                cad = cad & "(" & Sql & ")" & ","
                
                I = I + 1
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        'hasta aquí
        
        
        
        'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
        'de la factura
        If totimp <> BaseImp Then
    '        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
            'en SQL esta la ult linea introducida
            totimp = BaseImp - totimp
            totimp = ImpLinea + totimp '(+- diferencia)
            Sql2 = Sql2 & DBSet(totimp, "N") & ","
            If CCoste = "" Or CCoste = ValorNulo Then
                Sql2 = Sql2 & ValorNulo
            Else
                Sql2 = Sql2 & DBSet(CCoste, "T")
            End If
            If SQLaux <> "" Then 'hay mas de una linea
                cad = SQLaux & "(" & Sql2 & ")" & ","
            Else 'solo una linea
                cad = "(" & Sql2 & ")" & ","
            End If
            
    '        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
    '        cad = Replace(cad, SQL, Aux)
        End If
    
    
        'Insertar en la contabilidad
        If cad <> "" Then
            cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
            Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
            Sql = Sql & " VALUES " & cad
            ConnConta.Execute Sql
        End If
    
    End If
EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactPOZ = False
        cadErr = Err.Description
    Else
        InsertarLinFactPOZ = True
    End If
End Function



Private Function InsertarLinFactPOZContaNueva(cadTabla As String, cadWHERE As String, cadErr As String, tipoMov As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpConsumo As Currency, ImpCuota As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String
Dim ImpIvaAux As Currency

Dim TotImpIVA As Currency
Dim vImpIvaAux As Currency


Dim NumeroIVA As Byte
Dim k As Integer
Dim HayQueAjustar As Boolean

Dim ImpImva As Currency
Dim ImpREC As Currency




    On Error GoTo EInLinea
    

    If vParamAplic.Cooperativa = 7 Then ' si la cooperativa es quatretonda
        If vEmpresa.TieneAnalitica Then
            Sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu,rrecibpozos.tipoiva, rrecibpozos.porc_iva, sum(round(precio1*consumo1,2)) as importeconsumo,sum(round(precio2*consumo2,2) + impcuota) as importecuota, " & DBSet(vParamAplic.CodCCostPOZ, "T") & " as codccost "
        Else
            Sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu,rrecibpozos.tipoiva, rrecibpozos.porc_iva, sum(round(precio1*consumo1,2)) as importeconsumo,sum(round(precio2*consumo2,2) + impcuota) as importecuota "
        End If
        
        Sql = Sql & " FROM rrecibpozos inner join usuarios.stipom on rrecibpozos.codtipom=usuarios.stipom.codtipom "
        Sql = Sql & " WHERE " & cadWHERE
        
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & " GROUP BY 1,2,3,4,5,6,9 " '& cadCampo, codccost
        Else
            Sql = Sql & " GROUP BY 1,2,3,4,5,6" '& cadCampo
        End If
        Sql = Sql & " ORDER BY 5 "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        cad = ""
        I = 1
        totimp = 0
        TotImpIVA = 0
        
        SQLaux = ""
        If Not Rs.EOF Then
            SQLaux = cad
            
            ImpConsumo = DBLet(Rs!Importeconsumo, "N")
            ImpCuota = DBLet(Rs!importecuota, "N")
            totimp = totimp + ImpConsumo + ImpCuota
    
            vImpIvaAux = vImpIva(0)
    
            'concatenamos linea para insertar en la tabla de conta.linfact
            Sql = ""
            Sql2 = ""
            
            If ImpConsumo <> 0 Then
                Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & DBSet(I, "N") & "," & DBSet(vParamAplic.CtaVentasConsPOZ, "T") & ","
                
                If vEmpresa.TieneAnalitica Then
                    Sql = Sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBSet(Rs!CodCCost, "T")
                Else
                    Sql = Sql & ValorNulo
                    CCoste = ValorNulo
                End If
                
                Sql = Sql & "," & DBSet(Rs!fecfactu, "F")
                Sql = Sql & "," & DBSet(Rs!TipoIVA, "N") & "," & DBSet(Rs!porc_iva, "N") & "," & ValorNulo
                
                Sql2 = Sql & ","  'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
                
                ImpLinea = ImpConsumo
                
                'Calculo el importe de IVA y el de recargo de equivalencia
                ImpImva = vPorcIva(NumeroIVA) / 100
                ImpImva = Round2(ImpLinea * ImpImva, 2)
                If vPorcRec(NumeroIVA) = 0 Then
                    ImpREC = 0
                Else
                    ImpREC = vPorcRec(NumeroIVA) / 100
                    ImpREC = Round2(ImpLinea * ImpREC, 2)
                End If
                vImpIva(NumeroIVA) = vImpIva(NumeroIVA) - ImpImva
                vImpRec(NumeroIVA) = vImpRec(NumeroIVA) - ImpREC
                
                TotImpIVA = TotImpIVA + ImpImva
                
                
                ' baseimpo , impoiva, imporec, aplicret, CodCCost
                Sql = Sql & "," & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpREC, "N", "S") & ",0"
                
                cad = "(" & Sql & ")" & ","
                
                SQLaux = SQLaux & cad
                
                I = I + 1
            End If
            
            
            If ImpCuota <> 0 Then
                Sql = "('" & Rs!letraser & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & DBSet(I, "N") & "," & DBSet(vParamAplic.CtaVentasCuoPOZ, "T") & ","
                
                If vEmpresa.TieneAnalitica Then
                    Sql = Sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBSet(Rs!CodCCost, "T")
                Else
                    Sql = Sql & ValorNulo
                    CCoste = ValorNulo
                End If
                
                Sql = Sql & "," & DBSet(Rs!fecfactu, "F")
                Sql = Sql & "," & DBSet(Rs!TipoIVA, "N") & "," & DBSet(Rs!porc_iva, "N") & "," & ValorNulo
                
                Sql2 = Sql & ","   'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
                
                
                ImpLinea = ImpCuota
                
                'Calculo el importe de IVA y el de recargo de equivalencia
                ImpImva = vPorcIva(NumeroIVA) / 100
                ImpImva = Round2(ImpLinea * ImpImva, 2)
                If vPorcRec(NumeroIVA) = 0 Then
                    ImpREC = 0
                Else
                    ImpREC = vPorcRec(NumeroIVA) / 100
                    ImpREC = Round2(ImpLinea * ImpREC, 2)
                End If
                vImpIva(NumeroIVA) = vImpIva(NumeroIVA) - ImpImva
                vImpRec(NumeroIVA) = vImpRec(NumeroIVA) - ImpREC
                
                
                ' baseimpo , impoiva, imporec, aplicret, CodCCost
                Sql = Sql & "," & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpREC, "N", "S") & ",0"
                
                cad = "(" & Sql & ")" & ","
                
                TotImpIVA = TotImpIVA + ImpImva
                
                SQLaux = SQLaux & cad
                
            End If
            
            
            Rs.MoveNext
        End If
        
        Rs.Close
        Set Rs = Nothing
        
        'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
        'de la factura
        If TotImpIVA <> vImpIvaAux Then
    '        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
            'en SQL esta la ult linea introducida
            totimp = vImpIvaAux - TotImpIVA
            totimp = ImpImva + totimp '(+- diferencia)
            Sql2 = Sql2 & DBSet(ImpLinea, "N") & "," & DBSet(totimp, "N") & "," & DBSet(ImpREC, "N")
            
            If SQLaux <> "" Then 'hay mas de una linea
                cad = SQLaux & Sql2 & ")" & ","
            Else 'solo una linea
                cad = SQLaux & ")" & ","
            End If
        End If
    
    
        'Insertar en la contabilidad
        If cad <> "" Then
            cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
            Sql = "INSERT INTO factcli_lineas(numserie,numfactu,anofactu,numlinea,codmacta,codccost,fecfactu,codigiva,porciva,porcrec,baseimpo,impoiva,imporec)"
            Sql = Sql & " VALUES " & cad
            ConnConta.Execute Sql
        End If
    Else
        ' la cooperativa es utxera o escalona
        ' Dependiendo del tipo de movimiento cambiamos la cta de venta
        '[Monica]15/01/2016: añadimos los tipos de movimientos rectificativos RTA,RRM,RRV,RRT,RRC (DE CONSUMO)
        Select Case tipoMov
            Case "TAL", "RTA"
                ' Recibos de talla
                cadCampo = vParamAplic.CtaVentasTalPOZ
            Case "RMP", "RVP", "RRM", "RRV" '[Monica]28/06/2013: añadido el tipo de movimiento de contadores
                ' Recibos de mantenimiento o de contadores
                cadCampo = vParamAplic.CtaVentasMtoPOZ
            Case "RMT", "RRT"
                ' Recibos de consumo a manta
                cadCampo = vParamAplic.CtaVentasMantaPOZ
            Case Else
                ' Recibos de consumo y de contadores de momento
                cadCampo = vParamAplic.CtaVentasConsPOZ
        End Select
        
        If vEmpresa.TieneAnalitica Then
            Sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,rrecibpozos.tipoiva, rrecibpozos.porc_iva,sum(baseimpo-coalesce(imprecargo,0)) as importe, " & DBSet(vParamAplic.CodCCostPOZ, "T")
        Else
            Sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,rrecibpozos.tipoiva, rrecibpozos.porc_iva,sum(baseimpo-coalesce(imprecargo,0)) as importe "
        End If
        
        Sql = Sql & " FROM rrecibpozos inner join usuarios.stipom on rrecibpozos.codtipom=usuarios.stipom.codtipom "
        Sql = Sql & " WHERE " & cadWHERE
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & " GROUP BY 1,2,3,4,5,6,8 " '& cadCampo, codccost
        Else
            Sql = Sql & " GROUP BY 1,2,3,4,5,6 " '& cadCampo
        End If
        
        
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        cad = ""
        I = 1
        totimp = 0
        SQLaux = ""
        TotImpIVA = 0
        vImpIvaAux = vImpIva(0)
        
        While Not Rs.EOF
            SQLaux = cad
            'calculamos la Base Imp del total del importe para cada cta cble ventas
            ImpLinea = DBLet(Rs!Importe, "N")
            totimp = totimp + DBLet(Rs!Importe, "N")
    
            'concatenamos linea para insertar en la tabla de conta.linfact
            Sql = ""
            Sql2 = ""
            
            Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
            Sql = Sql & DBSet(Rs!cuenta, "T")
            Sql = Sql & ","
            
            If vEmpresa.TieneAnalitica Then
                Sql = Sql & DBSet(Rs!CodCCost, "T")
                CCoste = DBSet(Rs!CodCCost, "T")
            Else
                Sql = Sql & ValorNulo
                CCoste = ValorNulo
            End If
            
            Sql = Sql & "," & DBSet(Rs!fecfactu, "F")
            
            
            Sql = Sql & "," & DBSet(Rs!TipoIVA, "N") & "," & DBSet(Rs!porc_iva, "N") & "," & ValorNulo
            Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
            
            'Calculo el importe de IVA y el de recargo de equivalencia
            ImpImva = vPorcIva(NumeroIVA) / 100
            ImpImva = Round2(ImpLinea * ImpImva, 2)
            If vPorcRec(NumeroIVA) = 0 Then
                ImpREC = 0
            Else
                ImpREC = vPorcRec(NumeroIVA) / 100
                ImpREC = Round2(ImpLinea * ImpREC, 2)
            End If
            vImpIva(NumeroIVA) = vImpIva(NumeroIVA) - ImpImva
            vImpRec(NumeroIVA) = vImpRec(NumeroIVA) - ImpREC
            
            ' baseimpo , impoiva, imporec, aplicret, CodCCost
            Sql = Sql & "," & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpREC, "N", "S") & ",0"
            
            cad = cad & "(" & Sql & ")" & ","
            
            TotImpIVA = TotImpIVA + ImpImva
            
            SQLaux = SQLaux & cad
            
            I = I + 1
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        
        '[Monica]21/01/2016: faltaria añadir el recargo
        cadCampo = vParamAplic.CtaRecargosPOZ
        
        If vEmpresa.TieneAnalitica Then
            Sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,rrecibpozos.tipoiva, rrecibpozos.porc_iva,sum(coalesce(imprecargo,0)) as importe, " & DBSet(vParamAplic.CodCCostPOZ, "T")
        Else
            Sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,rrecibpozos.tipoiva, rrecibpozos.porc_iva,sum(coalesce(imprecargo,0)) as importe "
        End If
        
        Sql = Sql & " FROM rrecibpozos inner join usuarios.stipom on rrecibpozos.codtipom=usuarios.stipom.codtipom "
        Sql = Sql & " WHERE " & cadWHERE
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & " GROUP BY 1,2,3,4,5,6,8 " '& cadCampo, codccost
        Else
            Sql = Sql & " GROUP BY 1,2,3,4,5,6 " '& cadCampo
        End If
        
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        While Not Rs.EOF
            If DBLet(Rs!Importe, "N") <> 0 Then
                SQLaux = cad
                'calculamos la Base Imp del total del importe para cada cta cble ventas
                ImpLinea = DBLet(Rs!Importe, "N")
                totimp = totimp + DBLet(Rs!Importe, "N")
        
                'concatenamos linea para insertar en la tabla de conta.linfact
                Sql = ""
                Sql2 = ""
                
                Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
                Sql = Sql & DBSet(Rs!cuenta, "T") & ","
                
                If vEmpresa.TieneAnalitica Then
                    Sql = Sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBSet(Rs!CodCCost, "T")
                Else
                    Sql = Sql & ValorNulo
                    CCoste = ValorNulo
                End If
                Sql = Sql & "," & DBSet(Rs!fecfactu, "F")
                
                Sql = Sql & "," & DBSet(Rs!TipoIVA, "N") & "," & DBSet(Rs!porc_iva, "N") & "," & ValorNulo
                Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
                
                'Calculo el importe de IVA y el de recargo de equivalencia
                ImpImva = vPorcIva(NumeroIVA) / 100
                ImpImva = Round2(ImpLinea * ImpImva, 2)
                If vPorcRec(NumeroIVA) = 0 Then
                    ImpREC = 0
                Else
                    ImpREC = vPorcRec(NumeroIVA) / 100
                    ImpREC = Round2(ImpLinea * ImpREC, 2)
                End If
                vImpIva(NumeroIVA) = vImpIva(NumeroIVA) - ImpImva
                vImpRec(NumeroIVA) = vImpRec(NumeroIVA) - ImpREC
                
                ' baseimpo , impoiva, imporec, aplicret, CodCCost
                Sql = Sql & "," & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpREC, "N", "S") & ",0"
                
                cad = cad & "(" & Sql & ")" & ","
                
                TotImpIVA = TotImpIVA + ImpImva
                
                SQLaux = SQLaux & cad
                
                I = I + 1
                
            End If
            
            
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        'hasta aquí
        
        
        
        'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
        'de la factura
        If TotImpIVA <> vImpIvaAux Then
    '        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
            'en SQL esta la ult linea introducida
            totimp = vImpIvaAux - TotImpIVA
            totimp = ImpImva + totimp '(+- diferencia)
            Sql2 = Sql2 & DBSet(ImpLinea, "N") & "," & DBSet(totimp, "N") & "," & DBSet(ImpREC, "N")
            
            If SQLaux <> "" Then 'hay mas de una linea
                cad = SQLaux & Sql2 & ")" & ","
            Else 'solo una linea
                cad = SQLaux & ")" & ","
            End If
        End If
    
    
        'Insertar en la contabilidad
        If cad <> "" Then
            cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
            Sql = "INSERT INTO factcli_lineas(numserie,numfactu,anofactu,numlinea,codmacta,codccost,fecfactu,codigiva,porciva,porcrec,baseimpo,impoiva,imporec)"
            Sql = Sql & " VALUES " & cad
            ConnConta.Execute Sql
        End If
    
    End If
EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactPOZContaNueva = False
        cadErr = Err.Description
    Else
        InsertarLinFactPOZContaNueva = True
    End If
End Function





'###########################CONTABILIZACION DE FACTURAS DE TRANSPORTE INTERNAS


Private Function InsertarAsientoDiarioTRANS(cadWHERE As String, cadErr As String, ByRef Mc As Contadores, FechaFin As Date, Secci As String, Tipo As Byte, FecRec As Date, TipoM As String) As Boolean
' la contabilizacion de las facturas internas de horto, se han de insertar en el diario no en el registro de iva de proveedor
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Concepto As String
Dim letraser As String
Dim Obs As String
Dim cadMen As String
Dim B As Boolean
'Dim CtaSocio As String


    On Error GoTo EInsertar
       
    Sql = " SELECT fecfactu,year(" & DBSet(FecRec, "F") & ") as anofacpr," & DBSet(FecRec, "F") & ",numfactu,rtransporte.codmacpro,"
    Sql = Sql & "baseimpo, tipoiva, porc_iva,imporiva,basereten,porc_ret,impreten,baseaport,porc_apo,impapor,totalfac,"
    Sql = Sql & "rectif_codtipom, rectif_numfactu, rectif_fecfactu,"
    Sql = Sql & "rtransporte.codtrans, rtransporte.nomtrans, rtransporte.codbanco, rtransporte.codsucur, rtransporte.digcontr, rtransporte.cuentaba "
    Sql = Sql & ",rtransporte.iban "
    Sql = Sql & " FROM (" & "rfacttra "
    Sql = Sql & "INNER JOIN rtransporte ON rfacttra.codtrans=rtransporte.codtrans) "
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
        If Mc.ConseguirContador("1", (FecRec <= CDate(FechaFin) - 365), True) = 0 Then
        
            BaseImp = DBLet(Rs!baseimpo, "N")
            TotalFac = BaseImp + DBLet(Rs!ImporIva, "N")
            AnyoFacPr = Rs!anofacpr
            
            ImpReten = DBLet(Rs!ImpReten, "N")
            
            letraser = ""
            letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(TipoM, "T"))
            
            FacturaTRA = letraser & "-" & DBLet(Rs!numfactu, "N")
            FecFactuTRA = DBLet(Rs!fecfactu, "F")
            
            CodTipomRECT = DBLet(Rs!rectif_codtipom, "T")
            NumfactuRECT = DBLet(Rs!rectif_numfactu, "T")
            FecfactuRECT = DBLet(Rs!rectif_fecfactu, "T")
            
            CtaTransporte = Rs!codmacpro
            TipoFact = Tipo
            FecRecep = FecRec
            BancoTRA = DBLet(Rs!CodBanco, "N")
            SucurTRA = DBLet(Rs!CodSucur, "N")
            DigcoTRA = DBLet(Rs!digcontr, "T")
            CtaBaTRA = DBLet(Rs!CuentaBa, "T")
            IbanTRA = DBLet(Rs!Iban, "T")
            TotalTesor = DBLet(Rs!TotalFac, "N")
            
'            Variedades = VariedadesFactura(cadWhere)
            
            Obs = "Contabilización Factura Interna de Transporte de Fecha " & Format(FecRecep, "dd/mm/yyyy")
        
            'Insertar en la conta Cabecera Asiento
            B = InsertarCabAsientoDia(vEmpresa.NumDiarioInt, Mc.Contador, CStr(Format(FecRecep, "dd/mm/yyyy")), Obs, cadMen)
            cadMen = "Insertando Cab. Asiento: " & cadMen
    
            If B Then
                B = InsertarLinAsientoFactIntTRA("rfacttra", cadWHERE, cadMen, Tipo, CtaTransporte, Mc.Contador)
                cadMen = "Insertando Lin. Factura Interna: " & cadMen
            
                Set Mc = Nothing
            End If
            
            FacturaTRA = DBLet(Rs!numfactu, "N")
            
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarAsientoDiarioTRANS = False
        cadErr = Err.Description
    Else
        InsertarAsientoDiarioTRANS = True
    End If
End Function





Private Function InsertarLinAsientoFactIntTRA(cadTabla As String, cadWHERE As String, cadErr As String, Tipo As Byte, CtaSocio As String, Optional Contador As Long) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim numdocum As String
Dim Ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim I As Long
Dim B As Boolean
Dim cad As String
Dim cadMen As String
Dim FeFact As Date

Dim cadCampo As String

Dim TipoAnt As Byte
Dim TipoFact As String

Dim totimp As Currency
Dim SQLaux As String
Dim ImpLinea As String
Dim Sql3 As String
Dim ImpAnticipo As Currency
Dim NumFact As Long

    On Error GoTo EInLinea
    
    InsertarLinAsientoFactIntTRA = False
    
    TipoAnt = Tipo
'    TipoFactAnt = TipoFact
    
    If Tipo = 11 Then ' si es una factura rectificativa cojo el tipo de movimiento de la factura que rectifico
        Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(CodTipomRECT, "T"))
        
        TipoFact = CodTipomRECT
    Else
        TipoFact = "FTR"
    End If
    
    FeFact = FecFactuTRA
    NumFact = DevuelveValor("select numfactu from rfacttra where " & cadWHERE)
    
    If vEmpresa.TieneAnalitica Then
        Sql = " SELECT 1, variedades.ctatransporte as cuenta, sum(rfacttra_albaran.importe) as importe, variedades.codccost "
    Else
        Sql = " SELECT 1, variedades.ctatransporte as cuenta, sum(rfacttra_albaran.importe) as importe "
    End If
    Sql = Sql & " FROM rfacttra_albaran, variedades "
    Sql = Sql & " WHERE " & Replace(cadWHERE, "rfacttra", "rfacttra_albaran") & " and"
    Sql = Sql & " rfacttra_albaran.codvarie = variedades.codvarie "
    Sql = Sql & " group by 1,2 "
    Sql = Sql & " order by 1,2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    I = 0
    ImporteD = 0
    ImporteH = 0
    
    numdocum = Format(NumFact, "0000000")
    Ampliacion = FacturaTRA
    ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoIntProv, "N")) & " " & Ampliacion
    ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoIntProv, "N")) & " " & Ampliacion
    
    If Not Rs.EOF Then Rs.MoveFirst
    
    B = True

    cad = ""
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        ImpLinea = Rs!Importe
        
        totimp = totimp + ImpLinea
        
        I = I + 1
        
        cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
        cad = cad & DBSet(I, "N") & "," & DBSet(Rs!cuenta, "T") & "," & DBSet(numdocum, "T") & ","
        
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If Rs.Fields(2).Value > 0 Then
            ' importe al debe en positivo
            cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Rs.Fields(2).Value, "N") & ","
            cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaTransporte, "T") & "," & ValorNulo & ",0"
        
            ImporteD = ImporteD + (CCur(Rs.Fields(2).Value))
        Else
            ' importe al haber en positivo, cambiamos el signo
            cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
            cad = cad & DBSet((Rs.Fields(2).Value) * (-1), "N") & "," & ValorNulo & "," & DBSet(CtaTransporte, "T") & "," & ValorNulo & ",0"
        
            ImporteH = ImporteH + (CCur(Rs.Fields(2).Value) * (-1))
        End If
        
        cad = "(" & cad & ")"
        
        B = InsertarLinAsientoDia(cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & I

        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)

        If ImpLinea > 0 Then
            If vParamAplic.ContabilidadNueva Then
                Sql = "update hlinapu set timporteD = " & DBSet(totimp, "N")
            Else
                Sql = "update linapu set timporteD = " & DBSet(totimp, "N")
            End If
            Sql = Sql & " where numdiari = " & DBSet(vEmpresa.NumDiarioInt, "N")
            Sql = Sql & " and fechaent = " & DBSet(FecRecep, "F")
            Sql = Sql & " and numasien = " & DBSet(Contador, "N")
            Sql = Sql & " and linliapu = " & DBSet(I, "N")
            
            ConnConta.Execute Sql
        Else
            If vParamAplic.ContabilidadNueva Then
                Sql = "update hlinapu set timporteH = " & DBSet(totimp, "N")
            Else
                Sql = "update linapu set timporteH = " & DBSet(totimp, "N")
            End If
            Sql = Sql & " where numdiari = " & DBSet(vEmpresa.NumDiarioInt, "N")
            Sql = Sql & " and fechaent = " & DBSet(FecRecep, "F")
            Sql = Sql & " and numasien = " & DBSet(Contador, "N")
            Sql = Sql & " and linliapu = " & DBSet(I, "N")
            
            ConnConta.Execute Sql
        End If
    End If

    If B And I > 0 Then
        I = I + 1
        
        ' el Total es sobre la cuenta del transportista
        cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
        cad = cad & DBSet(I, "N") & ","
        cad = cad & DBSet(CtaTransporte, "T") & "," & DBSet(numdocum, "T") & ","
            
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If ImporteD - ImporteH < 0 Then
            ' importe al debe en positivo
            cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet((ImporteD - ImporteH) * (-1), "N") & ","
            cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
        Else
            'importe al haber en positivo, cambiamos el signo
            cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
            cad = cad & DBSet((ImporteD - ImporteH), "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
        End If
        
        cad = "(" & cad & ")"
        
        B = InsertarLinAsientoDia(cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & I
        
    End If

    If B Then
        ' las retenciones si las hay
        If ImpReten <> 0 Then
            I = I + 1
            
            cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
            cad = cad & DBSet(I, "N") & "," & DBSet(CtaTransporte, "T") & "," & DBSet(numdocum, "T") & ","
        
            If ImpReten > 0 Then
                ' importe al debe en positivo
                cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(ImpReten, "N") & ","
                cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaReten, "T") & "," & ValorNulo & ",0"
            Else
                ' importe al haber en positivo, cambiamos el signo
                cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                cad = cad & DBSet((ImpReten * (-1)), "N") & "," & ValorNulo & "," & DBSet(CtaReten, "T") & "," & ValorNulo & ",0"
            
            End If
            
            cad = "(" & cad & ")"
            
            B = InsertarLinAsientoDia(cad, cadMen)
            cadMen = "Insertando Lin. Asiento: " & I
            
            If B Then
                I = I + 1
                
                cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
                cad = cad & DBSet(I, "N") & "," & DBSet(CtaReten, "T") & "," & DBSet(numdocum, "T") & ","
                If ImpReten > 0 Then
                    ' importe al haber en positivo
                    cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet(ImpReten, "N") & "," & ValorNulo & "," & DBSet(CtaTransporte, "T") & "," & ValorNulo & ",0"
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & DBSet((ImpReten * (-1)), "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaTransporte, "T") & "," & ValorNulo & ",0"
                End If
                
                cad = "(" & cad & ")"
                
                B = InsertarLinAsientoDia(cad, cadMen)
                cadMen = "Insertando Lin. Asiento: " & I
            End If
            
        End If
    End If
    
    
    If B Then
        ' las aportaciones de fondo operativo si las hay
        If ImpAport <> 0 Then
            I = I + 1
            
            cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
            cad = cad & DBSet(I, "N") & "," & DBSet(CtaTransporte, "T") & "," & DBSet(numdocum, "T") & ","
        
            If ImpAport > 0 Then
                ' importe al debe en positivo
                cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(ImpAport, "N") & ","
                cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaAport, "T") & "," & ValorNulo & ",0"
            Else
                ' importe al haber en positivo, cambiamos el signo
                cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                cad = cad & DBSet((ImpAport * (-1)), "N") & "," & ValorNulo & "," & DBSet(CtaAport, "T") & "," & ValorNulo & ",0"
            
            End If
            
            cad = "(" & cad & ")"
            
            B = InsertarLinAsientoDia(cad, cadMen)
            cadMen = "Insertando Lin. Asiento: " & I
            
            If B Then
                I = I + 1
                
                cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
                cad = cad & DBSet(I, "N") & "," & DBSet(CtaAport, "T") & "," & DBSet(numdocum, "T") & ","
                If ImpAport > 0 Then
                    ' importe al haber en positivo
                    cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet(ImpAport, "N") & "," & ValorNulo & "," & DBSet(CtaTransporte, "T") & "," & ValorNulo & ",0"
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    cad = cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & DBSet((ImpAport * (-1)), "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaTransporte, "T") & "," & ValorNulo & ",0"
                End If
                
                cad = "(" & cad & ")"
                
                B = InsertarLinAsientoDia(cad, cadMen)
                cadMen = "Insertando Lin. Asiento: " & I
            End If
        End If
    End If
    
    Tipo = TipoAnt

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinAsientoFactIntTRA = False
        cadErr = Err.Description
    Else
        InsertarLinAsientoFactIntTRA = True
    End If
    Set Rs = Nothing
    InsertarLinAsientoFactIntTRA = B
    Exit Function
End Function




Public Function PasarAsientoGastoCampo(cadWHERE As String, FechaFin As String, FecRecep As String, CtaContra As String, Concep As String, Amplia As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariagro.rfactsoc --> conta.cabfactprov
' ariagro.rfactsoc_variedad --> conta.linfactprov
'Actualizar la tabla ariagro.rfactsoc.contabilizada=1 para indicar que ya esta contabilizada
Dim B As Boolean
Dim cadMen As String
Dim Sql As String
Dim Mc As Contadores


    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    Set Mc = New Contadores
    
    ' Insertamos en el diario
    B = InsertarAsientoGastoCampo(cadWHERE, cadMen, Mc, CDate(FechaFin), CDate(FecRecep), CtaContra, Concep, Amplia)
    cadMen = "Insertando Asiento de Gasto en Diario: " & cadMen
    
    If B Then
        '---- Poner contabilizado=1 en rcampos_gastos
        B = ActualizarCabFactSoc("rcampos_gastos", cadWHERE, cadMen)
        cadMen = "Actualizando Concepto Gasto Campo: " & cadMen
    End If
    
EContab:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Contabilizando Asiento Gasto de Campo", Err.Description
    End If
    If B Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarAsientoGastoCampo = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarAsientoGastoCampo = False
        If Not B Then
            InsertarTMPErrFacSoc cadMen, cadWHERE
        End If
    End If
End Function



Private Function InsertarAsientoGastoCampo(cadWHERE As String, cadErr As String, ByRef Mc As Contadores, FechaFin As Date, FecRec As Date, CtaContra As String, Concep As String, Amplia As String) As Boolean
' la contabilizacion de las facturas internas de horto, se han de insertar en el diario no en el registro de iva de proveedor
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Concepto As String
Dim letraser As String
Dim Obs As String
Dim cadMen As String
Dim B As Boolean
'Dim CtaSocio As String


    On Error GoTo EInsertar
       
    Sql = " SELECT rcampos_gastos.codgasto, rcampos_gastos.fecha, rcampos_gastos.importe, rconcepgasto.codmacgto "
    Sql = Sql & " FROM (rcampos_gastos "
    Sql = Sql & "INNER JOIN rconcepgasto ON rcampos_gastos.codgasto=rconcepgasto.codgasto) "
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
        If Mc.ConseguirContador("1", (FecRec <= CDate(FechaFin) - 365), True) = 0 Then
        
            Obs = "Contabilización Gasto de Campo de Fecha " & Format(FecRec, "dd/mm/yyyy")
        
            'Insertar en la conta Cabecera Asiento
            B = InsertarCabAsientoDia(vEmpresa.NumDiarioInt, Mc.Contador, CStr(Format(FecRec, "dd/mm/yyyy")), Obs, cadMen)
            cadMen = "Insertando Cab. Asiento: " & cadMen
    
            If B Then
                B = InsertarLinAsientoDiaGastos("rcampos_gastos", cadWHERE, cadMen, CtaContra, Mc.Contador, Concep, Amplia)
                cadMen = "Insertando Lin. Asiento Diario: " & cadMen
            
                Set Mc = Nothing
            End If
            
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarAsientoGastoCampo = False
        cadErr = Err.Description
    Else
        InsertarAsientoGastoCampo = True
    End If
End Function


Private Function InsertarLinAsientoDiaGastos(cadTabla As String, cadWHERE As String, cadErr As String, CtaContra As String, Contador As Long, Concep As String, Amplia As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim numdocum As String
Dim Ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim I As Long
Dim B As Boolean
Dim cad As String
Dim cadMen As String
Dim FeFact As Date

Dim cadCampo As String

    On Error GoTo eInsertarLinAsientoDiaGastos

    InsertarLinAsientoDiaGastos = False

    Sql = " SELECT rcampos_gastos.fecha, rcampos_gastos.codcampo, rconcepgasto.codmacgto cuenta, rcampos_gastos.importe as importe "
    Sql = Sql & " FROM rcampos_gastos Inner JOIN rconcepgasto ON rcampos_gastos.codgasto = rconcepgasto.codgasto "
    Sql = Sql & " where " & cadWHERE
    Sql = Sql & " order by 1, 2 " '& cadCampo

    
    Set Rs = New ADODB.Recordset
    
    Rs.Open Sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
            
    I = 0
    ImporteD = 0
    ImporteH = 0
    
    numdocum = Format(Rs!codcampo, "00000000")
'    Ampliacion = Format(Rs!codcampo, "00000000")
    ampliaciond = Amplia 'Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoInt, "N")) & " " & Ampliacion
    ampliacionh = Amplia 'Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoInt, "N")) & " " & Ampliacion
    
    If Not Rs.EOF Then Rs.MoveFirst
    
    B = True
    
    If Not Rs.EOF Then
        I = I + 1
        
        FeFact = Rs!Fecha
        
        cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(Rs!Fecha, "F") & "," & DBSet(Contador, "N") & ","
        cad = cad & DBSet(I, "N") & "," & DBSet(Rs!cuenta, "T") & "," & DBSet(numdocum, "T") & ","
        
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If DBLet(Rs!Importe, "N") > 0 Then
            ' importe al debe en positivo
            cad = cad & DBSet(Concep, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Rs!Importe, "N") & ","
            cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaContra, "T") & "," & ValorNulo & ",0"
        
            ImporteD = ImporteD + CCur(DBLet(Rs!Importe, "N"))
        Else
            ' importe al haber en positivo, cambiamos el signo
            cad = cad & DBSet(Concep, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
            cad = cad & DBSet(Rs!Importe, "N") & "," & ValorNulo & "," & DBSet(CtaContra, "T") & "," & ValorNulo & ",0"
        
            ImporteH = ImporteH + CCur(DBLet(Rs!Importe, "N"))
        End If
        
        cad = "(" & cad & ")"
        
        B = InsertarLinAsientoDia(cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & I

        I = I + 1
                
        ' el Total es sobre la cuenta del cliente
        cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(Rs!Fecha, "F") & "," & DBSet(Contador, "N") & ","
        cad = cad & DBSet(I, "N") & ","
        cad = cad & DBSet(CtaContra, "T") & "," & DBSet(numdocum, "T") & ","
            
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If DBLet(Rs!Importe, "N") > 0 Then
            ' importe al haber en positivo, cambiamos el signo
            cad = cad & DBSet(Concep, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
            cad = cad & DBSet(Rs!Importe, "N") & "," & ValorNulo & "," & DBSet(Rs!cuenta, "N") & "," & ValorNulo & ",0"
        Else
            ' importe al debe en positivo
            cad = cad & DBSet(Concep, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(DBLet(Rs!Importe, "N") * (-1), "N") & ","
            cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!cuenta, "N") & "," & ValorNulo & ",0"
        
        End If
        
        cad = "(" & cad & ")"
        
        B = InsertarLinAsientoDia(cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & I

    End If
    
    Set Rs = Nothing
    InsertarLinAsientoDiaGastos = B
    Exit Function
    
eInsertarLinAsientoDiaGastos:
    cadErr = "Insertar Linea Asiento Gastos: " & Err.Description
    cadErr = cadErr & cadMen
    InsertarLinAsientoDiaGastos = False
End Function


'----------------------------------------------------------------------
' FACTURAS VARIAS REGISTRO CLIENTE
'----------------------------------------------------------------------
Public Function PasarFacturaFVAR(cadWHERE As String, CodCCost As String, FechaFin As String, Seccion As String, TipoFact As Byte, FecVto As Date, ForpaPos As String, ForpaNeg As String, CtaBanc As String, TipoM As String, ByRef vContaFra As cContabilizarFacturas, Optional FecRecep As Date) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariagro.rfactsoc --> conta.cabfactprov
' ariagro.rfactsoc_variedad --> conta.linfactprov
'Actualizar la tabla ariagro.rfactsoc.contabilizada=1 para indicar que ya esta contabilizada
Dim B As Boolean
Dim cadMen As String
Dim Sql As String
Dim Mc As Contadores
Dim Obs As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
        
    Set Mc = New Contadores
        
    CtaSocio = ""
    FacturaSoc = ""
    
    BancoSoc = 0
    SucurSoc = 0
    DigcoSoc = ""
    CtaBaSoc = ""
    IbanSoc = ""
    
    ImpReten = 0
    CtaReten = ""
        
    If TipoM = "FVG" Then
        B = True
        ' tendriamos que insertar en el diario FALTA
    Else
        If TipoM = "FVP" Then 'registro de iva de proveedor
            B = InsertarCabFactFVARPro(cadWHERE, cadMen, Mc, CDate(FechaFin), Seccion, CStr(FecRecep), vContaFra)
            cadMen = "Insertando Cab. Factura Proveedor: " & cadMen
        Else ' registro de iva de cliente
            '---- Insertar en la conta Cabecera Factura
            B = InsertarCabFactFVAR(cadWHERE, cadMen, TipoFact, Seccion, vContaFra)
            cadMen = "Insertando Cab. Factura: " & cadMen
        End If
    End If
    
    If B Then
        FecVenci = FecVto
        ForpaPosi = ForpaPos
        ForpaNega = ForpaNeg
        CtaBanco = CtaBanc
        tipoMov = TipoM    ' codtipom de la factura de socio
        
        If TipoM = "FVP" Then ' registro de iva de proveedor
            B = InsertarEnTesoreriaNewFVARPro(cadWHERE, cadMen, CtaBanco, CStr(FecVenci))
            cadMen = "Insertando en Tesoreria: " & cadMen
        Else
            'si la factura es a un cliente o de socio a no descontar en liquidacion , se inserta en tesoreria
            If TipoFact = 1 Or (TipoFact = 0 And Not FraADescontarEnLiquidacion(cadWHERE)) Then
                B = InsertarEnTesoreriaNewFVAR(cadWHERE, CtaBanco, CStr(FecVenci), cadMen, TipoFact, Seccion)
                cadMen = "Insertando en Tesoreria: " & cadMen
            End If
        End If
        If B Then
            If TipoM = "FVP" Then ' registro de iva de proveedores
                CCoste = CodCCost
                '---- Insertar lineas de Factura en la Conta
                If vParamAplic.ContabilidadNueva Then
                    B = InsertarLinFactFVARContaNueva("fvarcabfactpro", cadWHERE, cadMen, CStr(FecRecep), Mc.Contador)
                Else
                    B = InsertarLinFactFVAR("fvarcabfactpro", cadWHERE, cadMen, Mc.Contador)
                End If
                cadMen = "Insertando Lin. Factura: " & cadMen
            Else
                If TipoM <> "FVG" Then
                    CCoste = CodCCost
                    '---- Insertar lineas de Factura en la Conta
                    If vParamAplic.ContabilidadNueva Then
                        B = InsertarLinFactFVARContaNueva("fvarcabfact", cadWHERE, cadMen)
                    Else
                        B = InsertarLinFactFVAR("fvarcabfact", cadWHERE, cadMen)
                    End If
                    cadMen = "Insertando Lin. Factura: " & cadMen
                End If
            End If
            
            If B Then
                If TipoM = "FVP" Then ' registro de iva de proveedor
                    If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaProv(vContaFra.NumeroFactura, vContaFra.Anofac)
                Else
                    If vParamAplic.ContabilidadNueva Then vContaFra.AnyadeElError vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
                End If
            End If
            
            
            
            If B Then
                '---- Poner intconta=1 en ariges.scafac
                If TipoM = "FVP" Then ' registro de iva de proveedores
                    B = ActualizarCabFact("fvarcabfactpro", cadWHERE, cadMen)
                Else
                    B = ActualizarCabFact("fvarcabfact", cadWHERE, cadMen)
                End If
                cadMen = "Actualizando Factura Varia: " & cadMen
            End If
        End If
    End If
    
    
EContab:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Contabilizando Facturas Varias", Err.Description
    End If
    If B Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaFVAR = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaFVAR = False
        If Not B Then
            InsertarTMPErrFacFVAR cadMen, cadWHERE
        End If
    End If
End Function


Private Function InsertarCabFactFVAR(cadWHERE As String, cadErr As String, Tipo As Byte, vSeccion As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Seccion As Integer

Dim IvaImp As Currency
Dim Sql2 As String
Dim CadenaInsertFaclin2 As String


    On Error GoTo EInsertar
    
    ' factura de cliente (socio)
    If Tipo = 0 Then
        Sql = "SELECT stipom.letraser,numfactu,fecfactu, rsocios_seccion.codmaccli as codmacta,year(fecfactu) as anofaccl,"
        Sql = Sql & "baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
        Sql = Sql & "totalfac,tipoiva1,tipoiva2,tipoiva3, porcrec1, porcrec2, porcrec3, imporec1, imporec2, imporec3, "
        Sql = Sql & "retfaccl, trefaccl, cuereten, codforpa, "
        Sql = Sql & "rsocios.nomsocio, rsocios.dirsocio,rsocios.pobsocio,rsocios.codpostal,rsocios.prosocio,nifsocio"
        Sql = Sql & " FROM ((" & "fvarcabfact inner join " & "usuarios.stipom on fvarcabfact.codtipom=usuarios.stipom.codtipom) "
        Sql = Sql & "INNER JOIN rsocios ON fvarcabfact.codsocio=rsocios.codsocio) "
        Sql = Sql & "INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vSeccion, "N")
        Sql = Sql & " WHERE " & cadWHERE
    Else
    ' factura de cliente (cliente)
        Sql = "SELECT stipom.letraser,numfactu,fecfactu, clientes.codmacta as codmacta,year(fecfactu) as anofaccl,"
        Sql = Sql & "baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
        Sql = Sql & "totalfac,tipoiva1,tipoiva2,tipoiva3, porcrec1, porcrec2, porcrec3, imporec1, imporec2, imporec3, "
        Sql = Sql & "retfaccl, trefaccl, cuereten, fvarcabfact.codforpa, "
        Sql = Sql & "clientes.nomclien nomsocio,clientes.domclien dirsocio,clientes.pobclien pobsocio,clientes.codpobla codpostal,clientes.proclien prosocio,clientes.cifclien nifsocio "
        Sql = Sql & " FROM ((" & "fvarcabfact inner join " & "usuarios.stipom on fvarcabfact.codtipom=usuarios.stipom.codtipom) "
        Sql = Sql & "INNER JOIN clientes ON fvarcabfact.codclien=clientes.codclien) "
        Sql = Sql & " WHERE " & cadWHERE
    End If
        
        
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
        vContaFra.NumeroFactura = DBLet(Rs!numfactu)
        vContaFra.Anofac = Year(Rs!fecfactu)
        vContaFra.Serie = DBLet(Rs!letraser)
        
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = 0
        DtoGnral = 0
        BaseImp = Rs!BaseIVA1 + CCur(DBLet(Rs!BaseIVA2, "N")) + CCur(DBLet(Rs!BaseIVA3, "N"))
        IvaImp = DBLet(Rs!impoiva1, "N") + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
        '---- Laura 10/10/2006:  añadir el totalfac para utilizarlo en insertar lineas
        TotalFac = Rs!TotalFac
        '----
        
        Sql = ""
        Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!fecfactu) & "," & ValorNulo & ","
        
        If vParamAplic.ContabilidadNueva Then
            ' para el caso de las rectificativas
            Dim vTipM As String
            vTipM = DevuelveValor("select codtipom from stipom where letraser = " & DBSet(Rs!letraser, "T"))
            If vTipM = "FAR" Then
                Sql = Sql & "'D',"
            Else
                Sql = Sql & "'0',"
            End If
            
            Sql = Sql & "0," & DBSet(Rs!Codforpa, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(IvaImp, "N") & ","
            Sql = Sql & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & ","
            Sql = Sql & DBSet(Rs!retfaccl, "N", "S") & "," & DBSet(Rs!trefaccl, "N", "S") & "," & DBSet(Rs!cuereten, "T", "S") & ","
            If DBLet(Rs!retfaccl, "N") = 0 Then
                Sql = Sql & "0,"
            Else
                Sql = Sql & "2,"
            End If
            Sql = Sql & DBSet(Rs!nomsocio, "T") & "," & DBSet(Rs!dirsocio, "T") & "," & DBSet(Rs!pobsocio, "T") & "," & DBSet(Rs!codpostal, "T") & ","
            Sql = Sql & DBSet(Rs!prosocio, "T") & "," & DBSet(Rs!nifSocio, "T") & ",'ES',1"
            
            Sql = "(" & Sql & ")"
            
            Sql2 = "INSERT INTO factcli (numserie,numfactu,fecfactu,fecliqcl,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
            Sql2 = Sql2 & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,nommacta,dirdatos,despobla,codpobla,desprovi,nifdatos,"
            Sql2 = Sql2 & "codpais,codagente)"
            Sql2 = Sql2 & " VALUES " & cad
            ConnConta.Execute Sql2 & Sql
    '***
            CadenaInsertFaclin2 = ""
                
            
            'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
            'IVA 1, siempre existe
            Sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
            Sql2 = Sql2 & "1," & DBSet(Rs!BaseIVA1, "N") & "," & Rs!TipoIVA1 & "," & DBSet(Rs!porciva1, "N") & ","
            Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
            
            'para las lineas
            vTipoIva(0) = Rs!TipoIVA1
            vPorcIva(0) = Rs!porciva1
            vPorcRec(0) = 0
            vImpIva(0) = Rs!impoiva1
            vImpRec(0) = 0
            vBaseIva(0) = Rs!BaseIVA1
            
            vTipoIva(1) = 0: vTipoIva(2) = 0
            
            If Not IsNull(Rs!porciva2) Then
                Sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
                Sql2 = Sql2 & "2," & DBSet(Rs!BaseIVA2, "N") & "," & Rs!TipoIVA2 & "," & DBSet(Rs!porciva2, "N") & ","
                Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva2, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
                vTipoIva(1) = Rs!TipoIVA2
                vPorcIva(1) = Rs!porciva2
                vPorcRec(1) = 0
                vImpIva(1) = Rs!impoiva2
                vImpRec(1) = 0
                vBaseIva(1) = Rs!BaseIVA2
            End If
            If Not IsNull(Rs!porciva3) Then
                Sql2 = "'" & Rs!letraser & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & Year(Rs!fecfactu) & ","
                Sql2 = Sql2 & "3," & DBSet(Rs!BaseIVA3, "N") & "," & Rs!TipoIVA3 & "," & DBSet(Rs!porciva3, "N") & ","
                Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva3, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
                vTipoIva(2) = Rs!TipoIVA3
                vPorcIva(2) = Rs!porciva3
                vPorcRec(2) = 0
                vImpIva(2) = Rs!impoiva3
                vImpRec(2) = 0
                vBaseIva(2) = Rs!BaseIVA3
            End If
    
            Sql = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
            Sql = Sql & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
            ConnConta.Execute Sql
        
        Else
            Sql = Sql & DBSet(Rs!BaseIVA1, "N") & "," & DBSet(Rs!BaseIVA2, "N", "S") & "," & DBSet(Rs!BaseIVA3, "N", "S") & "," & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", "S") & "," & DBSet(Rs!porciva3, "N", "S") & ","
            Sql = Sql & DBSet(Rs!porcrec1, "N", "S") & "," & DBSet(Rs!porcrec2, "N", "S") & "," & DBSet(Rs!porcrec3, "N", "S") & "," & DBSet(Rs!impoiva1, "N", "N") & "," & DBSet(Rs!impoiva2, "N", "S") & "," & DBSet(Rs!impoiva3, "N", "S") & ","
            Sql = Sql & DBSet(Rs!imporec1, "N", "S") & "," & DBSet(Rs!imporec2, "N", "S") & "," & DBSet(Rs!imporec3, "N", "S") & ","
            Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!TipoIVA1, "N") & "," & DBSet(Rs!TipoIVA2, "N", "S") & "," & DBSet(Rs!TipoIVA3, "N", "S") & ",0,"
            Sql = Sql & DBSet(Rs!retfaccl, "N", "S") & "," & DBSet(Rs!trefaccl, "N", "S") & "," & DBSet(Rs!cuereten, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
        
            cad = cad & "(" & Sql & ")"
        
            'Insertar en la contabilidad
            Sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,fecliqcl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
            Sql = Sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
            Sql = Sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien) "
            Sql = Sql & " VALUES " & cad
            ConnConta.Execute Sql
        
        End If
'        RS.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
        
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactFVAR = False
        cadErr = Err.Description
    Else
        InsertarCabFactFVAR = True
    End If
End Function



Public Function InsertarEnTesoreriaNewFVAR(cadWHERE As String, CtaBan As String, FecVen As String, MenError As String, Tipo As Byte, vSeccion As String) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
Dim B As Boolean
Dim Sql As String, Text33csb As String, Text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset
Dim rsVenci As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Long
Dim DigConta As String
Dim CC As String

Dim Iban As String
Dim CodBanco As String
Dim CodSucur As String
Dim CuentaBa As String
Dim Codmacta As String



Dim vrefer As String
Dim letraser As String
Dim Rsx As ADODB.Recordset
Dim FecVenci As Date
Dim ImpVenci As Currency
Dim ImpVenci1 As Currency
Dim AcumIva As Currency
Dim PorcIva As String

Dim Rsx7 As ADODB.Recordset
Dim Sql7 As String
Dim cadena As String

Dim CadRegistro As String
Dim CadRegistro1 As String

Dim vSocio As cSocio

    On Error GoTo EInsertarTesoreriaNewFac

    B = False
    InsertarEnTesoreriaNewFVAR = False
    
    CadValues = ""
    CadValues2 = ""
    
    Set Rsx = New ADODB.Recordset
    Sql = "select * from fvarcabfact where " & cadWHERE
    Rsx.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    If Not Rsx.EOF Then
        letraser = ""
        letraser = ObtenerLetraSerie2(Rsx!CodTipom)
    
        If Tipo = 0 Then
            ' socio
            
            Dim vSoc As cSocio
            Set vSoc = New cSocio
            
            
            If vSoc.LeerDatos(DBLet(Rsx!Codsocio, "N")) Then
                If vSoc.LeerDatosSeccion(DBLet(Rsx!Codsocio, "N"), CStr(vSeccion)) Then
                    B = True
                            
                    CC = DBLet(vSoc.Digcontrol, "T")
                    If DBLet(vSoc.Digcontrol, "T") = "**" Then CC = "00"
        
                    Iban = vSoc.Iban
                    CodBanco = vSoc.Banco
                    CodSucur = vSoc.Sucursal
                    CuentaBa = vSoc.CuentaBan
                    Codmacta = vSoc.CtaClien
                End If
            End If
    
        Else
            ' cliente
            Sql4 = "select codbanco, codsucur, digcontr, cuentaba, codmacta, iban, nomclien,domclien,pobclien,codpobla,proclien,cifclien  from clientes where codclien = " & DBLet(Rsx!CodClien, "N")
            Set Rs4 = New ADODB.Recordset
            
            Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not Rs4.EOF Then
                B = True
                
                CC = DBLet(Rs4!digcontr, "T")
                If DBLet(Rs4!digcontr, "T") = "**" Then CC = "00"
                
                Iban = DBLet(Rs4!Iban, "T")
                CodBanco = DBLet(Rs4!CodBanco, "N")
                CodSucur = DBLet(Rs4!CodSucur, "N")
                CuentaBa = DBLet(Rs4!CuentaBa, "T")
                Codmacta = DBLet(Rs4!Codmacta, "T")
            End If
        End If
            
        If B Then
            Text33csb = "'Factura:" & DBLet(letraser, "T") & "-" & DBLet(Rsx!numfactu, "T") & " " & Format(DBLet(Rsx!fecfactu, "F"), "dd/mm/yy") & "'"
            Text41csb = "de " & DBSet(Rsx!TotalFac, "N")
            
            'Obtener el Nº de Vencimientos de la forma de pago
            Sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & DBSet(Rsx!Codforpa, "N")
            Set rsVenci = New ADODB.Recordset
            rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

            If Not rsVenci.EOF Then
                If DBLet(rsVenci!numerove, "N") > 0 Then
            
                    CadValuesAux2 = "('" & Trim(letraser) & "', " & DBSet(Rsx!numfactu, "N") & ", " & DBSet(Rsx!fecfactu, "F") & ", "
                    '-------- Primer Vencimiento
                    I = 1
                    'FECHA VTO
                    FecVenci = DBLet(Rsx!fecfactu, "F")
                    '=== Laura 23/01/2007
                    'FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                    FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                    '===
                    
                    CadValues2 = CadValuesAux2 & I & ", "
                    
                    '[Monica]03/07/2013: añado trim(codmacta)
                    CadValues2 = CadValues2 & DBSet(Trim(Codmacta), "T") & ", " & DBSet(Rsx!Codforpa, "N") & ", " & DBSet(FecVenci, "F") & ", "
                    
                    'IMPORTE del Vencimiento
                    If rsVenci!numerove = 1 Then
                        ImpVenci = DBLet(Rsx!TotalFac, "N")
                    Else
                        ImpVenci = Round2(DBLet(Rsx!TotalFac, "N") / rsVenci!numerove, 2)
                        'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                        If ImpVenci * rsVenci!numerove <> DBLet(Rsx!TotalFac, "N") Then
                            ImpVenci = Round2(ImpVenci + (DBLet(Rsx!TotalFac, "N") - ImpVenci * rsVenci!numerove), 2)
                        End If
                    End If
                    
                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", "
                    
                    If Not vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & DBSet(CodBanco, "N", "S") & ", " & DBSet(CodSucur, "N", "S") & ", " & DBSet(CC, "T", "S") & ", " & DBSet(CuentaBa, "T", "S") & ", "
                        CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & Text33csb & "," & DBSet(Text41csb, "T") & ",1" '),"
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & ", " & DBSet(Iban, "T", "S") & "),"
                        Else
                            CadValues2 = CadValues2 & "),"
                        End If
                    Else
                        CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & Text33csb & "," & DBSet(Text41csb, "T") & ",1"
                        
                        vvIban = MiFormat(Iban, "") & MiFormat(CodBanco, "0000") & MiFormat(CodSucur, "0000") & MiFormat(CC, "00") & MiFormat(CuentaBa, "0000000000")
                        
                        CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                        
                        If Tipo = 0 Then ' socio
                            CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & ","
                            CadValues2 = CadValues2 & DBSet(vSoc.CPostal, "T") & "," & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'),"
                        Else ' cliente
                            'nomclien,domclien,pobclien,codpobla,proclien,cifclien
                            CadValues2 = CadValues2 & DBSet(Rs4!nomclien, "T") & "," & DBSet(Rs4!domclien, "T") & "," & DBSet(Rs4!pobclien, "T") & ","
                            CadValues2 = CadValues2 & DBSet(Rs4!CodPobla, "T") & "," & DBSet(Rs4!proclien, "T") & "," & DBSet(Rs4!cifclien, "T") & ",'ES'),"
                        End If
                    End If
                    
                
                    'Resto Vencimientos
                    '--------------------------------------------------------------------
                    For I = 2 To rsVenci!numerove
                       'FECHA Resto Vencimientos
                        '=== Laura 23/01/2007
                        'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                        FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                        '===
                            
                        CadValues2 = CadValues2 & CadValuesAux2 & I & ", " & DBSet(Trim(Rs4!Codmacta), "T") & ", " & DBSet(Rsx!Codforpa, "N") & ", '" & Format(FecVenci, FormatoFecha) & "', "
                        
                        'IMPORTE Resto de Vendimientos
                        ImpVenci = Round2(TotalFac / rsVenci!numerove, 2)
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", "
                        
                        If Not vParamAplic.ContabilidadNueva Then
                            CadValues2 = CadValues2 & DBSet(Rs4!CodBanco, "N", "S") & ", " & DBSet(Rs4!CodSucur, "N", "S") & ", " & DBSet(CC, "T", "S") & ", " & DBSet(Rs4!CuentaBa, "T", "S") & ", "
                            CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & Text33csb & "," & DBSet(Text41csb, "T") & ",1" '),"
                            '[Monica]22/11/2013: Tema iban
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                CadValues2 = CadValues2 & ", " & DBSet(Iban, "T", "S") & "),"
                            Else
                                CadValues2 = CadValues2 & "),"
                            End If
                        Else
                            CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & Text33csb & "," & DBSet(Text41csb, "T") & ",1"
                            
                            vvIban = MiFormat(Iban, "") & MiFormat(DBLet(Rs4!CodBanco), "0000") & MiFormat(DBLet(Rs4!CodSucur), "0000") & MiFormat(CC, "00") & MiFormat(DBLet(Rs4!CuentaBa), "0000000000")
                            
                            CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                            
                            If Tipo = 0 Then ' socio
                                CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & ","
                                CadValues2 = CadValues2 & DBSet(vSoc.CPostal, "T") & "," & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'),"
                            Else ' cliente
                                'nomclien,domclien,pobclien,codpobla,proclien,cifclien
                                CadValues2 = CadValues2 & DBSet(Rs4!nomclien, "T") & "," & DBSet(Rs4!domclien, "T") & "," & DBSet(Rs4!pobclien, "T") & ","
                                CadValues2 = CadValues2 & DBSet(Rs4!CodPobla, "T") & "," & DBSet(Rs4!proclien, "T") & "," & DBSet(Rs4!cifclien, "T") & ",'ES'),"
                            End If
                        End If
                    Next I
                    ' quitamos la ultima coma
                    CadValues2 = Mid(CadValues2, 1, Len(CadValues2) - 1)
                        
                    If vParamAplic.ContabilidadNueva Then
                        Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                        Sql = Sql & "ctabanc1,  fecultco, impcobro, "
                        Sql = Sql & " text33csb, text41csb, agente, iban, " ') "
                        Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
                        Sql = Sql & ") "
                    
                    Else
                        'Insertamos en la tabla scobro de la CONTA
                        Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                        Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                        Sql = Sql & " text33csb, text41csb, agente" ') "
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            Sql = Sql & ", iban) "
                        Else
                            Sql = Sql & ") "
                        End If
                    End If
                    Sql = Sql & " VALUES " & CadValues2
                    ConnConta.Execute Sql
                
                End If
            End If
        
            B = True

        End If
    
    End If
    
EInsertarTesoreriaNewFac:
    If Err.Number <> 0 Then
        B = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaNewFVAR = B
End Function





Private Function InsertarLinFactFVAR(cadTabla As String, cadWHERE As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String

    On Error GoTo EInLinea
    
    If cadTabla = "fvarcabfact" Then
        cadCampo = "fvarconce.codmacta"
    Else
        cadCampo = "fvarconce.codmacpr"
    End If
    
    If cadTabla = "fvarcabfact" Then
        If vEmpresa.TieneAnalitica Then
            Sql = " SELECT usuarios.stipom.letraser,fvarlinfact.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importe) as importe, fvarconce.codccost "
        Else
            Sql = " SELECT usuarios.stipom.letraser,fvarlinfact.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importe) as importe "
        End If
        
        Sql = Sql & " FROM (fvarlinfact inner join usuarios.stipom on fvarlinfact.codtipom=usuarios.stipom.codtipom) "
        Sql = Sql & " inner join fvarconce on fvarlinfact.codconce=fvarconce.codconce "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "fvarcabfact", "fvarlinfact")
    Else
        If vEmpresa.TieneAnalitica Then
            Sql = " SELECT usuarios.stipom.letraser,fvarlinfactpro.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importe) as importe, fvarconce.codccost "
        Else
            Sql = " SELECT usuarios.stipom.letraser,fvarlinfactpro.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importe) as importe "
        End If
        
        Sql = Sql & " FROM (fvarlinfactpro inner join usuarios.stipom on fvarlinfactpro.codtipom=usuarios.stipom.codtipom) "
        Sql = Sql & " inner join fvarconce on fvarlinfactpro.codconce=fvarconce.codconce "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "fvarcabfactpro", "fvarlinfactpro")
    End If
    
    If vEmpresa.TieneAnalitica Then
        Sql = Sql & " GROUP BY 5,7 " '& cadCampo, codccost
    Else
        Sql = Sql & " GROUP BY 5 " '& cadCampo
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
' --monica:no hay descuentos
'        '---- Laura: 10/10/2006
'        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
'        ImpLinea = RS!Importe - CCur(CalcularPorcentaje(RS!Importe, DtoPPago, 2))
'        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
'        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(RS!Importe, DtoGnral, 2))
'        'ImpLinea = Round(ImpLinea, 2)
'        '----
'        TotImp = TotImp + ImpLinea
        ImpLinea = DBLet(Rs!Importe, "N")
        totimp = totimp + DBLet(Rs!Importe, "N")

        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        If cadTabla = "fvarcabfact" Then
            Sql = "'" & Trim(Rs!letraser) & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
            Sql = Sql & DBSet(Trim(Rs!cuenta), "T")
            
        Else
            Sql = NumRegis & "," & Year(Rs!fecfactu) & "," & I & ","
            Sql = Sql & DBSet(Trim(Rs!cuenta), "T")
        
        End If
        
        Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        
        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & DBSet(Rs!CodCCost, "T")
            CCoste = DBSet(Rs!CodCCost, "T")
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        cad = cad & "(" & Sql & ")" & ","
        
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)
        Sql2 = Sql2 & DBSet(totimp, "N") & ","
        If CCoste = "" Or CCoste = ValorNulo Then
            Sql2 = Sql2 & ValorNulo
        Else
            Sql2 = Sql2 & DBSet(CCoste, "T")
        End If
        If SQLaux <> "" Then 'hay mas de una linea
            cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If

    If cadTabla = "fvarcabfactpro" Then
        ' las retenciones si las hay
        If ImpReten <> 0 Then
            Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
            Sql = Sql & DBSet(Trim(CtaSocio), "T")
            Sql = Sql & "," & DBSet(ImpReten, "N") & ","
            Sql = Sql & ValorNulo ' no llevan centro de coste
            
            cad = cad & "(" & Sql & ")" & ","
            I = I + 1
            
            Sql = NumRegis & "," & AnyoFacPr & "," & I & ","
            Sql = Sql & DBSet(Trim(CtaReten), "T")
            Sql = Sql & "," & DBSet(ImpReten * (-1), "N") & ","
            Sql = Sql & ValorNulo ' no llevan centro de coste
            
            cad = cad & "(" & Sql & ")" & ","
            I = I + 1
        End If
    End If


    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        If cadTabla = "fvarcabfact" Then
            Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Else
            Sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        End If
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactFVAR = False
        cadErr = Err.Description
    Else
        InsertarLinFactFVAR = True
    End If
End Function

Private Function InsertarLinFactFVARContaNueva(cadTabla As String, cadWHERE As String, cadErr As String, Optional FecRecep As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String

Dim NumeroIVA As Byte
Dim k As Integer
Dim HayQueAjustar As Boolean

Dim ImpImva As Currency
Dim ImpREC As Currency


    On Error GoTo EInLinea
    
    If cadTabla = "fvarcabfact" Then
        cadCampo = "fvarconce.codmacta"
    Else
        cadCampo = "fvarconce.codmacpr"
    End If
    
    If cadTabla = "fvarcabfact" Then
        If vEmpresa.TieneAnalitica Then
            Sql = " SELECT usuarios.stipom.letraser,fvarlinfact.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,fvarlinfact.tipoiva, tiposiva.porceiva porciva,  tiposiva.porcerec porcrec,sum(importe) as importe, fvarconce.codccost "
        Else
            Sql = " SELECT usuarios.stipom.letraser,fvarlinfact.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,fvarlinfact.tipoiva, tiposiva.porceiva porciva,  tiposiva.porcerec porcrec,sum(importe) as importe "
        End If
        
        Sql = Sql & " FROM ((fvarlinfact inner join usuarios.stipom on fvarlinfact.codtipom=usuarios.stipom.codtipom) "
        Sql = Sql & " inner join fvarconce on fvarlinfact.codconce=fvarconce.codconce) "
        Sql = Sql & " inner join ariconta" & vParamAplic.NumeroConta & ".tiposiva on ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = fvarlinfact.tipoiva "
        
        Sql = Sql & " WHERE " & Replace(cadWHERE, "fvarcabfact", "fvarlinfact")
    Else
        If vEmpresa.TieneAnalitica Then
            Sql = " SELECT usuarios.stipom.letraser,fvarlinfactpro.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,fvarlinfactpro.tipoiva, tiposiva.porceiva porciva,  tiposiva.porcerec porcrec,sum(importe) as importe, fvarconce.codccost "
        Else
            Sql = " SELECT usuarios.stipom.letraser,fvarlinfactpro.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,fvarlinfactpro.tipoiva, tiposiva.porceiva porciva,  tiposiva.porcerec porcrec,sum(importe) as importe "
        End If
        
        Sql = Sql & " FROM ((fvarlinfactpro inner join usuarios.stipom on fvarlinfactpro.codtipom=usuarios.stipom.codtipom) "
        Sql = Sql & " inner join fvarconce on fvarlinfactpro.codconce=fvarconce.codconce) "
        Sql = Sql & " inner join ariconta" & vParamAplic.NumeroConta & ".tiposiva on ariconta" & vParamAplic.NumeroConta & ".tiposiva.codigiva = fvarlinfactpro.tipoiva "
        
        Sql = Sql & " WHERE " & Replace(cadWHERE, "fvarcabfactpro", "fvarlinfactpro")
    End If
    
    If vEmpresa.TieneAnalitica Then
        Sql = Sql & " GROUP BY 5,6,7,8,10 " '& cadCampo, codccost
    Else
        Sql = Sql & " GROUP BY 5,6,7,8 " '& cadCampo
    End If
    Sql = Sql & " ORDER BY 6,5 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenKeyset, adLockPessimistic, adCmdText

    cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
' --monica:no hay descuentos
'        '---- Laura: 10/10/2006
'        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
'        ImpLinea = RS!Importe - CCur(CalcularPorcentaje(RS!Importe, DtoPPago, 2))
'        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
'        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(RS!Importe, DtoGnral, 2))
'        'ImpLinea = Round(ImpLinea, 2)
'        '----
'        TotImp = TotImp + ImpLinea
        ImpLinea = DBLet(Rs!Importe, "N")
        totimp = totimp + DBLet(Rs!Importe, "N")

        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        If cadTabla = "fvarcabfact" Then
            Sql = "'" & Trim(Rs!letraser) & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
            Sql = Sql & DBSet(Trim(Rs!cuenta), "T")
            
        Else
            Sql = DBSet(SerieFraPro, "T") & "," & NumRegis & "," & DBSet(FecRecep, "F") & "," & Year(Rs!fecfactu) & "," & I & ","
            Sql = Sql & DBSet(Trim(Rs!cuenta), "T")
        
        End If
        
        'Vemos que tipo de IVA es en el vector de importes
        NumeroIVA = 127
        For k = 0 To 2
            If Rs!TipoIVA = vTipoIva(k) Then
                NumeroIVA = k
                Exit For
            End If
        Next
        If NumeroIVA > 100 Then Err.Raise 513, "Error obteniendo IVA: " & Rs!Codigiva
        
        
        
        Sql = Sql & ","
        
        If vEmpresa.TieneAnalitica Then
            Sql = Sql & DBSet(Rs!CodCCost, "T")
            CCoste = DBSet(Rs!CodCCost, "T")
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        If cadTabla = "fvarcabfact" Then Sql = Sql & "," & DBSet(Rs!fecfactu, "F")
        
        vBaseIva(NumeroIVA) = vBaseIva(NumeroIVA) - ImpLinea   'Para ajustar el importe y que no haya descuadre
        
        'Caluclo el importe de IVA y el de recargo de equivalencia
        ImpImva = vPorcIva(NumeroIVA) / 100
        ImpImva = Round2(ImpLinea * ImpImva, 2)
        If vPorcRec(NumeroIVA) = 0 Then
            ImpREC = 0
        Else
            ImpREC = vPorcRec(NumeroIVA) / 100
            ImpREC = Round2(ImpLinea * ImpREC, 2)
        End If
        vImpIva(NumeroIVA) = vImpIva(NumeroIVA) - ImpImva
        vImpRec(NumeroIVA) = vImpRec(NumeroIVA) - ImpREC
        
        
        HayQueAjustar = False
        If vBaseIva(NumeroIVA) <> 0 Or vImpIva(NumeroIVA) <> 0 Or vImpRec(NumeroIVA) <> 0 Then
            'falta importe.
            'Puede ser que hayan mas lineas, o haya descuadre. Como esta ordenado por tipo de iva
            Rs.MoveNext
            If Rs.EOF Then
                'No hay mas lineas
                'Hay que ajustar SI o SI
                HayQueAjustar = True
            Else
                'Si que hay mas lineas.
                'Son del mismo tipo de IVA
                If Rs!TipoIVA <> vTipoIva(NumeroIVA) Then
                    'NO es el mismo tipo de IVA
                    'Hay que ajustar
                    HayQueAjustar = True
                End If
            End If
            Rs.MovePrevious
        End If
        
        Sql = Sql & "," & vTipoIva(NumeroIVA) & "," & DBSet(vPorcIva(NumeroIVA), "N") & "," & DBSet(vPorcRec(NumeroIVA), "N", "S") & ","
        
        If HayQueAjustar Then
            If vBaseIva(NumeroIVA) <> 0 Then ImpLinea = ImpLinea + vBaseIva(NumeroIVA)
            If vImpIva(NumeroIVA) <> 0 Then ImpImva = ImpImva + vImpIva(NumeroIVA)
            If vImpRec(NumeroIVA) <> 0 Then ImpREC = ImpREC + vImpRec(NumeroIVA)
        End If

        
        ' baseimpo , impoiva, imporec
        Sql = Sql & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpREC, "N", "S") & ",1"
        
        
        cad = cad & "(" & Sql & ")" & ","
        
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        If cadTabla = "fvarcabfact" Then
            Sql = "INSERT INTO factcli_lineas (numserie,numfactu,anofactu,numlinea,codmacta,codccost,fecfactu,codigiva,porciva,porcrec,baseimpo,impoiva,imporec,aplicret) "
        Else
            Sql = "INSERT INTO factpro_lineas (numserie,numregis,fecharec,anofactu,numlinea,codmacta,codccost,codigiva,porciva,porcrec,baseimpo,impoiva,imporec,aplicret) "
        End If
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactFVARContaNueva = False
        cadErr = Err.Description
    Else
        InsertarLinFactFVARContaNueva = True
    End If
End Function






Private Function FraADescontarEnLiquidacion(cWhere As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "select enliquidacion from fvarcabfact where " & cWhere
    
    FraADescontarEnLiquidacion = (DevuelveValor(Sql) > 0)

End Function




Private Function InsertarCabFactFVARPro(cadWHERE As String, cadErr As String, ByRef Mc As Contadores, FechaFin As String, Seccion As String, FecRecep As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim TipoOpera As Integer
Dim Aux As String

Dim Sql2 As String
Dim CadenaInsertFaclin2 As String
Dim ImporAux As Currency
Dim ImporAux2 As Currency

    On Error GoTo EInsertar
       
    
    Sql = " SELECT fecfactu," & Year(CDate(FecRecep)) & " as anofacpr,numfactu,rsocios_seccion.codmacpro codmacta,"
    Sql = Sql & "baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    Sql = Sql & "totalfac,tipoiva1,tipoiva2,tipoiva3,rsocios_seccion.codsocio, rsocios.nomsocio, fvarcabfactpro.codforpa, "
    Sql = Sql & "retfaccl, trefaccl, cuereten, rsocios.codbanco, rsocios.codsucur, rsocios.digcontr, rsocios.cuentaba, rsocios.iban,  "
    Sql = Sql & " rsocios.dirsocio, rsocios.pobsocio, rsocios.codpostal, rsocios.prosocio, rsocios.nifsocio "
    Sql = Sql & " FROM (fvarcabfactpro "
    Sql = Sql & " INNER JOIN rsocios_seccion ON fvarcabfactpro.codsocio=rsocios_seccion.codsocio) "
    Sql = Sql & " INNER JOIN rsocios ON fvarcabfactpro.codsocio = rsocios.codsocio"
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
    
        If Mc.ConseguirContador("1", (CDate(FecRecep) <= CDate(FechaFin) - 365), True) = 0 Then
        
            vContaFra.NumeroFactura = Mc.Contador
            vContaFra.Anofac = DBLet(Rs!anofacpr)
        
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            DtoPPago = 0
            DtoGnral = 0
            BaseImp = Rs!BaseIVA1 + CCur(DBLet(Rs!BaseIVA2, "N")) + CCur(DBLet(Rs!BaseIVA3, "N"))
            TotalFac = BaseImp + Rs!impoiva1 + CCur(DBLet(Rs!impoiva2, "N")) + CCur(DBLet(Rs!impoiva3, "N"))
            AnyoFacPr = Rs!anofacpr
            
            Nulo2 = "N"
            Nulo3 = "N"
            If DBLet(Rs!BaseIVA2, "N") = "0" Then Nulo2 = "S"
            If DBLet(Rs!BaseIVA3, "N") = "0" Then Nulo3 = "S"
            Sql = ""
            If vParamAplic.ContabilidadNueva Then Sql = "'" & SerieFraPro & "',"
            Sql = Sql & Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & Rs!anofacpr & "," & DBSet(FecRecep, "F") & "," & DBSet(FecRecep, "F") & "," & DBSet(Rs!numfactu, "T") & "," & DBSet(Trim(Rs!Codmacta), "T") & "," & ValorNulo & ","
            
            If vParamAplic.ContabilidadNueva Then
                Sql = Sql & DBSet(Rs!nomsocio, "T") & "," & DBSet(Rs!dirsocio, "T", "S") & ","
                Sql = Sql & DBSet(Rs!codpostal, "T", "S") & "," & DBSet(Rs!pobsocio, "T", "S") & "," & DBSet(Rs!prosocio, "T", "S") & ","
                Sql = Sql & DBSet(Rs!nifSocio, "F", "S") & ",'ES',"
                Sql = Sql & DBSet(Rs!Codforpa, "N") & ","
                
                TipoOpera = 0
                
                Aux = "0"
                Select Case TipoOpera
                Case 0
                    If Rs!TotalFac < 0 Then
                        Aux = "D"
                    Else
                        If Not IsNull(Rs!TipoIVA2) Then Aux = "C"
                    End If
                End Select
                
                'codopera,codconce340,codintra
                Sql = Sql & TipoOpera & "," & DBSet(Aux, "T") & "," & ValorNulo & ","
                
                
                'para las lineas
                'factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
                'IVA 1, siempre existe
                Aux = "'" & SerieFraPro & "'," & Mc.Contador & "," & DBSet(FecRecep, "F") & "," & Rs!anofacpr & ","
                
                Sql2 = Aux & "1," & DBSet(Rs!BaseIVA1, "N") & "," & Rs!TipoIVA1 & "," & DBSet(Rs!porciva1, "N") & ","
                Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & ValorNulo
                CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
                vTipoIva(0) = Rs!TipoIVA1
                vPorcIva(0) = Rs!porciva1
                vPorcRec(0) = 0
                vImpIva(0) = Rs!impoiva1
                vImpRec(0) = 0
                vBaseIva(0) = Rs!BaseIVA1
                
                vTipoIva(1) = 0: vTipoIva(2) = 0
                
                If Not IsNull(Rs!porciva2) Then
                    Sql2 = Aux & "2," & DBSet(Rs!BaseIVA2, "N") & "," & Rs!TipoIVA2 & "," & DBSet(Rs!porciva2, "N") & ","
                    Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva2, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
                    vTipoIva(1) = Rs!TipoIVA2
                    vPorcIva(1) = Rs!porciva2
                    vPorcRec(1) = 0
                    vImpIva(1) = Rs!impoiva2
                    vImpRec(1) = 0
                    vBaseIva(1) = Rs!BaseIVA2
                End If
                
                If Not IsNull(Rs!porciva3) Then
                    Sql2 = Aux & "3," & DBSet(Rs!BaseIVA3, "N") & "," & Rs!TipoIVA3 & "," & DBSet(Rs!porciva3, "N") & ","
                    Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva3, "N") & "," & ValorNulo
                    CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
                    vTipoIva(2) = Rs!TipoIVA3
                    vPorcIva(2) = Rs!porciva3
                    vPorcRec(2) = 0
                    vImpIva(2) = Rs!impoiva3
                    vImpRec(2) = 0
                    vBaseIva(2) = Rs!BaseIVA3
                End If
                    
                    
                'Los totales
                'totbases,totbasesret,totivas,totrecargo,totfacpr,
                ImporAux = Rs!BaseIVA1 + DBLet(Rs!BaseIVA2, "N") + DBLet(Rs!BaseIVA3, "N")
                Sql = Sql & DBSet(ImporAux, "N") & ","
                
                If DBLet(Rs!retfaccl) <> 0 Then
                    ImporAux2 = Rs!impoiva1 + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
                    Sql = Sql & DBSet(ImporAux + ImporAux2, "N")
                Else
                    Sql = Sql & ValorNulo
                End If
                Sql = Sql & ","

                        
                'totivas
                ImporAux = Rs!impoiva1 + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
                Sql = Sql & DBSet(ImporAux, "N") & "," & DBSet(Rs!TotalFac, "N") & ","
                
                If DBLet(Rs!retfaccl, "N") <> 0 Then
                    Sql = Sql & DBSet(Rs!retfaccl, "N") & "," & DBSet(Rs!trefaccl, "N") & "," & DBSet(Rs!cuereten, "T") & ",2"
                Else
                    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                End If
                
                cad = cad & "(" & Sql & ")"
            
                Sql = "INSERT INTO factpro(numserie,numregis,fecfactu,anofactu,fecharec,fecliqpr,numfactu,codmacta,observa,nommacta,"
                Sql = Sql & "dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,codforpa,codopera,codconce340,codintra,"
                Sql = Sql & "totbases,totbasesret,totivas,totfacpr,retfacpr,trefacpr,cuereten,tiporeten)"
                Sql = Sql & " VALUES " & cad
                ConnConta.Execute Sql
                
                'Las  lineas de IVA
                Sql = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)"
                Sql = Sql & " VALUES " & CadenaInsertFaclin2
                ConnConta.Execute Sql
                
            Else
                Sql = Sql & DBSet(Rs!BaseIVA1, "N") & "," & DBSet(Rs!BaseIVA2, "N", "S") & "," & DBSet(Rs!BaseIVA3, "N", "S") & ","
                Sql = Sql & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", Nulo2) & "," & DBSet(Rs!porciva3, "N", Nulo3) & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & DBSet(Rs!impoiva2, "N", Nulo2) & "," & DBSet(Rs!impoiva3, "N", Nulo3) & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & DBSet(TotalFac, "N") & "," & DBSet(Rs!TipoIVA1, "N") & "," & DBSet(Rs!TipoIVA2, "N", Nulo2) & "," & DBSet(Rs!TipoIVA3, "N", Nulo3) & ",0,"
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                cad = cad & "(" & Sql & ")"
                
                'Insertar en la contabilidad
                Sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,fecliqpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                Sql = Sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                Sql = Sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,nodeducible) "
                Sql = Sql & " VALUES " & cad
                ConnConta.Execute Sql
            End If
            
            'añadido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            Sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            Sql = Sql & ",'" & DevNombreSQL(Rs!numfactu) & " @ " & Format(Rs!fecfactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!nomsocio) & "'," & Rs!Codsocio & ")"
            conn.Execute Sql
            
            CtaSocio = DBLet(Rs!Codmacta, "T")
            FacturaSoc = DBLet(Rs!numfactu, "N")
            FecFactuSoc = DBLet(Rs!fecfactu)
            
                        
            IbanSoc = DBLet(Rs!Iban, "T")
            BancoSoc = DBLet(Rs!CodBanco, "N")
            SucurSoc = DBLet(Rs!CodSucur, "N")
            DigcoSoc = DBLet(Rs!digcontr, "T")
            CtaBaSoc = DBLet(Rs!CuentaBa, "T")
            
            ImpReten = DBLet(Rs!trefaccl, "N")
            CtaReten = DBLet(Rs!cuereten, "T")
            
            TotalFac = DBLet(Rs!TotalFac, "N")
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactFVARPro = False
        cadErr = Err.Description
    Else
        InsertarCabFactFVARPro = True
    End If
End Function



Private Function InsertarEnTesoreriaNewFVARPro(cadWHERE As String, MenError As String, CtaBanco As String, FecVenci As Date) As Boolean
'Guarda datos de Tesoreria en tablas: spagop o scobro dependiendo del signo de la factura
Dim B As Boolean
Dim Sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim CC As String
Dim letraser As String
Dim Text33csb As String
Dim Text41csb As String
Dim Text42csb As String
Dim GastosPie As Currency
Dim FactuRec As String

Dim Socio As String

    On Error GoTo EInsertarTesoreria

    InsertarEnTesoreriaNewFVARPro = False
    
    
    Dim vSoc As cSocio
    Set vSoc = New cSocio
    
    Socio = DevuelveValor("select codsocio from fvarcabfactpro where " & cadWHERE)
    
    If vSoc.LeerDatos(Socio) Then
        
        If TotalFac > 0 Then ' se insertara en la cartera de pagos (spagop)
            CadValues2 = ""
        
            'vamos creando la cadena para insertar en spagosp de la CONTA
            letraser = ""
            letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(tipoMov, "T"))
            
            '[Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
            
            '[Monica]03/07/2013: añado trim(codmacta)
            CadValuesAux2 = "("
            If vParamAplic.ContabilidadNueva Then CadValuesAux2 = CadValuesAux2 & "'" & SerieFraPro & "',"
            CadValuesAux2 = CadValuesAux2 & "'" & Trim(CtaSocio) & "', " & DBSet(FacturaSoc, "T") & ", '" & Format(FecFactuSoc, FormatoFecha) & "', "
        
            '------------------------------------------------------------
            I = 1
            CadValues2 = CadValuesAux2 & I
            
            CadValues2 = CadValues2 & ", " & ForpaPosi & ", '" & Format(FecVenci, FormatoFecha) & "', "
            CadValues2 = CadValues2 & DBSet(TotalFac, "N") & ", " & DBSet(CtaBanco, "T") & ","
        
            If Not vParamAplic.ContabilidadNueva Then
                'David. Para que ponga la cuenta bancaria (SI LA tiene)
                CadValues2 = CadValues2 & DBSet(BancoSoc, "T", "S") & "," & DBSet(SucurSoc, "T", "S") & ","
                CadValues2 = CadValues2 & DBSet(DigcoSoc, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & ","
            End If
            
            'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
            Sql = "Fact.: " & letraser & "-" & FacturaSoc & "-" & Format(FecFactuSoc, "dd/mm/yyyy")
                
            CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
            
            Sql = ""
            CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" ')"
            If vParamAplic.ContabilidadNueva Then
                vvIban = MiFormat(IbanSoc, "") & MiFormat(CStr(BancoSoc), "0000") & MiFormat(CStr(SucurSoc), "0000") & MiFormat(DigcoSoc, "00") & MiFormat(CtaBaSoc, "0000000000")
                
                CadValues2 = CadValues2 & ", " & DBSet(vvIban, "T") & ","
                'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                CadValues2 = CadValues2 & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES') "
            
            
            Else
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & ") "
                Else
                    CadValues2 = CadValues2 & ") "
                End If
            End If
        
            'Grabar tabla spagop de la CONTABILIDAD
            '-------------------------------------------------
            If CadValues2 <> "" Then
                'Insertamos en la tabla spagop de la CONTA
                'David. Cuenta bancaria y descripcion textos
                If vParamAplic.ContabilidadNueva Then
                    Sql = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, iban,"
                    Sql = Sql & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais)"
                Else
                    Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        Sql = Sql & ", iban) "
                    Else
                        Sql = Sql & ") "
                    End If
                End If
                Sql = Sql & " VALUES " & CadValues2
                ConnConta.Execute Sql
            End If
            
        Else
            ' si es negativo se inserta en positivo en la cartera de cobros (scobro)
    
            letraser = ""
            letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(tipoMov, "T"))
    
    '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
    '        Text33csb = "'Factura:" & DBLet(letraser, "T") & "-" & DBLet(numfactu, "T") & " " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
            Text33csb = "'Factura:" & DBLet(FacturaSoc, "T") & " " & Format(DBLet(FecFactuSoc, "F"), "dd/mm/yy") & "'"
            Text41csb = "de " & DBSet(TotalFac, "N")
            Text42csb = ""
    
            CC = DBLet(DigcoSoc, "T")
            If DBLet(DigcoSoc, "T") = "**" Then CC = "00"
                
            '[Monica]03/07/2013: añado trim(codmacta)
            CadValuesAux2 = "(" & DBSet(letraser, "T") & "," & DBSet(FacturaSoc, "N") & "," & DBSet(FecFactuSoc, "F") & ", 1," & DBSet(Trim(CtaSocio), "T") & ","
            CadValues2 = CadValuesAux2 & DBSet(ForpaNega, "N") & "," & DBSet(FecFactuSoc, "F") & "," & DBSet(TotalFac * (-1), "N") & ","
            If Not vParamAplic.ContabilidadNueva Then
                CadValues2 = CadValues2 & DBSet(CtaBanco, "T") & "," & DBSet(BancoSoc, "N", "S") & "," & DBSet(SucurSoc, "N", "S") & ","
                CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & ","
                CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            Else
                CadValues2 = CadValues2 & DBSet(CtaBanco, "T") & "," & ValorNulo & "," & ValorNulo & ","
            End If
            CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & "," & DBSet(Text42csb, "T") & ",1" ')"
            
            If vParamAplic.ContabilidadNueva Then
                vvIban = MiFormat(IbanSoc, "") & MiFormat(CStr(BancoSoc), "0000") & MiFormat(CStr(SucurSoc), "0000") & MiFormat(CC, "00") & MiFormat(CtaBaSoc, "0000000000")
                
                CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                CadValues2 = CadValues2 & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES') "
    
                'Insertamos en la tabla scobro de la CONTA
                Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                Sql = Sql & "ctabanc1,  fecultco, impcobro, "
                Sql = Sql & " text33csb, text41csb, text42csb, agente, iban, " ') "
                Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
                Sql = Sql & ") "
            
            Else
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    CadValues2 = CadValues2 & "," & DBSet(IbanSoc, "T", "S") & ") "
                Else
                    CadValues2 = CadValues2 & ") "
                End If
                
        
                'Insertamos en la tabla scobro de la CONTA
                Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                Sql = Sql & " text33csb, text41csb, text42csb, agente" ') "
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    Sql = Sql & ", iban) "
                Else
                    Sql = Sql & ") "
                End If
            End If
            
            Sql = Sql & " VALUES " & CadValues2
            ConnConta.Execute Sql
    
        End If
    
        B = True
    End If
    
    Set vSoc = Nothing
    
    
EInsertarTesoreria:
    If Err.Number <> 0 Then
        B = False
        MenError = "Error al insertar en Tesoreria: " & Err.Description
    End If
    InsertarEnTesoreriaNewFVARPro = B
End Function



'############################################################################
'################ INSERTAR EN DIARIO EL ASIENTO DE COBRO DE RMT
'############################################################################

Private Function InsertarAsientoCobroPOZOS(cadMen As String, ByRef Rs As ADODB.Recordset, FecRec As Date, CtaContra As String) As Boolean

' la contabilizacion de las facturas internas de horto, se han de insertar en el diario no en el registro de iva de proveedor
Dim Sql As String
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Concepto As String
Dim letraser As String
Dim Obs As String
Dim B As Boolean
'Dim CtaSocio As String

Dim Mc As Contadores
    On Error GoTo EInsertar
       
    cad = ""
    Set Mc = New Contadores

    If Mc.ConseguirContador("0", (DBLet(Rs!fecfactu) <= CDate(vEmpresa.FechaFin)), True) = 0 Then
    
        Sql = "select codmaccli from rsocios_seccion where codsecci = " & vParamAplic.SeccionPOZOS & " and codsocio = " & DBSet(Rs!Codsocio, "N")
        CtaSocio = DevuelveValor(Sql)
        
        '[Monica]18/06/2014: antes poniamos la fecha de factura, ahora la fecha de hoy
        Obs = "Contabilización Cobro Rec.Manta " & Format(Now, "dd/mm/yyyy")
    
        'Insertar en la conta Cabecera Asiento
        cadMen = ""
        B = InsertarCabAsientoDia(1, Mc.Contador, CStr(Format(Rs!fecfactu, "dd/mm/yyyy")), Obs, cadMen)
        cadMen = "Insertando Cab. Asiento: " & cadMen

        If B Then
            cadMen = ""
            B = InsertarLinAsientoCobroPOZOS(Rs, cadMen, CtaSocio, CtaContra, Mc.Contador)
            cadMen = "Insertando Lin. Asiento Diario: " & cadMen
        
        End If
        
        If B And Not vParamAplic.ContabilidadNueva Then
        
            ProcesoCorrecto = False
        
            frmActualizar2.Numasiento = Mc.Contador
            frmActualizar2.FechaAsiento = Rs!fecfactu
            frmActualizar2.numdiari = vEmpresa.NumDiarioInt
            frmActualizar2.OpcionActualizar = 1
            frmActualizar2.Show vbModal
            
            B = ProcesoCorrecto
        End If
            
        Set Mc = Nothing
        
        
    End If
    
EInsertar:
    If Err.Number <> 0 Or Not B Then
        InsertarAsientoCobroPOZOS = False
        cadMen = cadMen & Err.Description
    Else
        InsertarAsientoCobroPOZOS = B And ProcesoCorrecto
    End If
End Function


Private Function InsertarLinAsientoCobroPOZOS(ByRef Rs As ADODB.Recordset, cadErr As String, CtaSocio As String, CtaContra As String, Contador As Long) As Boolean
Dim Sql As String
Dim numdocum As String
Dim Ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim I As Long
Dim B As Boolean
Dim cad As String
Dim FeFact As Date
Dim cadMen As String

Dim letraser As String
Dim Concep As Integer
Dim Amplia As String

    On Error GoTo eInsertarLinAsientoCobroPOZOS

    InsertarLinAsientoCobroPOZOS = False
        
        
    I = 0
    ImporteD = 0
    ImporteH = 0
    
    letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(Rs!CodTipom, "T"))
    
    numdocum = letraser & Format(Rs!numfactu, "0000000")
    
    Concep = 3
    
    Amplia = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", CStr(Concep), "N"))
    
    ampliaciond = Amplia & " " & letraser & "/" & DBLet(Rs!numfactu, "N")
    ampliacionh = Amplia & " " & letraser & "/" & DBLet(Rs!numfactu, "N")
    
    B = True
    
    I = I + 1
    
    FeFact = Rs!fecfactu
    
    cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Contador, "N") & ","
    cad = cad & DBSet(I, "N") & "," & DBSet(CtaSocio, "T") & "," & DBSet(numdocum, "T") & ","
    
    ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
    If DBLet(Rs!TotalFact, "N") > 0 Then
        ' importe al haber en positivo
        cad = cad & DBSet(Concep, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
        cad = cad & DBSet(Rs!TotalFact, "N") & "," & ValorNulo & "," & DBSet(CtaContra, "T") & "," & ValorNulo & ",0"
    
        ImporteH = ImporteH + CCur(DBLet(Rs!TotalFact, "N"))
        
    Else
        ' importe al debe en positivo cambiamos signo
        cad = cad & DBSet(Concep, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(DBLet(Rs!TotalFact, "N") * (-1), "N") & ","
        cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaContra, "T") & "," & ValorNulo & ",0"
    
        ImporteD = ImporteD + CCur(DBLet(Rs!TotalFact, "N") * (-1))
    
    End If
    
    cad = "(" & cad & ")"
    
    B = InsertarLinAsientoDia(cad, cadMen)
    cadMen = "Insertando Lin. Asiento: " & I

    I = I + 1
            
    ' el Total es sobre la cuenta del cliente
    cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Contador, "N") & ","
    cad = cad & DBSet(I, "N") & ","
    cad = cad & DBSet(CtaContra, "T") & "," & DBSet(numdocum, "T") & ","
        
    ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
    If DBLet(Rs!TotalFact, "N") > 0 Then
        ' importe al debe en positivo
        cad = cad & DBSet(Concep, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(DBLet(Rs!TotalFact, "N"), "N") & ","
        cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaSocio, "N") & "," & ValorNulo & ",0"
    Else
        ' importe al haber en positivo, cambiamos el signo
        cad = cad & DBSet(Concep, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
        cad = cad & DBSet(DBLet(Rs!TotalFact, "N") * (-1), "N") & "," & ValorNulo & "," & DBSet(CtaSocio, "N") & "," & ValorNulo & ",0"
    End If
    
    cad = "(" & cad & ")"
    
    B = InsertarLinAsientoDia(cad, cadMen)
    cadMen = "Insertando Lin. Asiento: " & I

    InsertarLinAsientoCobroPOZOS = B
    Exit Function
    
eInsertarLinAsientoCobroPOZOS:
    cadErr = "Insertar Linea Asiento Cobro Pozos: " & Err.Description
    cadErr = cadErr & cadMen
    InsertarLinAsientoCobroPOZOS = False
End Function



Public Function ComprobarSociosSeccion(cadTabla As String, Seccion As Integer) As Boolean
'Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
'vamos a contabilizar
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim B As Boolean

    On Error GoTo ECompFactu

    ComprobarSociosSeccion = False
    
    If cadTabla = "rrecibpozos" Then
        Sql = "SELECT DISTINCT rrecibpozos.codsocio "
        Sql = Sql & " FROM (rrecibpozos LEFT JOIN rsocios_seccion ON rrecibpozos.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N") & ") "
        Sql = Sql & " INNER JOIN tmpFactu ON rrecibpozos.codtipom=tmpFactu.codtipom AND rrecibpozos.numfactu=tmpFactu.numfactu AND rrecibpozos.fecfactu=tmpFactu.fecfactu "

        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        B = True
        While Not Rs.EOF And B
            Sql2 = "select * from rsocios_seccion where (codsocio= " & DBSet(Rs!Codsocio, "N") & " and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionPOZOS, "N") & ")"
            If RegistrosAListar(Sql2, cAgro) = 0 Then
                B = False
                
                Select Case cadTabla
                    Case "rrecibpozos"
                        Sql = "Socio no existente en la sección de pozos: " & DBSet(Rs!Codsocio, "N") & vbCrLf
                End Select
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not B Then
            Sql = "Comprobando Socios en Sección ...       " & vbCrLf & vbCrLf & Sql
            
            MsgBox Sql, vbExclamation
            ComprobarSociosSeccion = False
        Else
            ComprobarSociosSeccion = True
        End If
    End If
     
    Exit Function
    
ECompFactu:
     If Err.Number <> 0 Then
        ComprobarSociosSeccion = False
        MuestraError Err.Number, "Comprobar Socios Sección", Err.Description
    End If
End Function



Public Function InsertarEnTesoreriaBajaCampo(cadWHERE As String, MenError As String, Fra As String, fecfactu As String, ForpaPosi As String, CtaBanco As String, Socio As String, Campos As String) As Boolean
'Guarda datos de Tesoreria en tablas: spagop o scobro dependiendo del signo de la factura
Dim B As Boolean
Dim Sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim CC As String
Dim letraser As String
Dim Text33csb As String
Dim Text41csb As String
Dim Text42csb As String
Dim GastosPie As Currency
Dim GastosVarias As Currency
Dim FactuRec As String
Dim rsVenci As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim FecVenci1 As Date
Dim ImpVenci As Currency

Dim TotalTesor1 As Currency

Dim UltimoVto As Integer

Dim CadValuesGastos As String
Dim CadValuesVarias As String
Dim SqlGastos As String
Dim J As Integer
Dim CtaSocio As String
Dim cadena As String
Dim vSoc As cSocio


    On Error GoTo EInsertarTesoreriaSoc

    InsertarEnTesoreriaBajaCampo = False
    
    
    Sql = "select sum(importe) from (" & cadWHERE & ") aa"
    TotalTesor = DevuelveValor(Sql)
    
    
    If TotalTesor > 0 Then ' se insertara en la cartera de pagos (spagop)
        
        '[Monica]09/05/2013: Añadido el nro de vencimientos
        CadValues2 = ""
        
        Sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & ForpaPosi
        Set rsVenci = New ADODB.Recordset
        rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not rsVenci.EOF Then
            If rsVenci!numerove > 0 Then
            
                Set vSoc = New cSocio
                If vSoc.LeerDatos(Socio) Then
                
                    'Obtener los dias de pago de la tabla de parametros: spara1
                    cadena = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
                    'La raiz de aportacion está fija
                    CtaSocio = "1501" & Format(Socio, cadena)
                
                    
                    '[Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
                    '[Monica]03/07/2013: añado trim(codmacta)
                    CadValuesAux2 = "("
                    If vParamAplic.ContabilidadNueva Then CadValuesAux2 = CadValuesAux2 & DBSet(SerieFraPro, "T") & ","
                    CadValuesAux2 = CadValuesAux2 & "'" & Trim(CtaSocio) & "', " & DBSet(Fra, "T") & ", '" & Format(fecfactu, FormatoFecha) & "', "
                    
                      'Primer Vencimiento
                      '------------------------------------------------------------
                      I = 1
                      'FECHA VTO
                      FecVenci = CDate(fecfactu)
                      '=== Modificado: Laura 23/01/2007
        '              FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                      FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                      '==================================
                      'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                      
                      FecVenci1 = FecVenci
        
        
                      CadValues2 = CadValuesAux2 & I
                      CadValues2 = CadValues2 & ", " & ForpaPosi & ", '" & Format(FecVenci, FormatoFecha) & "', "
                      
                      
                      'IMPORTE del Vencimiento
                      If rsVenci!numerove = 1 Then
                          ImpVenci = TotalTesor
                      Else
                          ImpVenci = Round(TotalTesor / rsVenci!numerove, 2)
                          'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                          If ImpVenci * rsVenci!numerove <> TotalTesor Then
                              ImpVenci = Round(ImpVenci + (TotalTesor - ImpVenci * rsVenci!numerove), 2)
                          End If
                      End If
                      
                      CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBanco, "T") & ","
                
                      If Not vParamAplic.ContabilidadNueva Then
                            'David. Para que ponga la cuenta bancaria (SI LA tiene)
                            CadValues2 = CadValues2 & DBSet(vSoc.Banco, "T", "S") & "," & DBSet(vSoc.Sucursal, "T", "S") & ","
                            CadValues2 = CadValues2 & DBSet(vSoc.Digcontrol, "T", "S") & "," & DBSet(vSoc.CuentaBan, "T", "S") & ","
                      End If
                
                      Sql = "Pago Capital Social Campos " & Fra & "-" & Format(fecfactu, "dd/mm/yyyy")
                        
                      CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
                    
                      Sql = "Campos: " & Campos
                      CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" '),"
                      
                      If vParamAplic.ContabilidadNueva Then

                            vvIban = MiFormat(vSoc.Iban, "") & MiFormat(vSoc.Banco, "0000") & MiFormat(vSoc.Sucursal, "0000") & MiFormat(vSoc.Digcontrol, "00") & MiFormat(vSoc.CuentaBan, "0000000000")
                            
                            CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                            'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                            CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                            CadValues2 = CadValues2 & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'," & ValorNulo & "," & ValorNulo & ",0),"
                            
                      Else
                            '[Monica]22/11/2013: Tema iban
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                CadValues2 = CadValues2 & ", " & DBSet(vSoc.Iban, "T", "S") & "),"
                            Else
                                CadValues2 = CadValues2 & "),"
                            End If
                      
                      End If
                      'Resto Vencimientos
                      '--------------------------------------------------------------------
                      UltimoVto = I
                      For J = 2 To rsVenci!numerove
                          UltimoVto = I + J - 1
                         'FECHA Resto Vencimientos
                          '==== Modificado: Laura 23/01/2007
                          'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                          FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                          '==================================================
        
                          CadValues2 = CadValues2 & CadValuesAux2 & UltimoVto 'i
                          CadValues2 = CadValues2 & ", " & ForpaPosi & ", '" & Format(FecVenci, FormatoFecha) & "', "
        
                          'IMPORTE Resto de Vendimientos
                          ImpVenci = Round(TotalTesor / rsVenci!numerove, 2)
        
                          CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & "," & DBSet(CtaBanco, "T") & ","
                          
                          If Not vParamAplic.ContabilidadNueva Then
                                'David. Para que ponga la cuenta bancaria (SI LA tiene)
                                CadValues2 = CadValues2 & DBSet(vSoc.Banco, "T", "S") & "," & DBSet(vSoc.Sucursal, "T", "S") & ","
                                CadValues2 = CadValues2 & DBSet(vSoc.Digcontrol, "T", "S") & "," & DBSet(vSoc.CuentaBan, "T", "S") & ","
                          End If
                          
                          Sql = "Pago Capital Social Campos " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                            
                          CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
                        
                          Sql = "Campos: " & Campos
                          CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" '),"
                          
                          If vParamAplic.ContabilidadNueva Then
                                
                                vvIban = MiFormat(vSoc.Iban, "") & MiFormat(vSoc.Banco, "0000") & MiFormat(vSoc.Sucursal, "0000") & MiFormat(vSoc.Digcontrol, "00") & MiFormat(vSoc.CuentaBan, "0000000000")
                                
                                CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                                'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                                CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                                CadValues2 = CadValues2 & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'," & ValorNulo & "," & ValorNulo & ",0),"
                          Else
                                
                                '[Monica]22/11/2013: Tema iban
                                If vEmpresa.HayNorma19_34Nueva = 1 Then
                                    CadValues2 = CadValues2 & ", " & DBSet(vSoc.Iban, "T", "S") & "),"
                                Else
                                    CadValues2 = CadValues2 & "),"
                                End If
                          End If
                      Next J
                      
                    If CadValues2 <> "" Then
                        CadValues2 = Mid(CadValues2, 1, Len(CadValues2) - 1)
                    
                        'Insertamos en la tabla spagop de la CONTA
                        'David. Cuenta bancaria y descripcion textos
                        
                        If vParamAplic.ContabilidadNueva Then
                            Sql = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, iban,"
                            Sql = Sql & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais, fecultpa, imppagad, situacion)"
                        
                        
                        Else
                            Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
                            '[Monica]22/11/2013: Tema iban
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                Sql = Sql & ", iban) "
                            Else
                                Sql = Sql & ") "
                            End If
                        End If
                        
                        Sql = Sql & " VALUES " & CadValues2
                        ConnConta.Execute Sql
                        
                    End If
                End If
                Set vSoc = Nothing
                    
            End If
        End If
        
    End If

    B = True

EInsertarTesoreriaSoc:
    If Err.Number <> 0 Then
        B = False
        MenError = "Error al insertar en Tesoreria Baja Campo: " & Err.Description
    End If
    InsertarEnTesoreriaBajaCampo = B
End Function
