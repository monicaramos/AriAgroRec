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


Public Function CrearTMPFacturas(cadTabla As String, cadwhere As String) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim sql As String
    
    On Error GoTo ECrear
    
    CrearTMPFacturas = False
    
    sql = "CREATE TEMPORARY TABLE tmpFactu ( "
    If cadTabla = "facturas" Then
        sql = sql & "codtipom char(3) NOT NULL default '',"
        sql = sql & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Else
        If cadTabla = "rfactsoc" Or cadTabla = "advfacturas" Or cadTabla = "rbodfacturas" Or cadTabla = "fvarcabfact" Or cadTabla = "fvarcabfactpro" Then
            sql = sql & "codtipom char(3) NOT NULL default '',"
            sql = sql & "numfactu int(7)  NOT NULL ,"
        Else
            If cadTabla = "rcabfactalmz" Then
                sql = sql & "tipofichero smallint(1) unsigned NOT NULL,"
                sql = sql & "codsocio smallint(3) unsigned NOT NULL default '0',"
                sql = sql & "numfactu int(7)  NOT NULL ,"
            Else
                If cadTabla = "rtelmovil" Then
                    sql = sql & "numserie varchar(1) NOT NULL,"
                    sql = sql & "numfactu int(7)     NOT NULL,"
                Else
                    If cadTabla = "rrecibpozos" Then
                        sql = sql & "codtipom char(3) NOT NULL,"
                        sql = sql & "numfactu int(7) unsigned NOT NULL,"
                    Else
                        If cadTabla = "rfacttra" Then
                            sql = sql & "codtipom char(3) NOT NULL default '',"
                            sql = sql & "numfactu int(7)  NOT NULL ,"
                        Else
                            sql = sql & "codsocio int(7) unsigned NOT NULL default '0',"
                            sql = sql & "numfactu varchar(10)  NOT NULL ,"
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    sql = sql & "fecfactu date NOT NULL default '0000-00-00' "
    
    If cadTabla = "rfacttra" Then
        sql = sql & ",codtrans varchar(10))"
    Else
        sql = sql & ")"
    End If
    
    conn.Execute sql
     
     
    If cadTabla = "facturas" Or cadTabla = "advfacturas" Or cadTabla = "rbodfacturas" Or cadTabla = "fvarcabfact" Or cadTabla = "fvarcabfactpro" Then
        sql = "SELECT codtipom, numfactu, fecfactu"
    Else
        If cadTabla = "rfactsoc" Then
            sql = "SELECT codtipom, numfactu, fecfactu"
        Else
            If cadTabla = "rcabfactalmz" Then
                sql = "SELECT tipofichero, codsocio, numfactu, fecfactu "
            Else
                If cadTabla = "rtelmovil" Then
                    sql = "SELECT numserie, numfactu, fecfactu "
                Else
                    If cadTabla = "rfacttra" Then
                        sql = "SELECT codtipom, numfactu, fecfactu, codtrans"
                    Else
                        If cadTabla = "rrecibpozos" Then
                            sql = "SELECT DISTINCT codtipom, numfactu, fecfactu "
                        Else
                            sql = "SELECT codsocio, numfactu, fecfactu"
                        End If
                    End If
                End If
            End If
        End If
    End If
    sql = sql & " FROM " & cadTabla
    sql = sql & " WHERE " & cadwhere
    sql = " INSERT INTO tmpFactu " & sql
    conn.Execute sql

    CrearTMPFacturas = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPFacturas = False
        'Borrar la tabla temporal
        sql = " DROP TABLE IF EXISTS tmpFactu;"
        conn.Execute sql
    End If
End Function


Public Sub BorrarTMPFacturas()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmpFactu;"
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub InsertarTMPErrFac(MenError As String, cadwhere As String)
Dim sql As String

    On Error Resume Next
    sql = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
    sql = sql & " Select *," & DBSet(Mid(MenError, 1, 200), "T") & " as error From tmpFactu "
    sql = sql & " WHERE " & Replace(cadwhere, "rfactsoc", "tmpFactu")
    conn.Execute sql
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub InsertarTMPErrFacSoc(MenError As String, cadwhere As String)
Dim sql As String

    On Error Resume Next
    sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
    sql = sql & " Select *," & DBSet(Mid(MenError, 1, 200), "T") & " as error From tmpFactu "
    sql = sql & " WHERE " & Replace(cadwhere, "rfactsoc", "tmpFactu")
    conn.Execute sql
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub InsertarTMPErrFacFVAR(MenError As String, cadwhere As String)
Dim sql As String

    On Error Resume Next
    sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
    sql = sql & " Select *," & DBSet(Mid(MenError, 1, 200), "T") & " as error From tmpFactu "
    sql = sql & " WHERE " & Replace(cadwhere, "fvarcabfact", "tmpFactu")
    conn.Execute sql
    
    If Err.Number <> 0 Then Err.Clear
End Sub





Public Function CrearTMPErrFact(cadTabla As String) As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim sql As String
    
    On Error GoTo ECrear
    
    CrearTMPErrFact = False
    
    sql = "CREATE TEMPORARY TABLE tmpErrFac ( "
    If cadTabla = "facturas" Or cadTabla = "rfactsoc" Or cadTabla = "rbodfacturas" Or cadTabla = "rrecibpozos" Or cadTabla = "fvarcabfact" Or cadTabla = "fvarcabfactpro" Then
        sql = sql & "codtipom char(3) NOT NULL default '',"
        sql = sql & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Else
        If cadTabla = "rcabfactalmz" Then
            sql = sql & "tipofichero smallint unsigned NOT NULL, "
            sql = sql & "numfactu int(7) NOT NULL ,"
        Else
            sql = sql & "codprove int(6) unsigned NOT NULL default '0',"
            sql = sql & "numfactu varchar(10) NOT NULL ,"
        End If
    End If
    sql = sql & "fecfactu date NOT NULL default '0000-00-00', "
    
    If cadTabla = "rcabfactalmz" Then sql = sql & "codsocio int(7) ,"
    
    sql = sql & "error varchar(200) NULL )"
    
    conn.Execute sql
     
    CrearTMPErrFact = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrFact = False
        'Borrar la tabla temporal
        sql = " DROP TABLE IF EXISTS tmpErrFac;"
        conn.Execute sql
    End If
End Function

Public Function CrearTMPErrComprob() As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim sql As String
    
    On Error GoTo ECrear
    
    CrearTMPErrComprob = False
    
    sql = "CREATE TEMPORARY TABLE tmperrcomprob ( "
    sql = sql & "error varchar(100) NULL )"
    conn.Execute sql
     
    CrearTMPErrComprob = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrComprob = False
        'Borrar la tabla temporal
        sql = " DROP TABLE IF EXISTS tmperrcomprob;"
        conn.Execute sql
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
Dim sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim Cad As String, devuelve As String
Dim Sql2 As String
Dim total As Long

On Error GoTo EComprobarLetra

    ComprobarLetraSerie = False
    
    'Comprobar que existe la letra de serie en contabilidad
    Select Case cadTabla
        Case "rfactsoc"
            'cargamos el RSConta con la tabla contadores de BD: Contabilidad
            'donde estan todas las letra de serie que existen en la contabilidad
            sql = "Select distinct tiporegi from contadores"
            Set RSconta = New ADODB.Recordset
            RSconta.Open sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
            If RSconta.EOF Then
                RSconta.Close
                Set RSconta = Nothing
                Exit Function
            End If
                
        
            'obtenemos los distintos tipos de movimiento que vamos a contabilizar
            'de las facturas seleccionadas
            sql = "select distinct rfactsoc.codtipom from " & cadTabla
            sql = sql & " INNER JOIN tmpFactu ON rfactsoc.codtipom=tmpFactu.codtipom AND rfactsoc.numfactu=tmpFactu.numfactu AND rfactsoc.fecfactu=tmpFactu.fecfactu "
    '        SQL = SQL & cadWHERE
            Set Rs = New ADODB.Recordset
            Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad = ""
            b = True
            While Not Rs.EOF And b
                'comprobar que todas las letras serie existen en usuarios
                sql = "letraser"
                Sql2 = "select letraser from usuarios.stipom where codtipom = " & DBSet(Rs!CodTipom, "T")
                'devuelve = DevuelveDesdeBDNew(cAgro, "stipom", "codtipom", "codtipom", RS!codtipom, "T", SQL)
                total = TotalRegistrosConsulta(Sql2)
                If total = 0 Then 'devuelve = "" Then
                    b = False
                    Cad = Rs!CodTipom & " en BD de Ariagrorec."
                ElseIf DevuelveValor(Sql2) <> "" Then 'SQL <> "" Then
                    'comprobar que todas las letras serie existen en la contabilidad
                    devuelve = "tiporegi= " & DBSet(DevuelveValor(Sql2), "T") 'SQL, "T")
                    RSconta.MoveFirst
                    RSconta.Find (devuelve), , adSearchForward
                    If RSconta.EOF Then
                        'no encontrado
                        b = False
                        'Cad = SQL & " en BD de Contabilidad."
                        Cad = DevuelveValor(Sql2) & " en BD de Contabilidad."
                    End If
                End If
                If b Then Cad = Cad & DBSet(Rs!CodTipom, "T") & ","
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
            RSconta.Close
            Set RSconta = Nothing
            
            If Not b Then 'Hay algun movimiento que no existe
                devuelve = "No existe el tipo de movimiento: " & Cad & vbCrLf
                devuelve = devuelve & "Consulte con el administrador."
                MsgBox devuelve, vbExclamation
                Exit Function
            End If
            
            'Todos los Tipo de movimiento existen
            If Cad <> "" Then
                Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitamos ult. coma
            
                'miramos si hay algun movimiento de factura que la letra serie sea nulo
                sql = "select count(*) from usuarios.stipom "
                sql = sql & "where codtipom IN (" & Cad & ") and (isnull(letraser) or letraser='')"
                If RegistrosAListar(sql) > 0 Then
                    sql = "Hay algun tipo de movimiento de Facturación que no tiene letra serie." & vbCrLf
                    sql = sql & "Comprobar en la tabla de tipos de movimiento: " & Cad
                    MsgBox sql, vbExclamation
                    Exit Function
                End If
            End If
            ComprobarLetraSerie = True
        Case "advfacturas"
            'cargamos el RSConta con la tabla contadores de BD: Contabilidad
            'donde estan todas las letra de serie que existen en la contabilidad
            sql = "Select distinct tiporegi from contadores"
            Set RSconta = New ADODB.Recordset
            RSconta.Open sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
            If RSconta.EOF Then
                RSconta.Close
                Set RSconta = Nothing
                Exit Function
            End If
                
        
            'obtenemos los distintos tipos de movimiento que vamos a contabilizar
            'de las facturas seleccionadas
            sql = "select distinct advfacturas.codtipom from " & cadTabla
            sql = sql & " INNER JOIN tmpFactu ON advfacturas.codtipom=tmpFactu.codtipom AND advfacturas.numfactu=tmpFactu.numfactu AND advfacturas.fecfactu=tmpFactu.fecfactu "
    '        SQL = SQL & cadWHERE
            Set Rs = New ADODB.Recordset
            Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad = ""
            b = True
            While Not Rs.EOF And b
                'comprobar que todas las letras serie existen en usuarios
                sql = "letraser"
                Sql2 = "select letraser from usuarios.stipom where codtipom = " & DBSet(Rs!CodTipom, "T")
                'devuelve = DevuelveDesdeBDNew(cAgro, "stipom", "codtipom", "codtipom", RS!codtipom, "T", SQL)
                total = TotalRegistrosConsulta(Sql2)
                If total = 0 Then 'devuelve = "" Then
                    b = False
                    Cad = Rs!CodTipom & " en BD de Ariagrorec."
                ElseIf DevuelveValor(Sql2) <> "" Then 'SQL <> "" Then
                    'comprobar que todas las letras serie existen en la contabilidad
                    devuelve = "tiporegi= " & DBSet(DevuelveValor(Sql2), "T") 'SQL, "T")
                    RSconta.MoveFirst
                    RSconta.Find (devuelve), , adSearchForward
                    If RSconta.EOF Then
                        'no encontrado
                        b = False
                        'Cad = SQL & " en BD de Contabilidad."
                        Cad = DevuelveValor(Sql2) & " en BD de Contabilidad."
                    End If
                End If
                If b Then Cad = Cad & DBSet(Rs!CodTipom, "T") & ","
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
            RSconta.Close
            Set RSconta = Nothing
            
            If Not b Then 'Hay algun movimiento que no existe
                devuelve = "No existe el tipo de movimiento: " & Cad & vbCrLf
                devuelve = devuelve & "Consulte con el administrador."
                MsgBox devuelve, vbExclamation
                Exit Function
            End If
            
            'Todos los Tipo de movimiento existen
            If Cad <> "" Then
                Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitamos ult. coma
            
                'miramos si hay algun movimiento de factura que la letra serie sea nulo
                sql = "select count(*) from usuarios.stipom "
                sql = sql & "where codtipom IN (" & Cad & ") and (isnull(letraser) or letraser='')"
                If RegistrosAListar(sql) > 0 Then
                    sql = "Hay algun tipo de movimiento de Facturación que no tiene letra serie." & vbCrLf
                    sql = sql & "Comprobar en la tabla de tipos de movimiento: " & Cad
                    MsgBox sql, vbExclamation
                    Exit Function
                End If
            End If
            ComprobarLetraSerie = True
    
        Case "rrecibpozos"
            'cargamos el RSConta con la tabla contadores de BD: Contabilidad
            'donde estan todas las letra de serie que existen en la contabilidad
            sql = "Select distinct tiporegi from contadores"
            Set RSconta = New ADODB.Recordset
            RSconta.Open sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
            If RSconta.EOF Then
                RSconta.Close
                Set RSconta = Nothing
                Exit Function
            End If
                
        
            'obtenemos los distintos tipos de movimiento que vamos a contabilizar
            'de las facturas seleccionadas
            sql = "select distinct rrecibpozos.codtipom from " & cadTabla
            sql = sql & " INNER JOIN tmpFactu ON rrecibpozos.codtipom=tmpFactu.codtipom AND rrecibpozos.numfactu=tmpFactu.numfactu AND rrecibpozos.fecfactu=tmpFactu.fecfactu "
    '        SQL = SQL & cadWHERE
            Set Rs = New ADODB.Recordset
            Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad = ""
            b = True
            While Not Rs.EOF And b
                'comprobar que todas las letras serie existen en usuarios
                sql = "letraser"
                Sql2 = "select letraser from usuarios.stipom where codtipom = " & DBSet(Rs!CodTipom, "T")
                'devuelve = DevuelveDesdeBDNew(cAgro, "stipom", "codtipom", "codtipom", RS!codtipom, "T", SQL)
                total = TotalRegistrosConsulta(Sql2)
                If total = 0 Then 'devuelve = "" Then
                    b = False
                    Cad = Rs!CodTipom & " en BD de Ariagrorec."
                ElseIf DevuelveValor(Sql2) <> "" Then 'SQL <> "" Then
                    'comprobar que todas las letras serie existen en la contabilidad
                    devuelve = "tiporegi= " & DBSet(DevuelveValor(Sql2), "T") 'SQL, "T")
                    RSconta.MoveFirst
                    RSconta.Find (devuelve), , adSearchForward
                    If RSconta.EOF Then
                        'no encontrado
                        b = False
                        'Cad = SQL & " en BD de Contabilidad."
                        Cad = DevuelveValor(Sql2) & " en BD de Contabilidad."
                    End If
                End If
                If b Then Cad = Cad & DBSet(Rs!CodTipom, "T") & ","
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
            RSconta.Close
            Set RSconta = Nothing
            
            If Not b Then 'Hay algun movimiento que no existe
                devuelve = "No existe el tipo de movimiento: " & Cad & vbCrLf
                devuelve = devuelve & "Consulte con el administrador."
                MsgBox devuelve, vbExclamation
                Exit Function
            End If
            
            'Todos los Tipo de movimiento existen
            If Cad <> "" Then
                Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitamos ult. coma
            
                'miramos si hay algun movimiento de factura que la letra serie sea nulo
                sql = "select count(*) from usuarios.stipom "
                sql = sql & "where codtipom IN (" & Cad & ") and (isnull(letraser) or letraser='')"
                If RegistrosAListar(sql) > 0 Then
                    sql = "Hay algun tipo de movimiento de Facturación que no tiene letra serie." & vbCrLf
                    sql = sql & "Comprobar en la tabla de tipos de movimiento: " & Cad
                    MsgBox sql, vbExclamation
                    Exit Function
                End If
            End If
            ComprobarLetraSerie = True
    
    
    
    
        Case "rbodfacturas"
            'cargamos el RSConta con la tabla contadores de BD: Contabilidad
            'donde estan todas las letra de serie que existen en la contabilidad
            sql = "Select distinct tiporegi from contadores"
            Set RSconta = New ADODB.Recordset
            RSconta.Open sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
            If RSconta.EOF Then
                RSconta.Close
                Set RSconta = Nothing
                Exit Function
            End If
                
        
            'obtenemos los distintos tipos de movimiento que vamos a contabilizar
            'de las facturas seleccionadas
            sql = "select distinct rbodfacturas.codtipom from " & cadTabla
            sql = sql & " INNER JOIN tmpFactu ON rbodfacturas.codtipom=tmpFactu.codtipom AND rbodfacturas.numfactu=tmpFactu.numfactu AND rbodfacturas.fecfactu=tmpFactu.fecfactu "
    '        SQL = SQL & cadWHERE
            Set Rs = New ADODB.Recordset
            Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad = ""
            b = True
            While Not Rs.EOF And b
                'comprobar que todas las letras serie existen en usuarios
                sql = "letraser"
                Sql2 = "select letraser from usuarios.stipom where codtipom = " & DBSet(Rs!CodTipom, "T")
                'devuelve = DevuelveDesdeBDNew(cAgro, "stipom", "codtipom", "codtipom", RS!codtipom, "T", SQL)
                total = TotalRegistrosConsulta(Sql2)
                If total = 0 Then 'devuelve = "" Then
                    b = False
                    Cad = Rs!CodTipom & " en BD de Ariagrorec."
                ElseIf DevuelveValor(Sql2) <> "" Then 'SQL <> "" Then
                    'comprobar que todas las letras serie existen en la contabilidad
                    devuelve = "tiporegi= " & DBSet(DevuelveValor(Sql2), "T") 'SQL, "T")
                    RSconta.MoveFirst
                    RSconta.Find (devuelve), , adSearchForward
                    If RSconta.EOF Then
                        'no encontrado
                        b = False
                        'Cad = SQL & " en BD de Contabilidad."
                        Cad = DevuelveValor(Sql2) & " en BD de Contabilidad."
                    End If
                End If
                If b Then Cad = Cad & DBSet(Rs!CodTipom, "T") & ","
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
            RSconta.Close
            Set RSconta = Nothing
            
            If Not b Then 'Hay algun movimiento que no existe
                devuelve = "No existe el tipo de movimiento: " & Cad & vbCrLf
                devuelve = devuelve & "Consulte con el administrador."
                MsgBox devuelve, vbExclamation
                Exit Function
            End If
            
            'Todos los Tipo de movimiento existen
            If Cad <> "" Then
                Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitamos ult. coma
            
                'miramos si hay algun movimiento de factura que la letra serie sea nulo
                sql = "select count(*) from usuarios.stipom "
                sql = sql & "where codtipom IN (" & Cad & ") and (isnull(letraser) or letraser='')"
                If RegistrosAListar(sql) > 0 Then
                    sql = "Hay algun tipo de movimiento de Facturación que no tiene letra serie." & vbCrLf
                    sql = sql & "Comprobar en la tabla de tipos de movimiento: " & Cad
                    MsgBox sql, vbExclamation
                    Exit Function
                End If
            End If
            ComprobarLetraSerie = True
    
    
        Case "rtelmovil"
            'cargamos el RSConta con la tabla contadores de BD: Contabilidad
            'donde estan todas las letra de serie que existen en la contabilidad
            sql = "Select distinct tiporegi from contadores"
            Set RSconta = New ADODB.Recordset
            RSconta.Open sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
            If RSconta.EOF Then
                RSconta.Close
                Set RSconta = Nothing
                Exit Function
            End If
                
        
            'obtenemos las distintas letras de serie de las facturas seleccionadas
            sql = "select distinct rtelmovil.numserie from " & cadTabla
            sql = sql & " INNER JOIN tmpFactu ON rtelmovil.numserie=tmpFactu.numserie AND rtelmovil.numfactu=tmpFactu.numfactu AND rtelmovil.fecfactu=tmpFactu.fecfactu "
            Set Rs = New ADODB.Recordset
            Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad = ""
            b = True
            While Not Rs.EOF And b
                'comprobar que todas las letras serie existen en la contabilidad
                devuelve = "tiporegi= " & DBSet(Rs!numserie, "T") 'SQL, "T")
                RSconta.MoveFirst
                RSconta.Find (devuelve), , adSearchForward
                If RSconta.EOF Then
                    'no encontrado
                    b = False
                    'Cad = SQL & " en BD de Contabilidad."
                    Cad = Rs!numserie & " en BD de Contabilidad."
                End If
                
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
            RSconta.Close
            Set RSconta = Nothing
            
            If Not b Then 'Hay algun movimiento que no existe
                devuelve = "No existe la letra de serie: " & Cad & vbCrLf
                devuelve = devuelve & "Consulte con el administrador."
                MsgBox devuelve, vbExclamation
                Exit Function
            End If
            
            ComprobarLetraSerie = True
    
    
    
        Case "fvarcabfact"
            'cargamos el RSConta con la tabla contadores de BD: Contabilidad
            'donde estan todas las letra de serie que existen en la contabilidad
            sql = "Select distinct tiporegi from contadores"
            Set RSconta = New ADODB.Recordset
            RSconta.Open sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
            If RSconta.EOF Then
                RSconta.Close
                Set RSconta = Nothing
                Exit Function
            End If
                
        
            'obtenemos los distintos tipos de movimiento que vamos a contabilizar
            'de las facturas seleccionadas
            sql = "select distinct fvarcabfact.codtipom from " & cadTabla
            sql = sql & " INNER JOIN tmpFactu ON fvarcabfact.codtipom=tmpFactu.codtipom AND fvarcabfact.numfactu=tmpFactu.numfactu AND fvarcabfact.fecfactu=tmpFactu.fecfactu "
    '        SQL = SQL & cadWHERE
            Set Rs = New ADODB.Recordset
            Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad = ""
            b = True
            While Not Rs.EOF And b
                'comprobar que todas las letras serie existen en usuarios
                sql = "letraser"
                Sql2 = "select letraser from usuarios.stipom where codtipom = " & DBSet(Rs!CodTipom, "T")
                'devuelve = DevuelveDesdeBDNew(cAgro, "stipom", "codtipom", "codtipom", RS!codtipom, "T", SQL)
                total = TotalRegistrosConsulta(Sql2)
                If total = 0 Then 'devuelve = "" Then
                    b = False
                    Cad = Rs!CodTipom & " en BD de Ariagrorec."
                ElseIf DevuelveValor(Sql2) <> "" Then 'SQL <> "" Then
                    'comprobar que todas las letras serie existen en la contabilidad
                    devuelve = "tiporegi= " & DBSet(DevuelveValor(Sql2), "T") 'SQL, "T")
                    RSconta.MoveFirst
                    RSconta.Find (devuelve), , adSearchForward
                    If RSconta.EOF Then
                        'no encontrado
                        b = False
                        'Cad = SQL & " en BD de Contabilidad."
                        Cad = DevuelveValor(Sql2) & " en BD de Contabilidad."
                    End If
                End If
                If b Then Cad = Cad & DBSet(Rs!CodTipom, "T") & ","
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
            RSconta.Close
            Set RSconta = Nothing
            
            If Not b Then 'Hay algun movimiento que no existe
                devuelve = "No existe el tipo de movimiento: " & Cad & vbCrLf
                devuelve = devuelve & "Consulte con el administrador."
                MsgBox devuelve, vbExclamation
                Exit Function
            End If
            
            'Todos los Tipo de movimiento existen
            If Cad <> "" Then
                Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitamos ult. coma
            
                'miramos si hay algun movimiento de factura que la letra serie sea nulo
                sql = "select count(*) from usuarios.stipom "
                sql = sql & "where codtipom IN (" & Cad & ") and (isnull(letraser) or letraser='')"
                If RegistrosAListar(sql) > 0 Then
                    sql = "Hay algun tipo de movimiento de Facturación que no tiene letra serie." & vbCrLf
                    sql = sql & "Comprobar en la tabla de tipos de movimiento: " & Cad
                    MsgBox sql, vbExclamation
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

'###### ESTE YA NO SE UTILIZA
'Public Function ComprobarNumFacturas(cadTabla As String, cadWConta) As Boolean
''Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
''vamos a contabilizar
'Dim SQL As String
'Dim RS As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
'Dim b As Boolean
'
'    On Error GoTo ECompFactu
'
'    ComprobarNumFacturas = False
'
'    SQL = "SELECT numserie,codfaccl,anofaccl FROM cabfact "
'    SQL = SQL & " WHERE " & cadWConta
'
'    Set RSconta = New ADODB.Recordset
'    RSconta.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    If Not RSconta.EOF Then
'        'Seleccionamos las distintas facturas que vamos a facturar
'        SQL = "SELECT DISTINCT " & cadTabla & ".codtipom,letraser,scafac.numfactu,scafac.fecfactu "
'        SQL = SQL & " FROM (" & cadTabla & " INNER JOIN stipom ON " & cadTabla & ".codtipom=stipom.codtipom) "
'        SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
''        SQL = SQL & " WHERE " & cadWHERE
'
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        b = True
'        While Not RS.EOF And b
'            SQL = "(numserie= " & DBSet(RS!LetraSer, "T") & " AND codfaccl=" & DBSet(RS!NumFactu, "N") & " AND anofaccl=" & Year(RS!FecFactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
'                b = False
'                SQL = "          Nº Fac.: " & Format(RS!NumFactu, "0000000") & vbCrLf
'                SQL = SQL & "          Fecha: " & RS!FecFactu
'            End If
'            RS.MoveNext
'        Wend
'        RS.Close
'        Set RS = Nothing
'
'        If Not b Then
'            SQL = "Ya existe la factura: " & vbCrLf & SQL
'            SQL = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & SQL
'
'            MsgBox SQL, vbExclamation
'            ComprobarNumFacturas = False
'        Else
'            ComprobarNumFacturas = True
'        End If
'    Else
'        ComprobarNumFacturas = True
'    End If
'    RSconta.Close
'    Set RSconta = Nothing
'
'ECompFactu:
'     If Err.Number <> 0 Then
'        MuestraError Err.Number, "Comprobar Nº Facturas", Err.Description
'    End If
'End Function


Public Function ComprobarNumFacturas_new(cadTabla As String, cadWConta) As Boolean
'Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
'vamos a contabilizar
Dim sql As String
Dim SQLconta As String
Dim Rs As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECompFactu

    ComprobarNumFacturas_new = False
    
    SQLconta = "SELECT count(*) FROM cabfact WHERE "
 
    
        'Seleccionamos las distintas facturas que vamos a facturar
    If cadTabla = "rtelmovil" Then
        sql = "SELECT DISTINCT " & cadTabla & ".numserie," & cadTabla & ".numfactu," & cadTabla & ".fecfactu "
        sql = sql & " FROM " & cadTabla
        sql = sql & " INNER JOIN tmpFactu ON " & cadTabla & ".numserie=tmpFactu.numserie AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "

        Set Rs = New ADODB.Recordset
        Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF And b
            sql = "(numserie= " & DBSet(Rs!numserie, "T") & " AND codfaccl=" & DBSet(Rs!numfactu, "N") & " AND anofaccl=" & Year(Rs!fecfactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
            sql = SQLconta & sql
            If RegistrosAListar(sql, cConta) Then
                b = False
                sql = "          Letra Serie: " & DBSet(Rs!numserie, "T") & vbCrLf
                sql = sql & "          Nº Fac.: " & Format(Rs!numfactu, "0000000") & vbCrLf
                sql = sql & "          Fecha: " & Rs!fecfactu
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            sql = "Ya existe la factura: " & vbCrLf & sql
            sql = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & sql
            
            MsgBox sql, vbExclamation
            ComprobarNumFacturas_new = False
        Else
            ComprobarNumFacturas_new = True
        End If
     
     
     Else
        sql = "SELECT DISTINCT " & cadTabla & ".codtipom,letraser," & cadTabla & ".numfactu," & cadTabla & ".fecfactu "
        sql = sql & " FROM (" & cadTabla & " INNER JOIN usuarios.stipom ON " & cadTabla & ".codtipom=usuarios.stipom.codtipom) "
        sql = sql & " INNER JOIN tmpFactu ON " & cadTabla & ".codtipom=tmpFactu.codtipom AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "

        
        Set Rs = New ADODB.Recordset
        Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF And b
            sql = "(numserie= " & DBSet(Rs!LetraSer, "T") & " AND codfaccl=" & DBSet(Rs!numfactu, "N") & " AND anofaccl=" & Year(Rs!fecfactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
            sql = SQLconta & sql
            If RegistrosAListar(sql, cConta) Then
                b = False
                sql = "          Letra Serie: " & DBSet(Rs!LetraSer, "T") & vbCrLf
                sql = sql & "          Nº Fac.: " & Format(Rs!numfactu, "0000000") & vbCrLf
                sql = sql & "          Fecha: " & Rs!fecfactu
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            sql = "Ya existe la factura: " & vbCrLf & sql
            sql = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & sql
            
            MsgBox sql, vbExclamation
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




'###### ESTE YA NO SE UTILIZA
'Public Function ComprobarCtaContable(cadTabla As String, Opcion As Byte) As Boolean
''Comprobar que todas las ctas contables de los distintos clientes de las facturas
''que vamos a contabilizar existan en la contabilidad
'Dim SQL As String
'Dim RS As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
'Dim b As Boolean
'Dim cadG As String
'
'    On Error GoTo ECompCta
'
'    ComprobarCtaContable = False
'
'    If Opcion = 3 Then 'si hay analitica comprobar que todas las cuentas
'                        'empiezan por el digito que hay en conta.parametros.grupogto o .grupovta
'        cadG = "grupovta"
'        SQL = DevuelveDesdeBDNew(conConta, "parametros", "grupogto", "", "", "", cadG)
'        If SQL <> "" And cadG <> "" Then
'            SQL = " AND (codmacta like '" & SQL & "%' OR codmacta like '" & cadG & "%')"
'        ElseIf SQL <> "" Then
'            SQL = " AND (codmacta like '" & SQL & "%')"
'        ElseIf cadG <> "" Then
'            SQL = " AND (codmacta like '" & cadG & "%')"
'        End If
'        cadG = SQL
'    End If
'
'    SQL = "SELECT codmacta FROM cuentas "
'    SQL = SQL & " WHERE apudirec='S'"
'    If cadG <> "" Then SQL = SQL & cadG
'
'    Set RSconta = New ADODB.Recordset
'    RSconta.Open SQL, ConnConta, adOpenStatic, adLockPessimistic, adCmdText
'
'    If Not RSconta.EOF Then
'        If Opcion = 1 Then
'            If cadTabla = "scafac" Then
'                'Seleccionamos los distintos clientes,cuentas que vamos a facturar
'                SQL = "SELECT DISTINCT scafac.codclien, sclien.codmacta "
'                SQL = SQL & " FROM (scafac INNER JOIN sclien ON scafac.codclien=sclien.codclien) "
'                SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
'            Else
'                'Seleccionamos los distintos proveedores,cuentas que vamos a facturar
'                SQL = "SELECT DISTINCT scafpc.codprove, sprove.codmacta "
'                SQL = SQL & " FROM (scafpc INNER JOIN sprove ON scafpc.codprove=sprove.codprove) "
'                SQL = SQL & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
'            End If
'
'        ElseIf Opcion = 2 Or Opcion = 3 Then
'            SQL = "SELECT distinct "
'            If Opcion = 2 Then SQL = SQL & " sartic.codfamia,"
'            If cadTabla = "scafac" Then
'                SQL = SQL & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((slifac "
'                SQL = SQL & " INNER JOIN tmpFactu ON slifac.codtipom=tmpFactu.codtipom AND slifac.numfactu=tmpFactu.numfactu AND slifac.fecfactu=tmpFactu.fecfactu) "
'                SQL = SQL & "INNER JOIN sartic ON slifac.codartic=sartic.codartic) "
'            Else
'                SQL = SQL & " sfamia.ctacompr as codmacta,sfamia.abocompr as ctaabono from ((slifpc "
'                SQL = SQL & " INNER JOIN tmpFactu ON slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu) "
'                SQL = SQL & "INNER JOIN sartic ON slifpc.codartic=sartic.codartic) "
'            End If
'            SQL = SQL & " LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia "
'        End If
'
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        b = True
'        While Not RS.EOF And b
'            SQL = "codmacta= " & DBSet(RS!Codmacta, "T")
'            RSconta.MoveFirst
'            RSconta.Find (SQL), , adSearchForward
'            If RSconta.EOF Then
'                b = False 'no encontrado
'                If Opcion = 1 Then
'                    If cadTabla = "scafac" Then
'                        SQL = RS!Codmacta & " del Cliente " & Format(RS!CodClien, "000000")
'                    Else
'                        SQL = RS!Codmacta & " del Proveedor " & Format(RS!codProve, "000000")
'                    End If
'                ElseIf Opcion = 2 Then
'                    SQL = RS!Codmacta & " de la familia " & Format(RS!codfamia, "0000")
'                ElseIf Opcion = 3 Then
'                    SQL = RS!Codmacta
'                End If
'            End If
'
'            If Opcion = 2 Then
'                'Comprobar que ademas de existir la cuenta de ventas exista tambien
'                'la cuenta ABONO ventas
'                SQL = "codmacta= " & DBSet(RS!ctaabono, "T")
'                RSconta.MoveFirst
'                RSconta.Find (SQL), , adSearchForward
'                If RSconta.EOF Then
'                    b = False 'no encontrado
'                    SQL = RS!ctaabono & " de la familia " & Format(RS!codfamia, "0000")
'                End If
'            End If
'
'            'comprobar cuentas alternativas solo para facturacion a clientes
'            If cadTabla = "scafac" Then
'                If Opcion = 2 Then
'                    ' Comprobar cuenta venta alternativa
'                    If DBLet(RS!ctavent1, "T") <> "" Then
'                        SQL = "codmacta= " & DBSet(RS!ctavent1, "T")
'                        RSconta.MoveFirst
'                        RSconta.Find (SQL), , adSearchForward
'                        If RSconta.EOF Then
'                            b = False 'no encontrado
'                            SQL = RS!ctavent1 & " de la familia " & Format(RS!codfamia, "0000")
'                        End If
'                    Else
'                        b = False
'                        SQL = " o la familia no tiene asignada cuenta venta alternativa."
'                    End If
'                End If
'                If Opcion = 2 Then
'                    ' Comprobar cuenta de abono alternativa
'                    If DBLet(RS!abovent1, "T") <> "" Then
'                        SQL = "codmacta= " & DBSet(RS!abovent1, "T")
'                        RSconta.MoveFirst
'                        RSconta.Find (SQL), , adSearchForward
'                        If RSconta.EOF Then
'                            b = False 'no encontrado
'                            SQL = RS!ctaabon1 & " de la familia " & Format(RS!codfamia, "0000")
'                        End If
'                    Else
'                        b = False
'                        SQL = " o la familia no tiene asignada cuenta abono alternativa."
'                    End If
'                End If
'            End If
'            RS.MoveNext
'        Wend
'        RS.Close
'        Set RS = Nothing
'
'        If Not b Then
'            If Opcion <> 3 Then
'                SQL = "No existe la cta contable " & SQL
'            Else
'                SQL = "La cuenta " & SQL & " no es del nivel correcto."
'            End If
'            SQL = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & SQL
'
'            MsgBox SQL, vbExclamation
'            ComprobarCtaContable = False
'        Else
'            ComprobarCtaContable = True
'        End If
'    Else
'        ComprobarCtaContable = True
'    End If
'    RSconta.Close
'    Set RSconta = Nothing
'
'ECompCta:
'     If Err.Number <> 0 Then
'        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
'    End If
'End Function




Public Function ComprobarCtaContable_new(cadTabla As String, Opcion As Byte, Optional Tipo As Byte, Optional Seccion As Integer, Optional cuenta As String) As Boolean
'Comprobar que todas las ctas contables de los distintos clientes de las facturas
'que vamos a contabilizar existan en la contabilidad
Dim sql As String
Dim Rs As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim b As Boolean
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
        sql = DevuelveDesdeBDNew(cConta, "parametros", "grupogto", "", "", "", cadG)
        If sql <> "" And cadG <> "" Then
            sql = " AND (codmacta like '" & sql & "%' OR codmacta like '" & cadG & "%')"
        ElseIf sql <> "" Then
            sql = " AND (codmacta like '" & sql & "%')"
        ElseIf cadG <> "" Then
            sql = " AND (codmacta like '" & cadG & "%')"
        End If
        cadG = sql
    End If


    SQLcuentas = "SELECT count(*) FROM cuentas WHERE apudirec='S' "
    If cadG <> "" Then SQLcuentas = SQLcuentas & cadG

    If Opcion = 1 Then
        If cadTabla = "rfactsoc" Then
            'Seleccionamos los distintos socios, cuentas que vamos a facturar
            sql = "SELECT DISTINCT rfactsoc.codsocio, rsocios_seccion.codmacpro as codmacta "
            sql = sql & " FROM (rfactsoc INNER JOIN rsocios_seccion ON rfactsoc.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N") & ") "
            sql = sql & " INNER JOIN tmpFactu ON rfactsoc.codtipom=tmpFactu.codtipom AND rfactsoc.numfactu=tmpFactu.numfactu AND rfactsoc.fecfactu=tmpFactu.fecfactu "
        Else
            If cadTabla = "rcabfactalmz" Then
                'Seleccionamos los distintos socios, cuentas que vamos a facturar
                sql = "SELECT DISTINCT rcabfactalmz.codsocio, rsocios_seccion.codmacpro as codmacta "
                sql = sql & " FROM (rcabfactalmz INNER JOIN rsocios_seccion ON rcabfactalmz.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N") & ") "
                sql = sql & " INNER JOIN tmpFactu ON rcabfactalmz.tipofichero=tmpFactu.tipofichero AND rcabfactalmz.numfactu=tmpFactu.numfactu AND rcabfactalmz.fecfactu=tmpFactu.fecfactu "
                sql = sql & " and rcabfactalmz.codsocio = tmpFactu.codsocio "
                
                '[Monica]29/07/2015: en el caso de catadau si es asociado tengo que mirar raiz asociado + codigo asociado
                '                                           si es socio entonces raiz socio + codigo socio
                If vParamAplic.Cooperativa = 0 Then
                    sql = "SELECT DISTINCT rcabfactalmz.codsocio, if(rsocios.tiporelacion = 1, concat(rseccion.raiz_cliente_asociado,right(concat('00000',rsocios.nroasociado),5)), rsocios_seccion.codmacpro) as codmacta "
                    sql = sql & " FROM (((rcabfactalmz INNER JOIN rsocios_seccion ON rcabfactalmz.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N") & ") "
                    sql = sql & " INNER JOIN tmpFactu ON rcabfactalmz.tipofichero=tmpFactu.tipofichero AND rcabfactalmz.numfactu=tmpFactu.numfactu AND rcabfactalmz.fecfactu=tmpFactu.fecfactu "
                    sql = sql & " and rcabfactalmz.codsocio = tmpFactu.codsocio) "
                    sql = sql & " INNER JOIN rsocios ON rcabfactalmz.codsocio = rsocios.codsocio) "
                    sql = sql & " INNER JOIN rseccion on rseccion.codsecci = rsocios_seccion.codsecci "
                
                
                End If
                
                
            Else
                If cadTabla = "advfacturas" Then
                    'Seleccionamos los distintos socios, cuentas que vamos a facturar
                    sql = "SELECT DISTINCT advfacturas.codsocio, rsocios_seccion.codmaccli as codmacta "
                    sql = sql & " FROM (advfacturas INNER JOIN rsocios_seccion ON advfacturas.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionADV, "N") & ") "
                    sql = sql & " INNER JOIN tmpFactu ON advfacturas.codtipom=tmpFactu.codtipom AND advfacturas.numfactu=tmpFactu.numfactu AND advfacturas.fecfactu=tmpFactu.fecfactu "
'                    SQL = SQL & " and advfacturas.codsocio = tmpFactu.codsocio "
                Else ' facturas de retirada de almazara
                    If cadTabla = "rbodfact1" Then
                        'Seleccionamos los distintos socios, cuentas que vamos a facturar
                        sql = "SELECT DISTINCT rbodfacturas.codsocio, rsocios_seccion.codmaccli as codmacta "
                        sql = sql & " FROM (rbodfacturas INNER JOIN rsocios_seccion ON rbodfacturas.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N") & ") "
                        sql = sql & " INNER JOIN tmpFactu ON rbodfacturas.codtipom=tmpFactu.codtipom AND rbodfacturas.numfactu=tmpFactu.numfactu AND rbodfacturas.fecfactu=tmpFactu.fecfactu "
                    Else ' facturas de retirada de bodega
                        If cadTabla = "rbodfact2" Then
                            'Seleccionamos los distintos socios, cuentas que vamos a facturar
                            sql = "SELECT DISTINCT rbodfacturas.codsocio, rsocios_seccion.codmaccli as codmacta "
                            sql = sql & " FROM (rbodfacturas INNER JOIN rsocios_seccion ON rbodfacturas.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionBodega, "N") & ") "
                            sql = sql & " INNER JOIN tmpFactu ON rbodfacturas.codtipom=tmpFactu.codtipom AND rbodfacturas.numfactu=tmpFactu.numfactu AND rbodfacturas.fecfactu=tmpFactu.fecfactu "
                        Else 'facturas de telefonia
                            If cadTabla = "rtelmovil" Then
                                sql = "SELECT DISTINCT rtelmovil.codsocio, rsocios_seccion.codmaccli as codmacta "
                                sql = sql & " FROM (rtelmovil INNER JOIN rsocios_seccion ON rtelmovil.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.Seccionhorto, "N") & ") "
                                sql = sql & " INNER JOIN tmpFactu ON rtelmovil.numserie=tmpFactu.numserie AND rtelmovil.numfactu=tmpFactu.numfactu AND rtelmovil.fecfactu=tmpFactu.fecfactu "
                            Else
                                If cadTabla = "rrecibpozos" Then
                                    sql = "SELECT DISTINCT rrecibpozos.codsocio, rsocios_seccion.codmaccli as codmacta "
                                    sql = sql & " FROM (rrecibpozos INNER JOIN rsocios_seccion ON rrecibpozos.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionPOZOS, "N") & ") "
                                    sql = sql & " INNER JOIN tmpFactu ON rrecibpozos.codtipom=tmpFactu.codtipom AND rrecibpozos.numfactu=tmpFactu.numfactu AND rrecibpozos.fecfactu=tmpFactu.fecfactu "
                                Else
                                    If cadTabla = "rfacttra" Then
                                        'Seleccionamos los distintos socios, cuentas que vamos a facturar
                                        sql = "SELECT DISTINCT rfacttra.codtrans, rtransporte.codmacpro as codmacta "
                                        sql = sql & " FROM (rfacttra INNER JOIN rtransporte ON rfacttra.codtrans=rtransporte.codtrans) "
                                        sql = sql & " INNER JOIN tmpFactu ON rfacttra.codtipom=tmpFactu.codtipom AND rfacttra.numfactu=tmpFactu.numfactu AND rfacttra.fecfactu=tmpFactu.fecfactu "
                                    Else
                                        If cadTabla = "fvarcabfact" Then
                                            If Tipo = 0 Then ' seleccionamos primero los socios
                                                'Seleccionamos los distintos socios de facturas varias, cuentas que vamos a facturar
                                                sql = "SELECT DISTINCT fvarcabfact.codsocio, rsocios_seccion.codmaccli as codmacta "
                                                sql = sql & " FROM (fvarcabfact INNER JOIN rsocios_seccion ON fvarcabfact.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N") & " and not fvarcabfact.codsocio is null and fvarcabfact.codsocio <> 0 ) "
                                                sql = sql & " INNER JOIN tmpFactu ON fvarcabfact.codtipom=tmpFactu.codtipom AND fvarcabfact.numfactu=tmpFactu.numfactu AND fvarcabfact.fecfactu=tmpFactu.fecfactu "
                                            Else
                                                'Seleccionamos los distintos clientes de facturas varias, cuentas que vamos a facturar
                                                sql = "SELECT DISTINCT fvarcabfact.codclien, clientes.codmacta as codmacta "
                                                sql = sql & " FROM (fvarcabfact INNER JOIN clientes ON fvarcabfact.codclien=clientes.codclien and not fvarcabfact.codclien is null and fvarcabfact.codclien <> 0) "
                                                sql = sql & " INNER JOIN tmpFactu ON fvarcabfact.codtipom=tmpFactu.codtipom AND fvarcabfact.numfactu=tmpFactu.numfactu AND fvarcabfact.fecfactu=tmpFactu.fecfactu "
                                            End If
                                        Else
                                            If cadTabla = "fvarcabfactpro" Then
                                                'Seleccionamos los distintos socios de facturas varias, cuentas que vamos a facturar
                                                sql = "SELECT DISTINCT fvarcabfactpro.codsocio, rsocios_seccion.codmacpro as codmacta "
                                                sql = sql & " FROM (fvarcabfactpro INNER JOIN rsocios_seccion ON fvarcabfactpro.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N") & " ) "
                                                sql = sql & " INNER JOIN tmpFactu ON fvarcabfactpro.codtipom=tmpFactu.codtipom AND fvarcabfactpro.numfactu=tmpFactu.numfactu AND fvarcabfactpro.fecfactu=tmpFactu.fecfactu "
                                            Else
                                                'Seleccionamos los distintos socios terceros, cuentas que vamos a facturar
                                                sql = "SELECT DISTINCT rcafter.codsocio, rsocios_seccion.codmacpro as codmacta "
                                                sql = sql & " FROM (rcafter INNER JOIN rsocios_seccion ON rcafter.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & vParamAplic.Seccionhorto & ") "
                                                sql = sql & " INNER JOIN tmpFactu ON rcafter.codsocio=tmpFactu.codsocio AND rcafter.numfactu=tmpFactu.numfactu AND rcafter.fecfactu=tmpFactu.fecfactu "
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
        sql = "SELECT distinct "
        If cadTabla = "rfactsoc" Then
            Select Case Tipo
                Case 1, 3, 7, 9  ' 1=anticipos , 3=anticipos venta campo, 7=anticipos almazara, 9=anticipo bodega
                    sql = sql & " rfactsoc_variedad.codvarie, variedades.ctaanticipo as codmacta from ((rfactsoc_variedad "
                Case 2, 4, 8, 10, 11 ' 2=liquidaciones, 4=liquidaciones venta campo, 8=liquidacion almazara, 10=liquidacion bodega
                    sql = sql & " rfactsoc_variedad.codvarie, variedades.ctaliquidacion as codmacta from ((rfactsoc_variedad "
                Case 6  '6=siniestros
                    sql = sql & " rfactsoc_variedad.codvarie, variedades.ctasiniestros as codmacta from ((rfactsoc_variedad "
                Case 12 ' le paso un tipo 12 a las liquidaciones de industria de terceros para comprobar las cuentas de terceros
                    sql = sql & " rfactsoc_variedad.codvarie, variedades.ctacomtercero as codmacta from ((rfactsoc_variedad "
                Case 13 ' facturas de acarreo recoleccion socio FTS
                    sql = sql & " rfactsoc_variedad.codvarie, variedades.ctaacarecol as codmacta from ((rfactsoc_variedad "
            End Select
            sql = sql & " INNER JOIN tmpFactu ON rfactsoc_variedad.codtipom=tmpFactu.codtipom AND rfactsoc_variedad.numfactu=tmpFactu.numfactu AND rfactsoc_variedad.fecfactu=tmpFactu.fecfactu) "
            sql = sql & "INNER JOIN variedades ON rfactsoc_variedad.codvarie=variedades.codvarie) "
        Else
            If cadTabla = "rfacttra" Then
                sql = sql & " rfacttra_albaran.codvarie, variedades.ctatransporte as codmacta from ((rfacttra_albaran "
                sql = sql & " INNER JOIN tmpFactu ON rfacttra_albaran.codtipom=tmpFactu.codtipom AND rfacttra_albaran.numfactu=tmpFactu.numfactu AND rfacttra_albaran.fecfactu=tmpFactu.fecfactu) "
                sql = sql & "INNER JOIN variedades ON rfacttra_albaran.codvarie=variedades.codvarie) "
            Else
                sql = sql & " rlifter.codvarie, variedades.ctacomtercero as codmacta from ((rlifter "
                sql = sql & " INNER JOIN tmpFactu ON rlifter.codsocio=tmpFactu.codsocio AND rlifter.numfactu=tmpFactu.numfactu AND rlifter.fecfactu=tmpFactu.fecfactu) "
                sql = sql & "INNER JOIN variedades ON rlifter.codvarie=variedades.codvarie) "
            End If
        End If
    ElseIf Opcion = 2 Then
            If cadTabla = "advfacturas" Then
                sql = "SELECT distinct advartic.codfamia,"
                sql = sql & " advfamia.ctaventa as codmacta,advfamia.aboventa as ctaabono from ((advfacturas_lineas "
                sql = sql & " INNER JOIN tmpFactu ON advfacturas_lineas.codtipom=tmpFactu.codtipom AND advfacturas_lineas.numfactu=tmpFactu.numfactu AND advfacturas_lineas.fecfactu=tmpFactu.fecfactu) "
                sql = sql & "INNER JOIN advartic ON advfacturas_lineas.codartic=advartic.codartic) "
                sql = sql & "INNER JOIN advfamia ON advartic.codfamia = advfamia.codfamia "
            Else
                If cadTabla = "rbodfacturas" Then
                    sql = "SELECT distinct rbodfacturas_lineas.codvarie, variedades.ctavtasotros as codmacta from (rbodfacturas_lineas "
                    sql = sql & " INNER JOIN tmpFactu ON rbodfacturas_lineas.codtipom=tmpFactu.codtipom AND rbodfacturas_lineas.numfactu=tmpFactu.numfactu AND rbodfacturas_lineas.fecfactu=tmpFactu.fecfactu) "
                    sql = sql & " INNER JOIN variedades ON rbodfacturas_lineas.codvarie = variedades.codvarie "
                Else
                    If cadTabla = "rbodfact1" Then ' rbodfacturas de almazara "FZA"
                        sql = "select distinct " & vParamAplic.CtaVentasAlmz & " as codmacta "
                        sql = sql & " FROM rbodfacturas "
                    Else
                        If cadTabla = "rbodfact2" Then ' rbodfacturas de bodega "FAB"
                            sql = "select distinct " & vParamAplic.CtaVentasBOD & " as codmacta "
                            sql = sql & " FROM rbodfacturas "
                        Else
                            If cadTabla = "rtelmovil" Then
                                sql = "select distinct " & vParamAplic.CtaVentasTel & " as codmacta "
                                sql = sql & " FROM rtelmovil "
                            Else
                                If cadTabla = "rrecibpozos" Then
                                    Select Case Tipo
                                        Case 1   ' cuenta de ventas de consumo
                                            sql = "select distinct " & vParamAplic.CtaVentasConsPOZ & " as codmacta "
                                            sql = sql & " FROM rrecibpozos "
                                        Case 2   ' cuenta de ventas de cuotas
                                            sql = "select distinct " & vParamAplic.CtaVentasCuoPOZ & " as codmacta "
                                            sql = sql & " FROM rrecibpozos "
                                        Case 3   ' cuenta de ventas de talla
                                            sql = "select distinct " & vParamAplic.CtaVentasTalPOZ & " as codmacta "
                                            sql = sql & " FROM rrecibpozos "
                                        Case 4   ' cuenta de ventas de mantenimiento
                                            sql = "select distinct " & vParamAplic.CtaVentasMtoPOZ & " as codmacta "
                                            sql = sql & " FROM rrecibpozos "
                                        Case 5   ' cuenta de ventas de consumo a manta
                                            sql = "select distinct " & vParamAplic.CtaVentasMantaPOZ & " as codmacta "
                                            sql = sql & " FROM rrecibpozos "
                                    End Select
                                Else
                                    If cadTabla = "fvarcabfact" Then
                                        sql = "select distinct fvarconce.codmacta as codmacta "
                                        sql = sql & " FROM ((fvarlinfact "
                                        sql = sql & " INNER JOIN tmpFactu ON fvarlinfact.codtipom=tmpFactu.codtipom AND fvarlinfact.numfactu=tmpFactu.numfactu AND fvarlinfact.fecfactu=tmpFactu.fecfactu) "
                                        sql = sql & "INNER JOIN fvarconce ON fvarlinfact.codconce=fvarconce.codconce) "
                                    Else
                                        If cadTabla = "fvarcabfactpro" Then
                                            sql = "select distinct fvarconce.codmacpr as codmacta "
                                            sql = sql & " FROM ((fvarlinfactpro "
                                            sql = sql & " INNER JOIN tmpFactu ON fvarlinfactpro.codtipom=tmpFactu.codtipom AND fvarlinfactpro.numfactu=tmpFactu.numfactu AND fvarlinfactpro.fecfactu=tmpFactu.fecfactu) "
                                            sql = sql & "INNER JOIN fvarconce ON fvarlinfactpro.codconce=fvarconce.codconce) "
                                        Else
                                            sql = "select distinct " & vParamAplic.CtaVentasAlmz & " as codmacta "
                                            sql = sql & " FROM rcabfactalmz "
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
    ElseIf Opcion = 3 Then
            sql = "select distinct " & vParamAplic.CtaGastosAlmz & " as codmacta "
            sql = sql & " FROM rcabfactalmz "
    ElseIf Opcion = 4 Then
        sql = "select distinct " & DBSet(vParamAplic.CtaTerReten, "T") & " as codmacta from tcafpc "
    ElseIf Opcion = 7 Then
            If cadTabla = "rfactsoc" Then
                Select Case Tipo
                    Case 1, 3, 7, 9  ' 1=anticipos , 3=anticipos venta campo, 7=anticipos almazara, 9=anticipos bodega
                        sql = " SELECT variedades.ctaanticipo as cuenta "
                        sql = sql & " FROM rfactsoc_variedad, variedades, tmpFactu  WHERE "
                        sql = sql & " rfactsoc_variedad.codtipom=tmpFactu.codtipom and rfactsoc_variedad.numfactu=tmpFactu.numfactu and rfactsoc_variedad.fecfactu=tmpFactu.fecfactu and "
                        sql = sql & " rfactsoc_variedad.codvarie=variedades.codvarie "
                        sql = sql & " group by 1 "
                    Case 2, 4, 6, 8, 10 ' 2=liquidaciones, 4=liquidaciones venta campo, 6=siniestros, 8=liquidacion almazara, 10=liquidacion bodega
                        sql = " SELECT variedades.ctaliquidacion as cuenta "
                        sql = sql & " FROM rfactsoc_variedad, variedades, tmpFactu  WHERE "
                        sql = sql & " rfactsoc_variedad.codtipom=tmpFactu.codtipom and rfactsoc_variedad.numfactu=tmpFactu.numfactu and rfactsoc_variedad.fecfactu=tmpFactu.fecfactu and "
                        sql = sql & " rfactsoc_variedad.codvarie=variedades.codvarie "
                        sql = sql & " group by 1 "
                End Select
            Else
                If cadTabla = "advfacturas" Then
                    sql = "SELECT distinct advartic.codfamia,"
                    sql = sql & " advfamia.ctaventa as cuenta,advfamia.aboventa as ctaabono from ((advfacturas_lineas "
                    sql = sql & " INNER JOIN tmpFactu ON advfacturas_lineas.codtipom=tmpFactu.codtipom AND advfacturas_lineas.numfactu=tmpFactu.numfactu AND advfacturas_lineas.fecfactu=tmpFactu.fecfactu) "
                    sql = sql & "INNER JOIN advartic ON advfacturas_lineas.codartic=advartic.codartic) "
                    sql = sql & "INNER JOIN advfamia ON advartic.codfamia = advfamia.codfamia "
                Else
                    If cadTabla = "rbodfacturas" Then
                        sql = "SELECT distinct "
                        sql = sql & " variedades.ctavtasotros as cuenta from (rbodfacturas_lineas "
                        sql = sql & " INNER JOIN tmpFactu ON rbodfacturas_lineas.codtipom=tmpFactu.codtipom AND rbodfacturas_lineas.numfactu=tmpFactu.numfactu AND rbodfacturas_lineas.fecfactu=tmpFactu.fecfactu) "
                        sql = sql & "INNER JOIN variedades ON rbodfacturas_lineas.codvarie=variedades.codvarie "
                    Else
                        If cadTabla = "rbodfact1" Then ' facturas de retirada de almazara
                            sql = "select distinct " & vParamAplic.CtaVentasAlmz & " as cuenta "
                            sql = sql & " FROM rbodfacturas "
                        Else
                            If cadTabla = "rbodfact2" Then ' facturas de retirada de bodega
                                sql = "select distinct " & vParamAplic.CtaVentasBOD & " as cuenta "
                                sql = sql & " FROM rbodfacturas "
                            Else
                                If cadTabla = "rtelmovil" Then
                                    sql = "select distinct " & vParamAplic.CtaVentasTel & " as cuenta "
                                    sql = sql & " FROM rtelmovil "
                                Else
                                    If cadTabla = "rrecibpozos" Then
                                        sql = "select distinct " & vParamAplic.CtaVentasConsPOZ & " as cuenta "
                                        sql = sql & " FROM rrecibpozos "
                                    Else
                                        If cadTabla = "rfacttra" Then
                                            sql = " SELECT variedades.ctatransporte as cuenta "
                                            sql = sql & " FROM rfacttra_albaran, variedades, tmpFactu  WHERE "
                                            sql = sql & " rfacttra_albaran.codtipom=tmpFactu.codtipom and rfacttra_albaran.numfactu=tmpFactu.numfactu and rfacttra_albaran.fecfactu=tmpFactu.fecfactu and "
                                            sql = sql & " rfacttra_albaran.codvarie=variedades.codvarie "
                                            sql = sql & " group by 1 "
                                        Else
                                            If cadTabla = "fvarcabfact" Then
                                                sql = "SELECT distinct "
                                                sql = sql & " fvarconce.codmacta as cuenta from (fvarlinfact "
                                                sql = sql & " INNER JOIN tmpFactu ON fvarlinfact.codtipom=tmpFactu.codtipom AND fvarlinfact.numfactu=tmpFactu.numfactu AND fvarlinfact.fecfactu=tmpFactu.fecfactu) "
                                                sql = sql & "INNER JOIN fvarconce ON fvarlinfact.codconce=fvarconce.codconce "
                                            Else
                                                If cadTabla = "fvarcabfactpro" Then
                                                    sql = "SELECT distinct "
                                                    sql = sql & " fvarconce.codmacpr as cuenta from (fvarlinfactpro "
                                                    sql = sql & " INNER JOIN tmpFactu ON fvarlinfactpro.codtipom=tmpFactu.codtipom AND fvarlinfactpro.numfactu=tmpFactu.numfactu AND fvarlinfactpro.fecfactu=tmpFactu.fecfactu) "
                                                    sql = sql & "INNER JOIN fvarconce ON fvarlinfactpro.codconce=fvarconce.codconce "
                                                Else
                                            '       terceros
                                                    sql = " SELECT variedades.ctacomtercero as cuenta "
                                                    sql = sql & " FROM rlifter, variedades, tmpFactu  WHERE "
                                                    sql = sql & " rlifter.codsocio=tmpFactu.codsocio and rlifter.numfactu=tmpFactu.numfactu and rlifter.fecfactu=tmpFactu.fecfactu and "
                                                    sql = sql & " rlifter.codvarie=variedades.codvarie "
                                                    sql = sql & " group by 1 "
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
            sql = " select distinct " & vParamAplic.CtaVentasAlmz & " as cuenta "
            sql = sql & " from tmpFactu "
    ElseIf Opcion = 11 Then
            sql = " select distinct " & vParamAplic.CtaGastosAlmz & " as cuenta "
            sql = sql & " from tmpFactu "
    ElseIf Opcion = 12 Then
            sql = "SELECT rconcepgasto.codmacta as cuenta "
            sql = sql & " from (rconcepgasto INNER JOIN rfactsoc_gastos  ON rconcepgasto.codgasto = rfactsoc_gastos.codgasto) "
            sql = sql & " INNER JOIN tmpFactu ON rfactsoc_gastos.codtipom=tmpFactu.codtipom AND rfactsoc_gastos.numfactu=tmpFactu.numfactu AND rfactsoc_gastos.fecfactu=tmpFactu.fecfactu "
    ElseIf Opcion = 13 Then
        sql = "SELECT distinct "
        sql = sql & " rcafter.concepcargo, fvarconce.codmacpr as codmacta from ((rcafter "
        sql = sql & " INNER JOIN tmpFactu ON rcafter.codsocio=tmpFactu.codsocio AND rcafter.numfactu=tmpFactu.numfactu AND rcafter.fecfactu=tmpFactu.fecfactu) "
        sql = sql & "INNER JOIN fvarconce ON rcafter.concepcargo=fvarconce.codconce) "
    ElseIf Opcion = 14 Then
        'Seleccionamos los distintos socios asociados , cuentas que vamos a facturar
        sql = "SELECT DISTINCT rfactsoc.codsocio, replace(codmacpro,mid(codmacpro,1,length(rseccion.raiz_cliente_socio)), rseccion.raiz_cliente_asociado) as codmacta "
        sql = sql & " FROM (((rfactsoc INNER JOIN rsocios_seccion ON rfactsoc.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N") & ") INNER JOIN rseccion ON rsocios_seccion.codsecci = rseccion.codsecci) INNER JOIN rsocios ON rsocios.codsocio = rsocios_seccion.codsocio and rsocios.tiporelacion = 1) "
        sql = sql & " INNER JOIN tmpFactu ON rfactsoc.codtipom=tmpFactu.codtipom AND rfactsoc.numfactu=tmpFactu.numfactu AND rfactsoc.fecfactu=tmpFactu.fecfactu "
    
    End If

    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    sql = ""
    b = True

    While Not Rs.EOF And b
        If Opcion < 4 Or Opcion = 8 Or Opcion = 13 Or Opcion = 14 Then
            sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!Codmacta, "T")
        ElseIf Opcion = 4 Then
            sql = SQLcuentas & " AND codmacta= " & DBSet(vParamAplic.CtaTerReten, "T")
        ElseIf Opcion = 7 Then
            sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!cuenta, "T")
        ElseIf Opcion = 9 Or Opcion = 10 Or Opcion = 11 Or Opcion = 12 Then
            sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!cuenta, "T")
        End If


        If Not (RegistrosAListar(sql, cConta) > 0) Then
        'si no lo encuentra
            b = False 'no encontrado
            If Opcion = 1 Then
                If cadTabla = "facturas" Then
                    sql = DBLet(Rs!Codmacta, "T") & " del Socio " & Format(Rs!CodClien, "000000")
                Else
                    If cadTabla = "rfacttra" Then
                        sql = DBLet(Rs!Codmacta, "T") & " del transportista " & DBLet(Rs!codTrans, "T")
                    Else
                        If cadTabla = "rfactsoc" Or cadTabla = "advfacturas" Or cadTabla = "rbodfact1" Or cadTabla = "rbodfact2" Or cadTabla = "rtelmovil" Or cadTabla = "rrecibpozos" Or cadTabla = "fvarcabfact" Or cadTabla = "fvarcabfactpro" Then
                            sql = DBLet(Rs!Codmacta, "T") & " del Socio " & Format(Rs!Codsocio, "000000")
                        Else
                            sql = DBLet(Rs!Codmacta, "T") & " del Socio " & Format(Rs!Codsocio, "000000")
                        End If
                    End If
                End If
            ElseIf Opcion = 2 Then
                If cadTabla = "advfacturas" Then
                    sql = DBLet(Rs!Codmacta, "T") & " de la familia " & DBLet(Rs!codfamia, "N")
                Else
                    If cadTabla = "rbodfacturas" Then
                        sql = DBLet(Rs!Codmacta, "T") & " de la variedad " & DBLet(Rs!codvarie, "N")
                    Else
                        If cadTabla = "rbodfact1" Then
                            sql = DBLet(Rs!Codmacta, "T") & " de ventas de Almazara"
                        Else
                            If cadTabla = "rbodfact2" Then
                                sql = DBLet(Rs!Codmacta, "T") & " de ventas de Bodega"
                            Else
                                If cadTabla = "rrecibpozos" Then
                                    Select Case Tipo
                                        Case 1
                                            sql = DBLet(Rs!Codmacta, "T") & " de ventas consumo de Pozos"
                                        Case 2
                                            sql = DBLet(Rs!Codmacta, "T") & " de ventas cuotas de Pozos"
                                        Case 3
                                            sql = DBLet(Rs!Codmacta, "T") & " de ventas talla de Pozos"
                                        Case 4
                                            sql = DBLet(Rs!Codmacta, "T") & " de ventas mantenimiento de Pozos"
                                        Case 5
                                            sql = DBLet(Rs!Codmacta, "T") & " de vevntas consumo a manta Pozos"
                                    End Select
                                Else
                                    If cadTabla = "fvarcabfact" Then
                                        sql = DBLet(Rs!Codmacta, "T") & " del concepto de factura varia cliente"
                                    Else
                                        If cadTabla = "fvarcabfactpro" Then
                                            sql = DBLet(Rs!Codmacta, "T") & " del concepto de factura varia proveedor"
                                        Else
                                            If cadTabla = "rtelmovil" Then
                                                sql = DBLet(Rs!Codmacta, "T") & " de ventas de Telefonia"
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            ElseIf Opcion = 4 Then
                sql = vParamAplic.CtaTerReten
            ElseIf Opcion = 7 Then
                sql = DBLet(Rs!cuenta, "T")
            ElseIf Opcion = 8 Then
                sql = DBLet(Rs!Codmacta, "T") & " de la variedad " & Format(Rs!codvarie, "0000")
            ElseIf Opcion = 9 Then
                sql = DBLet(Rs!cuenta, "T") & " de ventas de almazara "
            ElseIf Opcion = 11 Then
                sql = DBLet(Rs!cuenta, "T") & " de gastos de almazara "
            ElseIf Opcion = 12 Then
                sql = DBLet(Rs!cuenta, "T") & " de gasto de concepto a pie de factura "
            ElseIf Opcion = 13 Then
                sql = DBLet(Rs!Codmacta, "T") & " del concepto de gasto "
            ElseIf Opcion = 14 Then
                sql = DBLet(Rs!Codmacta, "T") & " del Socio Asociado " & Format(Rs!Codsocio, "000000")
            End If
        End If

        If b And (Opcion = 2 Or Opcion = 7) Then
            If cadTabla = "advfacturas" Then
                'Comprobar que ademas de existir la cuenta de ventas exista tambien
                'la cuenta ABONO ventas (sfamia.aboventa)
                '---------------------------------------------
                sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!ctaabono, "T")
    '            RSconta.MoveFirst
    '            RSconta.Find (SQL), , adSearchForward
    '            If RSconta.EOF Then
                If Not (RegistrosAListar(sql, cConta) > 0) Then
                    b = False 'no encontrado
                    If Opcion = 2 Then
                        sql = DBLet(Rs!ctaabono, "T") & " de la familia " & Format(Rs!codfamia, "0000")
                    ElseIf Opcion = 7 Then
                        sql = DBLet(Rs!ctaabono, "T")
                    End If
                End If
            End If
        End If

        Rs.MoveNext
    Wend

    If Not b Then
        If Not (Opcion = 7 Or Opcion = 9 Or Opcion = 10 Or Opcion = 11 Or Opcion = 12) Then
            sql = "No existe la cta contable " & sql
            If Opcion = 4 Then sql = sql & " de retención."
        Else
            sql = "La cuenta " & sql & " no es del nivel correcto. "
        End If
        sql = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & sql

        MsgBox sql, vbExclamation
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
Dim sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim I As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVA = False
    
    sql = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open sql, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        For I = 1 To 3
            If cadTabla = "advfacturas" Then
                sql = "SELECT DISTINCT advfacturas.codiiva" & I
                sql = sql & " FROM advfacturas "
                sql = sql & " INNER JOIN tmpFactu ON advfacturas.codtipom=tmpFactu.codtipom AND advfacturas.numfactu=tmpFactu.numfactu AND advfacturas.fecfactu=tmpFactu.fecfactu "
                sql = sql & " WHERE not isnull(codiiva" & I & ")"
'                SQL = SQL & " WHERE " & " codigiv" & i & " <> 0 "
            Else
                If cadTabla = "rbodfacturas" Then
                    sql = "SELECT DISTINCT rbodfacturas.codiiva" & I
                    sql = sql & " FROM rbodfacturas "
                    sql = sql & " INNER JOIN tmpFactu ON rbodfacturas.codtipom=tmpFactu.codtipom AND rbodfacturas.numfactu=tmpFactu.numfactu AND rbodfacturas.fecfactu=tmpFactu.fecfactu "
                    sql = sql & " WHERE not isnull(codiiva" & I & ")"
                Else
                    If cadTabla = "scafpc" Then
                        sql = "SELECT DISTINCT scafpc.tipoiva" & I
                        sql = sql & " FROM " & cadTabla
                        sql = sql & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
                        sql = sql & " WHERE not isnull(tipoiva" & I & ")"
        '                SQL = SQL & " WHERE " & " tipoiva" & i & " <> 0 "
                    Else
                        If cadTabla = "rrecibpozos" Then
                            sql = "SELECT DISTINCT tipoiva"
                            sql = sql & " FROM " & cadTabla
                            sql = sql & " INNER JOIN tmpFactu ON rrecibpozos.codtipom=tmpFactu.codtipom AND rrecibpozos.numfactu=tmpFactu.numfactu AND rrecibpozos.fecfactu=tmpFactu.fecfactu "
                            sql = sql & " WHERE not isnull(tipoiva)"
            '                SQL = SQL & " WHERE " & " tipoiva" & i & " <> 0 "
                        Else
                            sql = "SELECT DISTINCT rcafter.tipoiva" & I
                            sql = sql & " FROM " & cadTabla
                            sql = sql & " INNER JOIN tmpFactu ON rcafter.codsocio=tmpFactu.codsocio AND rcafter.numfactu=tmpFactu.numfactu AND rcafter.fecfactu=tmpFactu.fecfactu "
                            sql = sql & " WHERE not isnull(tipoiva" & I & ")"
            '                SQL = SQL & " WHERE " & " tipoiva" & i & " <> 0 "
                    
                        End If
                    End If
               End If
            End If
'            SQL = SQL & " WHERE " & cadWHERE & " AND codigiv" & i & " <> 0 "

            Set Rs = New ADODB.Recordset
            Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            b = True
            While Not Rs.EOF And b
                sql = "codigiva= " & DBSet(Rs.Fields(0), "N")
                RSconta.MoveFirst
                RSconta.Find (sql), , adSearchForward
                If RSconta.EOF Then
                    b = False 'no encontrado
                    sql = "Tipo de IVA: " & Rs.Fields(0)
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        
            If Not b Then
                sql = "No existe el " & sql
                sql = "Comprobando Tipos de IVA en contabilidad..." & vbCrLf & vbCrLf & sql
            
                MsgBox sql, vbExclamation
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
Dim sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim I As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarIVA = False
    
    sql = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open sql, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
            Select Case cadTabla
                Case "rfactsoc"
                    sql = "SELECT DISTINCT rfactsoc.tipoiva"
                    sql = sql & " FROM rfactsoc "
                    sql = sql & " INNER JOIN tmpFactu ON rfactsoc.codtipom=tmpFactu.codtipom AND rfactsoc.numfactu=tmpFactu.numfactu AND rfactsoc.fecfactu=tmpFactu.fecfactu "
                    sql = sql & " WHERE not isnull(tipoiva)"
                Case "rcabfactalmz"
                    sql = "SELECT DISTINCT rcabfactalmz.tipoiva"
                    sql = sql & " FROM rcabfactalmz "
                    sql = sql & " INNER JOIN tmpFactu ON rcabfactalmz.tipofichero=tmpFactu.tipofichero AND rcabfactalmz.numfactu=tmpFactu.numfactu AND rcabfactalmz.fecfactu=tmpFactu.fecfactu "
                    sql = sql & " and rcabfactalmz.codsocio = tmpFactu.codsocio "
                    sql = sql & " WHERE not isnull(tipoiva)"
                Case "rtelmovil"
                    sql = "SELECT DISTINCT " & CodIva
                    sql = sql & " FROM rtelmovil  "
                Case "rfacttra"
                    sql = "SELECT DISTINCT rfacttra.tipoiva"
                    sql = sql & " FROM rfacttra "
                    sql = sql & " INNER JOIN tmpFactu ON rfacttra.codtipom=tmpFactu.codtipom AND rfacttra.numfactu=tmpFactu.numfactu AND rfacttra.fecfactu=tmpFactu.fecfactu "
                    sql = sql & " WHERE not isnull(tipoiva)"
                
            End Select

            Set Rs = New ADODB.Recordset
            Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            b = True
            While Not Rs.EOF And b
                sql = "codigiva= " & DBSet(Rs.Fields(0), "N")
                RSconta.MoveFirst
                RSconta.Find (sql), , adSearchForward
                If RSconta.EOF Then
                    b = False 'no encontrado
                    sql = "Tipo de IVA: " & Rs.Fields(0)
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        
            If Not b Then
                sql = "No existe el " & sql
                sql = "Comprobando Tipos de IVA en contabilidad..." & vbCrLf & vbCrLf & sql
            
                MsgBox sql, vbExclamation
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
Dim sql As String

    On Error GoTo ECCoste

    ComprobarCCoste = False
    sql = vUsu.Login
    If sql <> "" Then
        cadCC = DevuelveDesdeBDNew(cAgro, "straba", "codccost", "login", sql, "T")
        If cadCC <> "" Then
            'comprobar que el Centro de Coste existe en la Contabilidad
            sql = DevuelveDesdeBDNew(cConta, "cabccost", "codccost", "codccost", cadCC, "T")
            If sql <> "" Then
                ComprobarCCoste = True
            Else
                sql = "No existe el CC: " & cadCC
                sql = "Comprobando Centros de Coste en contabilidad..." & vbCrLf & vbCrLf & sql
                MsgBox sql, vbExclamation
            End If
        Else 'el usuario no tiene asignado un centro de coste
            sql = "El trabajador conectado no tiene asignado un centro de coste."
            sql = "Comprobando Centros de Coste ..." & vbCrLf & vbCrLf & sql
            MsgBox sql, vbExclamation
        End If
    End If
    
ECCoste:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Centros de Coste", Err.Description
    End If
End Function


Public Function ComprobarCCoste_new(cadCC As String, cadTabla As String, Optional Opcion As Byte) As Boolean
Dim sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECCoste

    ComprobarCCoste_new = False
    Select Case cadTabla
        Case "facturas" ' facturas de venta
            If Opcion = 1 Then
                sql = "select distinct variedades.codccost from facturas_variedad, albaran_variedad, variedades, tmpFactu where "
                sql = sql & " albaran_variedad.codvarie=variedades.codvarie and "
                sql = sql & " facturas_variedad.codtipom=tmpFactu.codtipom AND facturas_variedad.numfactu=tmpFactu.numfactu AND facturas_variedad.fecfactu=tmpFactu.fecfactu and  "
                sql = sql & " albaran_variedad.numalbar = facturas_variedad.numalbar and "
                sql = sql & " albaran_variedad.numlinea = facturas_variedad.numlinealbar "
            Else
                sql = sql & " select distinct sfamia.codccost from facturas_envases, sartic, sfamia, tmpFactu where "
                sql = sql & " facturas_envases.codtipom=tmpFactu.codtipom AND facturas_envases.numfactu=tmpFactu.numfactu AND facturas_envases.fecfactu=tmpFactu.fecfactu and  "
                sql = sql & " facturas_envases.codartic = sartic.codartic and "
                sql = sql & " sartic.codfamia = sfamia.codfamia "
            End If
        Case "scafpc" ' facturas de compra
            sql = sql & " select distinct sfamia.codccost from slifpc, sartic, sfamia, tmpFactu where "
            sql = sql & " slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu and  "
            sql = sql & " slifpc.codartic = sartic.codartic and "
            sql = sql & " sartic.codfamia = sfamia.codfamia "
        
        Case "rfactsoc" ' facturas de socio
            sql = "select distinct variedades.codccost from rfactsoc_variedad,  variedades, tmpFactu where "
            sql = sql & " rfactsoc_variedad.codvarie=variedades.codvarie and "
            sql = sql & " rfactsoc_variedad.codtipom=tmpFactu.codtipom AND rfactsoc_variedad.numfactu=tmpFactu.numfactu AND rfactsoc_variedad.fecfactu=tmpFactu.fecfactu "
        
        Case "rfacttra" ' facturas de transporte
            sql = "select distinct variedades.codccost from rfacttra_albaran,  variedades, tmpFactu where "
            sql = sql & " rfacttra_albaran.codvarie=variedades.codvarie and "
            sql = sql & " rfacttra_albaran.codtipom=tmpFactu.codtipom AND rfacttra_albaran.numfactu=tmpFactu.numfactu AND rfacttra_albaran.fecfactu=tmpFactu.fecfactu "
        
        Case "rcafter" ' facturas de terceros
            sql = "select distinct variedades.codccost from rlifter,  variedades, tmpFactu where "
            sql = sql & " rlifter.codvarie=variedades.codvarie and "
            sql = sql & " rlifter.codsocio=tmpFactu.codsocio AND rlifter.numfactu=tmpFactu.numfactu AND rlifter.fecfactu=tmpFactu.fecfactu "
    
        Case "advfacturas" ' facturas de venta adv
            sql = sql & "select distinct advfamia.codccost from advfacturas_lineas, advartic, advfamia, tmpFactu where "
            sql = sql & " advfacturas_lineas.codtipom=tmpFactu.codtipom AND advfacturas_lineas.numfactu=tmpFactu.numfactu AND advfacturas_lineas.fecfactu=tmpFactu.fecfactu and  "
            sql = sql & " advfacturas_lineas.codartic = advartic.codartic and "
            sql = sql & " advartic.codfamia = advfamia.codfamia "
        
        Case "rrecibpozos" ' recibos de consumo de pozos
            sql = sql & "select distinct " & DBSet(vParamAplic.CodCCostPOZ, "T") & " as codccost from rrecibpozos where 1=1 "
        
        Case "rbodfacturas" ' facturas de retirada de bodega / almazara
            sql = "select distinct variedades.codccost from rbodfacturas_lineas, variedades, tmpFactu where "
            sql = sql & " rbodfacturas_lineas.codvarie=variedades.codvarie and rbodfacturas_lineas.codtipom=tmpFactu.codtipom and "
            sql = sql & " rbodfacturas_lineas.numfactu=tmpFactu.numfactu AND rbodfacturas_lineas.fecfactu=tmpFactu.fecfactu "
    
        Case "fvarcabfact" ' facturas de tipo clientes varias
            sql = "select distinct fvarconce.codccost from fvarlinfact, fvarconce, tmpFactu where "
            sql = sql & " fvarlinfact.codconce=fvarconce.codconce and fvarlinfact.codtipom=tmpFactu.codtipom and "
            sql = sql & " fvarlinfact.numfactu=tmpFactu.numfactu AND fvarlinfact.fecfactu=tmpFactu.fecfactu "
    
        Case "fvarcabfactpro" ' facturas de tipo proveedor varias
            sql = "select distinct fvarconce.codccost from fvarlinfactpro, fvarconce, tmpFactu where "
            sql = sql & " fvarlinfactpro.codconce=fvarconce.codconce and fvarlinfactpro.codtipom=tmpFactu.codtipom and "
            sql = sql & " fvarlinfactpro.numfactu=tmpFactu.numfactu AND fvarlinfactpro.fecfactu=tmpFactu.fecfactu "
    End Select
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = True

    While Not Rs.EOF And b
        sql = DevuelveDesdeBDNew(cConta, "cabccost", "codccost", "codccost", DBLet(Rs.Fields(0).Value), "T")
        If sql = "" Then
            b = False
            Sql2 = "Centro de Coste: " & Rs.Fields(0)
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
    If Not b Then
        sql = "No existe el " & Sql2
        sql = "Comprobando Centros de Coste en contabilidad..." & vbCrLf & vbCrLf & sql
    
        MsgBox sql, vbExclamation
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
Dim sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECCoste

    ComprobarFormadePago = False
    
    Select Case cadCC
        Case "advfacturas"
            sql = "select distinct advfacturas.codforpa from advfacturas, tmpFactu where "
            sql = sql & " advfacturas.codtipom=tmpFactu.codtipom AND advfacturas.numfactu=tmpFactu.numfactu AND advfacturas.fecfactu=tmpFactu.fecfactu  "
        Case "rbodfacturas"
            sql = "select distinct rbodfacturas.codforpa from rbodfacturas, tmpFactu where "
            sql = sql & " rbodfacturas.codtipom=tmpFactu.codtipom AND rbodfacturas.numfactu=tmpFactu.numfactu AND rbodfacturas.fecfactu=tmpFactu.fecfactu  "
    End Select
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = True

    While Not Rs.EOF And b
        sql = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", Rs.Fields(0).Value, "N")
        If sql = "" Then
            b = False
            Sql2 = "Formas de Pago: " & Rs.Fields(0)
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
    If Not b Then
        sql = "No existe la " & Sql2
        sql = "Comprobando Formas de Pago en contabilidad..." & vbCrLf & vbCrLf & sql
    
        MsgBox sql, vbExclamation
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




Public Function PasarFactura(cadwhere As String, CodCCost As String, CtaBan As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariges.scafac --> conta.cabfact
' ariges.slifac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim sql As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    
    'Insertar en la conta Cabecera Factura
    b = InsertarCabFact(cadwhere, cadMen)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
        CCoste = CodCCost
        'Insertar lineas de Factura en la Conta
        b = InsertarLinFact_new("facturas", cadwhere, cadMen)
        cadMen = "Insertando Lin. Factura: " & cadMen

        '++monica:añadida la parte de insertar en tesoreria
        If b Then
            b = InsertarEnTesoreriaNewFac(cadwhere, CtaBan, cadMen)
            cadMen = "Insertando en Tesoreria: " & cadMen
        End If
        
        '++


        If b Then
            'Poner intconta=1 en ariagro.facturas
            b = ActualizarCabFact("facturas", cadwhere, cadMen)
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
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFactura = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFactura = False
        
        sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
        sql = sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
        sql = sql & " WHERE " & Replace(cadwhere, "facturas", "tmpFactu")
        conn.Execute sql
    End If
End Function

Public Function PasarFacturaADV(cadwhere As String, CodCCost As String, CtaBan As String, FecVen As String, TipoM As String, FecFac As Date, Observac As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariges.scafac --> conta.cabfact
' ariges.slifac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim sql As String
Dim Mc As Contadores
Dim Obs As String


    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    
    'Insertar en la conta Cabecera Factura
    
    If TipoM <> "FIN" Then
        
        b = InsertarCabFactADV(cadwhere, Observac, cadMen)
        cadMen = "Insertando Cab. Factura: " & cadMen
        
        If b Then
            CCoste = CodCCost
            'Insertar lineas de Factura en la Conta
            b = InsertarLinFactADV("advfacturas", cadwhere, cadMen)
            cadMen = "Insertando Lin. Factura: " & cadMen
    
            '++monica:añadida la parte de insertar en tesoreria
            If b Then
                b = InsertarEnTesoreriaNewADV(cadwhere, CtaBan, FecVen, cadMen)
                cadMen = "Insertando en Tesoreria: " & cadMen
            End If
        End If
            '++
    Else
        ' No insertamos la factura sino un asiento en el diario
        Set Mc = New Contadores
        
        If Mc.ConseguirContador("0", (CDate(FecFac) <= CDate(vEmpresa.FechaFin)), True) = 0 Then
        
            Obs = "Contabilización Factura Interna de Fecha " & Format(FecFac, "dd/mm/yyyy")
        
            'Insertar en la conta Cabecera Asiento
            b = InsertarCabAsientoDia(vEmpresa.NumDiarioInt, Mc.Contador, CStr(Format(FecFac, "dd/mm/yyyy")), Obs, cadMen)
            cadMen = "Insertando Cab. Asiento: " & cadMen
        Else
            b = False
        End If
    
        If b Then
            Socio = DevuelveValor("select codsocio from advfacturas where " & cadwhere)
            CtaSocio = DevuelveValor("select codmaccli from rsocios_seccion where codsocio = " & Socio & " and codsecci = " & vParamAplic.SeccionADV)
        
        
            b = InsertarLinAsientoFactInt("advfacturas", cadwhere, cadMen, CtaSocio, Mc.Contador)
            cadMen = "Insertando Lin. Factura Interna: " & cadMen
        
            Set Mc = Nothing
        End If
    
        '++monica:añadida la parte de insertar en tesoreria
        If b Then
            CCoste = CodCCost
            b = InsertarEnTesoreriaNewADV(cadwhere, CtaBan, FecVen, cadMen)
            cadMen = "Insertando en Tesoreria: " & cadMen
        End If
    
    End If

    If b Then
        'Poner intconta=1 en ariagro.facturas
        b = ActualizarCabFact("advfacturas", cadwhere, cadMen)
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
        b = False
        MuestraError Err.Number, "Contabilizando Factura ADV", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaADV = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaADV = False
        
        sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
        sql = sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
        sql = sql & " WHERE " & Replace(cadwhere, "advfacturas", "tmpFactu")
        conn.Execute sql
    End If
End Function

Public Function InsertarCabAsientoDia(Diario As String, Asiento As String, Fecha As String, Obs As String, cadErr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String

    On Error GoTo EInsertar
       
    
    Cad = Format(Diario, "00") & ", " & DBSet(Fecha, "F") & "," & Format(Asiento, "000000") & ","
    Cad = Cad & "''," & ValorNulo & "," & DBSet(Obs, "T")
    Cad = "(" & Cad & ")"

    'Insertar en la contabilidad
    sql = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) "
    sql = sql & " VALUES " & Cad
    ConnConta.Execute sql
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabAsientoDia = False
        cadErr = Err.Description
    Else
        InsertarCabAsientoDia = True
    End If
End Function


Private Function InsertarLinAsientoFactInt(cadTabla As String, cadwhere As String, cadErr As String, CtaSocio As String, Optional Contador As Long) As Boolean
Dim sql As String
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
Dim b As Boolean
Dim Cad As String
Dim cadMen As String
Dim FeFact As Date

Dim cadCampo As String

    On Error GoTo eInsertarLinAsientoFactInt

    InsertarLinAsientoFactInt = False

    TotalFac = DevuelveValor("select totalfac from advfacturas where " & cadwhere)
    'utilizamos sfamia.ctaventa o sfamia.aboventa
    If TotalFac >= 0 Then
        cadCampo = "advfamia.ctaventa"
    Else
        cadCampo = "advfamia.aboventa"
    End If
    
    If vEmpresa.TieneAnalitica Then
        sql = " SELECT usuarios.stipom.letraser,advfacturas_lineas.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, advfamia.codccost "
    Else
        sql = " SELECT usuarios.stipom.letraser,advfacturas_lineas.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe "
    End If
    
    sql = sql & " FROM ((advfacturas_lineas inner join usuarios.stipom on advfacturas_lineas.codtipom=usuarios.stipom.codtipom) "
    sql = sql & " inner join advartic on advfacturas_lineas.codartic=advartic.codartic) "
    sql = sql & " inner join advfamia on advartic.codfamia=advfamia.codfamia "
    sql = sql & " WHERE " & Replace(cadwhere, "advfacturas", "advfacturas_lineas")
    If vEmpresa.TieneAnalitica Then
        sql = sql & " GROUP BY 5,7 " '& cadCampo, codccost
    Else
        sql = sql & " GROUP BY 5 " '& cadCampo
    End If

    
    Set Rs = New ADODB.Recordset
    
    Rs.Open sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
            
    I = 0
    ImporteD = 0
    ImporteH = 0
    
    numdocum = Format(Rs!numfactu, "0000000")
    Ampliacion = Rs.Fields(0).Value & "-" & Format(Rs!numfactu, "0000000")
    ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoInt, "N")) & " " & Ampliacion
    ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoInt, "N")) & " " & Ampliacion
    
    If Not Rs.EOF Then Rs.MoveFirst
    
    b = True
    
    
    
    While Not Rs.EOF And b
        I = I + 1
        
        FeFact = Rs!fecfactu
        
        Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Contador, "N") & ","
        Cad = Cad & DBSet(I, "N") & "," & DBSet(Rs!cuenta, "T") & "," & DBSet(numdocum, "T") & ","
        
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If Rs.Fields(5).Value < 0 Then
            ' importe al debe en positivo
            Cad = Cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Rs.Fields(5).Value * (-1), "N") & ","
            Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
        
            ImporteD = ImporteD + (CCur(Rs.Fields(5).Value) * (-1))
        Else
            ' importe al haber en positivo, cambiamos el signo
            Cad = Cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
            Cad = Cad & DBSet((Rs.Fields(5).Value), "N") & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
        
            ImporteH = ImporteH + CCur(Rs.Fields(5).Value)
        End If
        
        Cad = "(" & Cad & ")"
        
        b = InsertarLinAsientoDia(Cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & I

        Rs.MoveNext
    Wend
    
    If b And I > 0 Then
        I = I + 1
                
        ' el Total es sobre la cuenta del cliente
        Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FeFact, "F") & "," & DBSet(Contador, "N") & ","
        Cad = Cad & DBSet(I, "N") & ","
        Cad = Cad & DBSet(CtaSocio, "T") & "," & DBSet(numdocum, "T") & ","
            
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If ImporteD - ImporteH > 0 Then
            ' importe al debe en positivo
            Cad = Cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliaciond, "T") & "," & ValorNulo & ","
            Cad = Cad & DBSet(ImporteD - ImporteH, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
        Else
            ' importe al haber en positivo, cambiamos el signo
            Cad = Cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliacionh, "T") & "," & DBSet(((ImporteD - ImporteH) * -1), "N") & ","
            Cad = Cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
        End If
        
        Cad = "(" & Cad & ")"
        
        b = InsertarLinAsientoDia(Cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & I
        
    End If
        
    Set Rs = Nothing
    InsertarLinAsientoFactInt = b
    Exit Function
    
eInsertarLinAsientoFactInt:
    cadErr = "Insertar Linea Asiento Factura Interna: " & Err.Description
    cadErr = cadErr & cadMen
    InsertarLinAsientoFactInt = False
End Function


Public Function InsertarLinAsientoDia(Cad As String, cadErr As String) As Boolean
' el Tipo me indica desde donde viene la llamada
' tipo = 0 srecau.codmacta
' tipo = 1 scaalb.codmacta

Dim Rs As ADODB.Recordset
Dim Aux As String
Dim sql As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

 
    sql = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
    sql = sql & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) "
    sql = sql & " VALUES " & Cad
    
    ConnConta.Execute sql

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinAsientoDia = False
        cadErr = Err.Description
    Else
        InsertarLinAsientoDia = True
    End If
End Function



Public Function PasarFacturaBOD(cadwhere As String, CodCCost As String, CtaBan As String, FecVen As String, Tipo As Byte) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariagro.rbodfacturas --> conta.cabfact
' ariagro.rbodfacturas_variedad --> conta.linfact
'Actualizar la tabla ariagro.rbodfacturas.inconta=1 para indicar que ya esta contabilizada
'Tipo : 0 = facturas de retirada de almazara
'       1 = facturas de retirada de bodega

Dim b As Boolean
Dim cadMen As String
Dim sql As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    
    'Insertar en la conta Cabecera Factura
    b = InsertarCabFactBOD(cadwhere, cadMen, Tipo)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
        CCoste = CodCCost
        'Insertar lineas de Factura en la Conta
        Select Case Tipo
            Case 0
                b = InsertarLinFactBOD("rbodfact1", cadwhere, cadMen)
            Case 1
                b = InsertarLinFactBOD("rbodfact2", cadwhere, cadMen)
        End Select
        
        'b = InsertarLinFactBOD("rbodfacturas", cadWHERE, cadMen)
        cadMen = "Insertando Lin. Factura: " & cadMen

        '++monica:añadida la parte de insertar en tesoreria
        If b Then
            b = InsertarEnTesoreriaNewBOD(cadwhere, CtaBan, FecVen, cadMen, Tipo)
            cadMen = "Insertando en Tesoreria: " & cadMen
        End If
        
        '++


        If b Then
            'Poner intconta=1 en ariagro.facturas
            b = ActualizarCabFact("rbodfacturas", cadwhere, cadMen)
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
        b = False
        MuestraError Err.Number, "Contabilizando Factura Retirada", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaBOD = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaBOD = False
        
        sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
        sql = sql & " Select tmpfactu.*," & DBSet(cadMen, "T") & " as error From tmpFactu "
        sql = sql & " WHERE " & Replace(cadwhere, "rbodfacturas", "tmpFactu")
        conn.Execute sql
    End If
End Function


Public Function PasarFacturaTel(cadwhere As String, CodCCost As String, CtaVtas As String, CodIva As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariagro.rbodfacturas --> conta.cabfact
' ariagro.rbodfacturas_variedad --> conta.linfact
'Actualizar la tabla ariagro.rbodfacturas.inconta=1 para indicar que ya esta contabilizada
'Tipo : 0 = facturas de retirada de almazara
'       1 = facturas de retirada de bodega

Dim b As Boolean
Dim cadMen As String
Dim sql As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    CodiIVA = CodIva
    
    'Insertar en la conta Cabecera Factura
    b = InsertarCabFactTEL(cadwhere, cadMen)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
        CCoste = CodCCost
        'Insertar lineas de Factura en la Conta
        b = InsertarLinFactTEL(CtaVtas, cadwhere, cadMen)
        
        cadMen = "Insertando Lin. Factura: " & cadMen

'--Monica: quitado de momento
'        '++monica:añadida la parte de insertar en tesoreria
'        If b Then
'            b = InsertarEnTesoreriaNewBOD(cadWHERE, CtaBan, FecVen, cadMen, Tipo)
'            cadMen = "Insertando en Tesoreria: " & cadMen
'        End If
'
        '++


        If b Then
            'Poner intconta=1 en ariagro.facturas
            b = ActualizarCabFact("rtelmovil", cadwhere, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
    
EContab:
    
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura Telefonia", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaTel = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaTel = False
        
        sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
        sql = sql & " Select tmpfactu.*," & DBSet(cadMen, "T") & " as error From tmpFactu "
        sql = sql & " WHERE " & Replace(cadwhere, "rtelmovil", "tmpFactu")
        conn.Execute sql
    End If
End Function





Private Function InsertarCabFact(cadwhere As String, cadErr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String


    On Error GoTo EInsertar
    
    sql = sql & " SELECT stipom.letraser,numfactu,fecfactu, clientes.codmacta,clientes.cliabono,year(fecfactu) as anofaccl,"
    sql = sql & "baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    sql = sql & "totalfac,codiiva1,codiiva2,codiiva3, porcrec1, porcrec2, porcrec3, imporec1, imporec2, imporec3 "
    sql = sql & " FROM (" & "facturas inner join " & "stipom on facturas.codtipom=stipom.codtipom) "
    sql = sql & "INNER JOIN " & "clientes ON facturas.codclien=clientes.codclien "
    sql = sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = 0
        DtoGnral = 0
        BaseImp = Rs!baseimp1 + CCur(DBLet(Rs!baseimp2, "N")) + CCur(DBLet(Rs!baseimp3, "N"))
        '---- Laura 10/10/2006:  añadir el totalfac para utilizarlo en insertar lineas
        TotalFac = Rs!TotalFac
        '----
        conCtaAlt = Rs!cliAbono
        
        sql = ""
        sql = "'" & Rs!LetraSer & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!fecfactu) & "," & ValorNulo & ","
        sql = sql & DBSet(Rs!baseimp1, "N") & "," & DBSet(Rs!baseimp2, "N", "S") & "," & DBSet(Rs!baseimp3, "N", "S") & "," & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", "S") & "," & DBSet(Rs!porciva3, "N", "S") & ","
        sql = sql & DBSet(Rs!porcrec1, "N") & "," & DBSet(Rs!porcrec2, "N", "S") & "," & DBSet(Rs!porcrec3, "N", "S") & "," & DBSet(Rs!ImpoIva1, "N", "N") & "," & DBSet(Rs!impoIVA2, "N", "S") & "," & DBSet(Rs!impoIVA3, "N", "S") & ","
        sql = sql & DBSet(Rs!imporec1, "N", "N") & "," & DBSet(Rs!imporec2, "N", "S") & "," & DBSet(Rs!imporec3, "N", "S") & ","
        sql = sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!codiiva1, "N") & "," & DBSet(Rs!codiiva2, "N", "S") & "," & DBSet(Rs!codiiva3, "N", "S") & ",0,"
        sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        sql = sql & DBSet(Rs!fecfactu, "F")
        Cad = Cad & "(" & sql & ")"
'        RS.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
    
    'Insertar en la contabilidad
    sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
    sql = sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
    sql = sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
    sql = sql & " VALUES " & Cad
    ConnConta.Execute sql
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFact = False
        cadErr = Err.Description
    Else
        InsertarCabFact = True
    End If
End Function


Private Function InsertarCabFactADV(cadwhere As String, Observac As String, cadErr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String


    On Error GoTo EInsertar
    
    sql = "SELECT stipom.letraser,numfactu,fecfactu, rsocios_seccion.codmaccli as codmacta,year(fecfactu) as anofaccl,"
    sql = sql & "baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    sql = sql & "totalfac,codiiva1,codiiva2,codiiva3, porcrec1, porcrec2, porcrec3, imporec1, imporec2, imporec3 "
    sql = sql & " FROM ((" & "advfacturas inner join " & "usuarios.stipom on advfacturas.codtipom=usuarios.stipom.codtipom) "
    sql = sql & "INNER JOIN rsocios ON advfacturas.codsocio=rsocios.codsocio) "
    sql = sql & "INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionADV, "N")
    sql = sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = 0
        DtoGnral = 0
        BaseImp = Rs!baseimp1 + CCur(DBLet(Rs!baseimp2, "N")) + CCur(DBLet(Rs!baseimp3, "N"))
        '---- Laura 10/10/2006:  añadir el totalfac para utilizarlo en insertar lineas
        TotalFac = Rs!TotalFac
        '----
        
        sql = ""
        sql = "'" & Rs!LetraSer & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!fecfactu) & ","
        '[Monica]02/05/2012: añadido campo observaciones del frame, antes valor nulo
        sql = sql & DBSet(Observac, "T") & "," '& ValorNulo & ","
        
        sql = sql & DBSet(Rs!baseimp1, "N") & "," & DBSet(Rs!baseimp2, "N", "S") & "," & DBSet(Rs!baseimp3, "N", "S") & "," & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", "S") & "," & DBSet(Rs!porciva3, "N", "S") & ","
        sql = sql & DBSet(Rs!porcrec1, "N", "S") & "," & DBSet(Rs!porcrec2, "N", "S") & "," & DBSet(Rs!porcrec3, "N", "S") & "," & DBSet(Rs!ImpoIva1, "N", "N") & "," & DBSet(Rs!impoIVA2, "N", "S") & "," & DBSet(Rs!impoIVA3, "N", "S") & ","
        sql = sql & DBSet(Rs!imporec1, "N", "S") & "," & DBSet(Rs!imporec2, "N", "S") & "," & DBSet(Rs!imporec3, "N", "S") & ","
        sql = sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!codiiva1, "N") & "," & DBSet(Rs!codiiva2, "N", "S") & "," & DBSet(Rs!codiiva3, "N", "S") & ",0,"
        sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        sql = sql & DBSet(Rs!fecfactu, "F")
        Cad = Cad & "(" & sql & ")"
'        RS.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    'Insertar en la contabilidad
    sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
    sql = sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
    sql = sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
    sql = sql & " VALUES " & Cad
    ConnConta.Execute sql
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactADV = False
        cadErr = Err.Description
    Else
        InsertarCabFactADV = True
    End If
End Function

Private Function InsertarCabFactBOD(cadwhere As String, cadErr As String, Tipo As Byte) As Boolean
'Insertando en tabla conta.cabfact
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Seccion As Integer

    On Error GoTo EInsertar
    
    Select Case Tipo
        Case 0
            Seccion = vParamAplic.SeccionAlmaz
        Case 1
            Seccion = vParamAplic.SeccionBodega
    End Select
    
    
    sql = "SELECT stipom.letraser,numfactu,fecfactu, rsocios_seccion.codmaccli as codmacta,year(fecfactu) as anofaccl,"
    sql = sql & "baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    sql = sql & "totalfac,codiiva1,codiiva2,codiiva3, porcrec1, porcrec2, porcrec3, imporec1, imporec2, imporec3 "
    sql = sql & " FROM ((" & "rbodfacturas inner join " & "usuarios.stipom on rbodfacturas.codtipom=usuarios.stipom.codtipom) "
    sql = sql & "INNER JOIN rsocios ON rbodfacturas.codsocio=rsocios.codsocio) "
    sql = sql & "INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N")
    sql = sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = 0
        DtoGnral = 0
        BaseImp = Rs!baseimp1 + CCur(DBLet(Rs!baseimp2, "N")) + CCur(DBLet(Rs!baseimp3, "N"))
        '---- Laura 10/10/2006:  añadir el totalfac para utilizarlo en insertar lineas
        TotalFac = Rs!TotalFac
        '----
        
        sql = ""
        sql = "'" & Rs!LetraSer & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!fecfactu) & "," & ValorNulo & ","
        sql = sql & DBSet(Rs!baseimp1, "N") & "," & DBSet(Rs!baseimp2, "N", "S") & "," & DBSet(Rs!baseimp3, "N", "S") & "," & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", "S") & "," & DBSet(Rs!porciva3, "N", "S") & ","
        sql = sql & DBSet(Rs!porcrec1, "N", "S") & "," & DBSet(Rs!porcrec2, "N", "S") & "," & DBSet(Rs!porcrec3, "N", "S") & "," & DBSet(Rs!ImpoIva1, "N", "N") & "," & DBSet(Rs!impoIVA2, "N", "S") & "," & DBSet(Rs!impoIVA3, "N", "S") & ","
        sql = sql & DBSet(Rs!imporec1, "N", "S") & "," & DBSet(Rs!imporec2, "N", "S") & "," & DBSet(Rs!imporec3, "N", "S") & ","
        sql = sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!codiiva1, "N") & "," & DBSet(Rs!codiiva2, "N", "S") & "," & DBSet(Rs!codiiva3, "N", "S") & ",0,"
        sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        sql = sql & DBSet(Rs!fecfactu, "F")
        Cad = Cad & "(" & sql & ")"
'        RS.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    'Insertar en la contabilidad
    sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
    sql = sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
    sql = sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
    sql = sql & " VALUES " & Cad
    ConnConta.Execute sql
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactBOD = False
        cadErr = Err.Description
    Else
        InsertarCabFactBOD = True
    End If
End Function



Private Function InsertarCabFactTEL(cadwhere As String, cadErr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Seccion As Integer
Dim PorcIva As String

    On Error GoTo EInsertar
    
    Seccion = vParamAplic.Seccionhorto
    
    sql = "SELECT numserie,numfactu,fecfactu, rsocios_seccion.codmaccli as codmacta,year(fecfactu) as anofaccl,"
    sql = sql & "baseimpo,cuotaiva,totalfac"
    sql = sql & " FROM (rtelmovil "
    sql = sql & "INNER JOIN rsocios ON rtelmovil.codsocio=rsocios.codsocio) "
    sql = sql & "INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N")
    sql = sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = 0
        DtoGnral = 0
        BaseImp = Rs!baseimpo
        '---- Laura 10/10/2006:  añadir el totalfac para utilizarlo en insertar lineas
        TotalFac = Rs!TotalFac
        '----
        
        PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CodiIVA, "N")
        
        sql = ""
        sql = DBSet(Rs!numserie, "T") & "," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!fecfactu) & "," & ValorNulo & ","
        sql = sql & DBSet(Rs!baseimpo, "N") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(PorcIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
        sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!CuotaIva, "N", "N") & "," & ValorNulo & "," & ValorNulo & ","
        sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        sql = sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(CodiIVA, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
        sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        sql = sql & DBSet(Rs!fecfactu, "F")
        Cad = Cad & "(" & sql & ")"
'        RS.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    'Insertar en la contabilidad
    sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
    sql = sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
    sql = sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
    sql = sql & " VALUES " & Cad
    ConnConta.Execute sql
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactTEL = False
        cadErr = Err.Description
    Else
        InsertarCabFactTEL = True
    End If
End Function



Private Function InsertarLinFact(cadTabla As String, cadwhere As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

    If cadTabla = "scafac" Then
        sql = " SELECT stipom.letraser,slifac.codtipom,numfactu,fecfactu,sartic.codfamia,sfamia.ctaventa,sfamia.ctavent1,sfamia.aboventa,sfamia.abovent1,sum(importel) as importe "
        sql = sql & " FROM ((slifac inner join stipom on slifac.codtipom=stipom.codtipom) "
        sql = sql & " inner join sartic on slifac.codartic=sartic.codartic) "
        sql = sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        sql = sql & " WHERE " & Replace(cadwhere, "scafac", "slifac")
        sql = sql & " GROUP BY sfamia.codfamia "
    Else
        sql = " SELECT slifpc.codprove,numfactu,fecfactu,sartic.codfamia,sfamia.ctacompr,sfamia.abocompr,sum(importel) as importe "
        sql = sql & " FROM (slifpc  "
        sql = sql & " inner join sartic on slifpc.codartic=sartic.codartic) "
        sql = sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        sql = sql & " WHERE " & Replace(cadwhere, "scafpc", "slifpc")
        sql = sql & " GROUP BY sfamia.codfamia "
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = Cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
        ImpLinea = Rs!Importe - CalcularPorcentaje(Rs!Importe, DtoPPago, 2)
        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
        ImpLinea = ImpLinea - CalcularPorcentaje(Rs!Importe, DtoGnral, 2)
        'ImpLinea = Round(ImpLinea, 2)
        '----
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        sql = ""
        Sql2 = ""
        If cadTabla = "scafac" Then
            sql = "'" & Rs!LetraSer & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
            If Not conCtaAlt Then 'cliente no tiene cuenta alternativa
                If ImpLinea >= 0 Then
                    sql = sql & DBSet(Rs!ctaventa, "T")
                Else
                    sql = sql & DBSet(Rs!aboventa, "T")
                End If
            Else
                If ImpLinea >= 0 Then
                    sql = sql & DBSet(Rs!ctavent1, "T")
                Else
                    sql = sql & DBSet(Rs!abovent1, "T")
                End If
            End If
        Else
            sql = NumRegis & "," & Year(Rs!fecfactu) & "," & I & ","
            If ImpLinea >= 0 Then
                sql = sql & DBSet(Rs!ctacompr, "T")
            Else
                sql = sql & DBSet(Rs!abocompr, "T")
            End If
        End If
        Sql2 = sql & ","
        sql = sql & "," & DBSet(ImpLinea, "N") & ","
        
        If CCoste = "" Then
            sql = sql & ValorNulo
        Else
            sql = sql & DBSet(CCoste, "T")
        End If
        
        Cad = Cad & "(" & sql & ")" & ","
        
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
        If CCoste = "" Then
            Sql2 = Sql2 & ValorNulo
        Else
            Sql2 = Sql2 & DBSet(CCoste, "T")
        End If
        If SQLaux <> "" Then 'hay mas de una linea
            Cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            Cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If


    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        If cadTabla = "scafac" Then
            sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Else
            sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        End If
        sql = sql & " VALUES " & Cad
        ConnConta.Execute sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact = False
        cadErr = Err.Description
    Else
        InsertarLinFact = True
    End If
End Function





Private Function InsertarLinFact_new(cadTabla As String, cadwhere As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
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
            sql = " SELECT stipom.letraser,facturas_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sfamia.codccost "
        Else
            sql = " SELECT stipom.letraser,facturas_envases.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe "
        End If
        
        sql = sql & " FROM ((facturas_envases inner join stipom on facturas_envases.codtipom=stipom.codtipom) "
        sql = sql & " inner join sartic on facturas_envases.codartic=sartic.codartic) "
        sql = sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        sql = sql & " WHERE " & Replace(cadwhere, "facturas", "facturas_envases")
        If vEmpresa.TieneAnalitica Then
            sql = sql & " GROUP BY 5,7 " '& cadCampo, codccost
        Else
            sql = sql & " GROUP BY 5 " '& cadCampo
        End If
        sql = sql & "Union "
        
        If vEmpresa.TieneAnalitica Then
            sql = sql & " SELECT stipom.letraser,facturas_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe, variedades.codccost "
        Else
            sql = sql & " SELECT stipom.letraser,facturas_variedad.codtipom,numfactu,fecfactu," & CadCampo1 & " as cuenta,sum(impornet) as importe "
        End If
        sql = sql & " FROM (((((facturas_variedad inner join stipom on facturas_variedad.codtipom=stipom.codtipom) "
        sql = sql & " inner join albaran on facturas_variedad.numalbar = albaran.numalbar) "
        sql = sql & " inner join tipomer on albaran.codtimer = tipomer.codtimer) "
        sql = sql & " inner join albaran_variedad on facturas_variedad.numalbar = albaran_variedad.numalbar and facturas_variedad.numlinealbar = albaran_variedad.numlinea) "
        sql = sql & " inner join variedades on albaran_variedad.codvarie=variedades.codvarie) "
        sql = sql & " WHERE " & Replace(cadwhere, "facturas", "facturas_variedad")
        If vEmpresa.TieneAnalitica Then
            sql = sql & " GROUP BY 5,7 " '& cadCampo1, codccost
        Else
            sql = sql & " GROUP BY 5 " '& cadCampo1
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
                sql = " SELECT slifpc.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sfamia.codccost"
            Else
                sql = " SELECT slifpc.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe"
            End If
            sql = sql & " FROM (slifpc  "
            sql = sql & " inner join sartic on slifpc.codartic=sartic.codartic) "
            sql = sql & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
            sql = sql & " WHERE " & Replace(cadwhere, "scafpc", "slifpc")
            sql = sql & " GROUP BY " & cadCampo
            If vEmpresa.TieneAnalitica Then
                sql = sql & ", sfamia.codccost "
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
                 sql = " SELECT 2, variedades.ctacomtercero as cuenta, sum(rlifter.importel) as importe, variedades.codccost "
            Else
                 sql = " SELECT 2, variedades.ctacomtercero as cuenta, sum(rlifter.importel) as importe "
            End If
             sql = sql & " FROM rlifter, variedades "
             sql = sql & " WHERE " & Replace(cadwhere, "rcafter", "rlifter") & " and"
             sql = sql & " rlifter.codvarie = variedades.codvarie "
             sql = sql & " group by 1,2 "
             
             '[Monica]23/09/2013: concepto de gasto
             sql = sql & " union "
             If vEmpresa.TieneAnalitica Then
                sql = sql & " select 1, fvarconce.codmacpr as cuenta, rcafter.impcargo as importe, '' "
             Else
                sql = sql & " select 1, fvarconce.codmacpr as cuenta, rcafter.impcargo as importe "
             End If
             sql = sql & " FROM rcafter, fvarconce "
             sql = sql & " WHERE " & cadwhere & " and"
             sql = sql & " rcafter.concepcargo = fvarconce.codconce "
             sql = sql & " group by 1,2 "
             
             sql = sql & " order by 1,2 "


        End If
    End If
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = Cad
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
        sql = ""
        Sql2 = ""
        
        If cadTabla = "facturas" Then 'VENTAS a clientes
            sql = "'" & Rs!LetraSer & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
            sql = sql & DBSet(Rs!cuenta, "T")
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
                sql = NumRegis & "," & AnyoFacPr & "," & I & ","
                
    '            If ImpLinea >= 0 Then
                    sql = sql & DBSet(Rs!cuenta, "T")
    '            Else
    '                SQL = SQL & DBSet(RS!abocompr, "T")
    '            End If
            Else 'TRANSPORTE
                sql = NumRegis & "," & AnyoFacPr & "," & I & ","
                sql = sql & DBSet(Rs!cuenta, "T")
            End If
        End If
        
        Sql2 = sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        sql = sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
            If cadTabla = "rcafter" Then
                If DBLet(Rs!CodCCost, "T") = "----" Then
                    sql = sql & DBSet(CCoste, "T")
                Else
                    sql = sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBLet(Rs!CodCCost, "T")
                End If
            Else
                sql = sql & DBSet(Rs!CodCCost, "T")
                CCoste = DBSet(Rs!CodCCost, "T")
            End If
        Else
            sql = sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        Cad = Cad & "(" & sql & ")" & ","
        
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
            Cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            Cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If


    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        If cadTabla = "facturas" Then
            sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Else
            sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        End If
        sql = sql & " VALUES " & Cad
        ConnConta.Execute sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact_new = False
        cadErr = Err.Description
    Else
        InsertarLinFact_new = True
    End If
End Function


Private Function InsertarLinFactADV(cadTabla As String, cadwhere As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
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
            sql = " SELECT usuarios.stipom.letraser,advfacturas_lineas.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, advfamia.codccost "
        Else
            sql = " SELECT usuarios.stipom.letraser,advfacturas_lineas.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe "
        End If
        
        sql = sql & " FROM ((advfacturas_lineas inner join usuarios.stipom on advfacturas_lineas.codtipom=usuarios.stipom.codtipom) "
        sql = sql & " inner join advartic on advfacturas_lineas.codartic=advartic.codartic) "
        sql = sql & " inner join advfamia on advartic.codfamia=advfamia.codfamia "
        sql = sql & " WHERE " & Replace(cadwhere, "advfacturas", "advfacturas_lineas")
        If vEmpresa.TieneAnalitica Then
            sql = sql & " GROUP BY 5,7 " '& cadCampo, codccost
        Else
            sql = sql & " GROUP BY 5 " '& cadCampo
        End If
        
    End If
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = Cad
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
        sql = ""
        Sql2 = ""
        
        If cadTabla = "advfacturas" Then 'VENTAS a socios
            sql = "'" & Rs!LetraSer & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
            sql = sql & DBSet(Rs!cuenta, "T")
        End If
        
        Sql2 = sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        sql = sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
            sql = sql & DBSet(Rs!CodCCost, "T")
            CCoste = DBSet(Rs!CodCCost, "T")
        Else
            sql = sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        Cad = Cad & "(" & sql & ")" & ","
        
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
            Cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            Cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If


    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        sql = sql & " VALUES " & Cad
        ConnConta.Execute sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactADV = False
        cadErr = Err.Description
    Else
        InsertarLinFactADV = True
    End If
End Function


Private Function InsertarLinFactBOD(cadTabla As String, cadwhere As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
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
        sql = " SELECT usuarios.stipom.letraser,rbodfacturas_lineas.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, variedades.codccost "
    Else
        sql = " SELECT usuarios.stipom.letraser,rbodfacturas_lineas.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe "
    End If
    
    sql = sql & " FROM (rbodfacturas_lineas inner join usuarios.stipom on rbodfacturas_lineas.codtipom=usuarios.stipom.codtipom) "
    sql = sql & " inner join variedades on rbodfacturas_lineas.codvarie=variedades.codvarie "
    sql = sql & " WHERE " & Replace(cadwhere, "rbodfacturas", "rbodfacturas_lineas")
    If vEmpresa.TieneAnalitica Then
        sql = sql & " GROUP BY 5,7 " '& cadCampo, codccost
    Else
        sql = sql & " GROUP BY 5 " '& cadCampo
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = Cad
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
        sql = ""
        Sql2 = ""
        
        sql = "'" & Rs!LetraSer & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
        sql = sql & DBSet(Rs!cuenta, "T")
        
        Sql2 = sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        sql = sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
            sql = sql & DBSet(Rs!CodCCost, "T")
            CCoste = DBSet(Rs!CodCCost, "T")
        Else
            sql = sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        Cad = Cad & "(" & sql & ")" & ","
        
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
            Cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            Cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If


    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        sql = sql & " VALUES " & Cad
        ConnConta.Execute sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactBOD = False
        cadErr = Err.Description
    Else
        InsertarLinFactBOD = True
    End If
End Function


Private Function InsertarLinFactTEL(CtaVtas As String, cadwhere As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
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
        sql = " SELECT numserie,numfactu,fecfactu," & cadCampo & " as cuenta,baseimpo as importe, " & CCoste
    Else
        sql = " SELECT numserie,numfactu,fecfactu," & cadCampo & " as cuenta,baseimpo as importe "
    End If
    
    sql = sql & " FROM rtelmovil "
    sql = sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    If Not Rs.EOF Then
        SQLaux = Cad
        
        ImpLinea = DBLet(Rs!Importe, "N")
        totimp = totimp + DBLet(Rs!Importe, "N")

        'concatenamos linea para insertar en la tabla de conta.linfact
        sql = "'" & Rs!numserie & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
        sql = sql & DBSet(Rs!cuenta, "T")
        sql = sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
            sql = sql & DBSet(Rs!CodCCost, "T")
            CCoste = DBSet(Rs!CodCCost, "T")
        Else
            sql = sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        Cad = Cad & "(" & sql & ")"
        
        I = I + 1
        Rs.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
    

    'Insertar en la contabilidad
    If Cad <> "" Then
        sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        sql = sql & " VALUES " & Cad
        ConnConta.Execute sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactTEL = False
        cadErr = Err.Description
    Else
        InsertarLinFactTEL = True
    End If
End Function



Private Function InsertarLinFactSoc(cadTabla As String, cadwhere As String, cadErr As String, Tipo As Byte, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
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
        sql = "select mid(rfactsoc.codtipom,1,3) from " & cadTabla & " where " & cadwhere
        TipoFact = DevuelveValor(sql)
    
    End If
    
    If Tipo = 2 And TipoFact = "FLI" Then
        sql = "select rfactsoc.codsocio from " & cadTabla & " where " & cadwhere
        Socio = DevuelveValor(sql)
        
        Set vSocio = New cSocio
        If vSocio.LeerDatos(Socio) Then
            If vEmpresa.TieneAnalitica Then
                If vSocio.TipoProd = 1 Then
                    sql = " SELECT 1, variedades.ctacomtercero as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Else
                    sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                End If
            Else
                If vSocio.TipoProd = 1 Then
                    sql = " SELECT 1, variedades.ctacomtercero as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                Else
                    sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                End If
            End If
            
            '[Monica]14/11/2014: solo en el caso de Catadau en liquidacion de industria cogemos la ctacomtercero
            If vParamAplic.Cooperativa = 0 Then
                If vEmpresa.TieneAnalitica Then
                    sql = " SELECT 1, variedades.ctacomtercero as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Else
                    sql = " SELECT 1, variedades.ctacomtercero as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                End If
            End If
            
            sql = sql & " FROM rfactsoc_variedad, variedades "
            sql = sql & " WHERE " & Replace(cadwhere, "rfactsoc", "rfactsoc_variedad") & " and"
            sql = sql & " rfactsoc_variedad.codvarie = variedades.codvarie "
            sql = sql & " group by 1,2 "
            sql = sql & " order by 1,2 "
            
        Else
            InsertarLinFactSoc = False
            Exit Function
        End If
    Else
    ' fin de lo añadido
    
        If vEmpresa.TieneAnalitica Then
            Select Case Tipo
                Case 1, 3, 7, 9
                     sql = " SELECT 1, variedades.ctaanticipo as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Case 2, 4, 8, 10
                     sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Case 6 ' siniestros
                     sql = " SELECT 1, variedades.ctasiniestros as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
            End Select
            If TipoFact = "FTS" Then
                sql = " SELECT 1, variedades.ctaacarecol as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
            End If
        Else
            Select Case Tipo
                Case 1, 3, 7, 9
                     sql = " SELECT 1, variedades.ctaanticipo as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                Case 2, 4, 8, 10
                     sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                Case 6 ' siniestros
                     sql = " SELECT 1, variedades.ctasiniestros as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
            End Select
            '[Monica]16/07/2014: añadido el caso de tipo transporte tercero de Picassent
            If TipoFact = "FTS" Or TipoFact = "FTT" Then
                sql = " SELECT 1, variedades.ctaacarecol as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
            End If
        End If
        sql = sql & " FROM rfactsoc_variedad, variedades "
        sql = sql & " WHERE " & Replace(cadwhere, "rfactsoc", "rfactsoc_variedad") & " and"
        sql = sql & " rfactsoc_variedad.codvarie = variedades.codvarie "
        sql = sql & " group by 1,2 "
        sql = sql & " order by 1,2 "

    End If



    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = Cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        ImpLinea = Rs!Importe
        
        ' si se trata de una liquidacion hemos de descontar los anticipos por variedad
        ' que no sean anticipo de gastos
        If (Tipo = 2 Or Tipo = 4 Or Tipo = 8 Or Tipo = 10) And TipoFact <> "FTS" Then
            Sql3 = "select sum(baseimpo) from rfactsoc_anticipos, variedades "
            Sql3 = Sql3 & " where " & Replace(cadwhere, "rfactsoc", "rfactsoc_anticipos")
            Sql3 = Sql3 & " and rfactsoc_anticipos.codvarieanti = variedades.codvarie "
            Sql3 = Sql3 & " and variedades.ctaliquidacion = " & DBSet(Rs!cuenta, "N")
            
            ImpAnticipo = DevuelveValor(Sql3)
            
            ImpLinea = ImpLinea - ImpAnticipo
        End If
        '----
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        sql = ""
        Sql2 = ""
        
        sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        sql = sql & DBSet(Rs!cuenta, "T")
        
        Sql2 = sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        sql = sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
                If DBLet(Rs!CodCCost, "T") = "----" Then
                    sql = sql & DBSet(CCoste, "T")
                Else
                    sql = sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBLet(Rs!CodCCost, "T")
                End If
        Else
            sql = sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        Cad = Cad & "(" & sql & ")" & ","
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
            Cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            Cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If

    ' las retenciones si las hay
    If ImpReten <> 0 Then
        sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        sql = sql & DBSet(CtaSocio, "T")
        sql = sql & "," & DBSet(ImpReten, "N") & ","
        sql = sql & ValorNulo ' no llevan centro de coste
        
        Cad = Cad & "(" & sql & ")" & ","
        I = I + 1
        
        sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        sql = sql & DBSet(CtaReten, "T")
        sql = sql & "," & DBSet(ImpReten * (-1), "N") & ","
        sql = sql & ValorNulo ' no llevan centro de coste
        
        Cad = Cad & "(" & sql & ")" & ","
        I = I + 1
    End If
    
    ' las aportaciones de fondo operativo si las hay
    If ImpAport <> 0 Then
        sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        sql = sql & DBSet(CtaSocio, "T")
        sql = sql & "," & DBSet(ImpAport, "N") & ","
        sql = sql & ValorNulo ' no llevan centro de coste
        
        Cad = Cad & "(" & sql & ")" & ","
        I = I + 1
    
        sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        sql = sql & DBSet(CtaAport, "T")
        sql = sql & "," & DBSet(ImpAport * (-1), "N") & ","
        sql = sql & ValorNulo ' no llevan centro de coste
        
        Cad = Cad & "(" & sql & ")" & ","
        I = I + 1
    End If
    
    '[Monica]20/12/2013: si es montifrut no descontamos el descuento que tengo grabado a pie
        '[Monica]09/03/2015: para el caso de Catadau tampoco se tienen que insertar las bases correspondientes a gastos
    If vParamAplic.Cooperativa <> 12 And vParamAplic.Cooperativa <> 0 Then
        ' insertamos todos los gastos a pie de factura rfactsoc_gastos
        sql = " SELECT rconcepgasto.codmacta as cuenta, sum(rfactsoc_gastos.importe) as importe "
        sql = sql & " from rconcepgasto INNER JOIN rfactsoc_gastos ON rconcepgasto.codgasto = rfactsoc_gastos.codgasto "
        sql = sql & " where " & Replace(cadwhere, "rfactsoc", "rfactsoc_gastos")
        sql = sql & " group by 1 "
        sql = sql & " order by 1 "
        
        Set Rs = New ADODB.Recordset
        Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        While Not Rs.EOF
            sql = NumRegis & "," & AnyoFacPr & "," & I & ","
            sql = sql & DBSet(CtaSocio, "T")
            sql = sql & "," & DBSet(Rs!Importe, "N") & ","
            sql = sql & ValorNulo ' no llevan centro de coste
            
            Cad = Cad & "(" & sql & ")" & ","
            I = I + 1
        
            sql = NumRegis & "," & AnyoFacPr & "," & I & ","
            sql = sql & DBSet(Rs!cuenta, "T")
            sql = sql & "," & DBSet(Rs!Importe * (-1), "N") & ","
            sql = sql & ValorNulo ' no llevan centro de coste
            
            Cad = Cad & "(" & sql & ")" & ","
            I = I + 1
        
            Rs.MoveNext
        Wend
        Set Rs = Nothing
    End If
    
    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        sql = sql & " VALUES " & Cad
        ConnConta.Execute sql
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







Private Function ActualizarCabFact(cadTabla As String, cadwhere As String, cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim sql As String

    On Error GoTo EActualizar
    
    Select Case cadTabla
        Case "rrecibpozos"
    
            sql = "UPDATE " & cadTabla & " SET contabilizado=1 "
            
        Case Else
            sql = "UPDATE " & cadTabla & " SET intconta=1"
            
    End Select
    sql = sql & " WHERE " & cadwhere

    conn.Execute sql
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCabFact = False
        cadErr = Err.Description
    Else
        ActualizarCabFact = True
    End If
End Function



Private Function ActualizarCabFactSoc(cadTabla As String, cadwhere As String, cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim sql As String

    On Error GoTo EActualizar
    
        
    sql = "UPDATE " & cadTabla & " SET contabilizado=1 "
    sql = sql & " WHERE " & cadwhere

    conn.Execute sql
    
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

Public Function PasarFacturaSoc(cadwhere As String, CodCCost As String, FechaFin As String, Seccion As String, TipoFact As Byte, FecRecep As Date, FecVto As Date, ForpaPos As String, ForpaNeg As String, CtaBanc As String, CtaRete As String, CtaApor As String, TipoM As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariagro.rfactsoc --> conta.cabfactprov
' ariagro.rfactsoc_variedad --> conta.linfactprov
'Actualizar la tabla ariagro.rfactsoc.contabilizada=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim sql As String
Dim Mc As Contadores
Dim Obs As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
        
    Set Mc = New Contadores
        
    '[Monica]29/04/2011: INTERNAS
    If EsFacturaInterna(cadwhere) Then
        CtaReten = CtaRete
        CtaAport = CtaApor
        ' Insertamos en el diario
        b = InsertarAsientoDiario(cadwhere, cadMen, Mc, CDate(FechaFin), Seccion, TipoFact, FecRecep, TipoM)
        cadMen = "Insertando Factura en Diario: " & cadMen
    Else
       '---- Insertar en la conta Cabecera Factura
        b = InsertarCabFactSoc(cadwhere, cadMen, Mc, CDate(FechaFin), Seccion, TipoFact, FecRecep, TipoM)
        cadMen = "Insertando Cab. Factura: " & cadMen
    End If
    
    If b Then
        FecVenci = FecVto
        ForpaPosi = ForpaPos
        ForpaNega = ForpaNeg
        CtaBanco = CtaBanc
        CtaReten = CtaRete
        CtaAport = CtaApor
        tipoMov = TipoM    ' codtipom de la factura de socio
        
        '[Monica]09/05/2013: si la cooperativa es Montifrut, las formas de pago estan en la propia factura
        If vParamAplic.Cooperativa = 12 Then
            ForpaPosi = DevuelveValor("select codforpa from rfactsoc where " & cadwhere)
            ForpaNega = ForpaPosi
        End If
        
'01-06-2009
        b = InsertarEnTesoreriaSoc(cadwhere, cadMen, FacturaSoc, FecFactuSoc)
        cadMen = "Insertando en Tesoreria: " & cadMen

        If b Then
            CCoste = CodCCost
            '[Monica]29/04/2011: INTERNAS
            If Not EsFacturaInterna(cadwhere) Then
                '---- Insertar lineas de Factura en la Conta
                b = InsertarLinFactSoc("rfactsoc", cadwhere, cadMen, TipoFact, Mc.Contador)
                cadMen = "Insertando Lin. Factura: " & cadMen
            End If
            
            If b Then
                '---- Poner intconta=1 en ariges.scafac
                b = ActualizarCabFactSoc("rfactsoc", cadwhere, cadMen)
                cadMen = "Actualizando Factura Socio: " & cadMen
            End If
        End If
    End If
    
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura Socio", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaSoc = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaSoc = False
        If Not b Then
            InsertarTMPErrFacSoc cadMen, cadwhere
        End If
    End If
End Function


Private Function InsertarCabFactProv(cadwhere As String, cadErr As String, ByRef Mc As Contadores, FechaFin As String) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String

    On Error GoTo EInsertar
       
    
    sql = " SELECT fecfactu,year(fecrecep) as anofacpr,fecrecep,numfactu,proveedor.codmacta,"
    sql = sql & "scafpc.dtoppago,scafpc.dtognral,baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    sql = sql & "totalfac,tipoiva1,tipoiva2,tipoiva3,proveedor.codprove, proveedor.nomprove "
    sql = sql & " FROM " & "scafpc "
    sql = sql & "INNER JOIN " & "proveedor ON scafpc.codprove=proveedor.codprove "
    sql = sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
    
        If Mc.ConseguirContador("1", (Rs!FecRecep <= CDate(FechaFin) - 365), True) = 0 Then
        
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            DtoPPago = Rs!DtoPPago
            DtoGnral = Rs!DtoGnral
            BaseImp = Rs!BaseIVA1 + CCur(DBLet(Rs!BaseIVA2, "N")) + CCur(DBLet(Rs!BaseIVA3, "N"))
            TotalFac = Rs!TotalFac
            AnyoFacPr = Rs!anofacpr
            
            Nulo2 = "N"
            Nulo3 = "N"
            If DBLet(Rs!BaseIVA2, "N") = "0" Then Nulo2 = "S"
            If DBLet(Rs!BaseIVA3, "N") = "0" Then Nulo3 = "S"
            sql = ""
            sql = Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & Rs!anofacpr & "," & DBSet(Rs!FecRecep, "F") & "," & DBSet(Rs!numfactu, "T") & "," & DBSet(Rs!Codmacta, "T") & "," & ValorNulo & ","
            sql = sql & DBSet(Rs!BaseIVA1, "N") & "," & DBSet(Rs!BaseIVA2, "N", "S") & "," & DBSet(Rs!BaseIVA3, "N", "S") & ","
            sql = sql & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", Nulo2) & "," & DBSet(Rs!porciva3, "N", Nulo3) & ","
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!ImpoIva1, "N") & "," & DBSet(Rs!impoIVA2, "N", Nulo2) & "," & DBSet(Rs!impoIVA3, "N", Nulo3) & ","
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!TipoIVA1, "N") & "," & DBSet(Rs!TipoIVA2, "N", Nulo2) & "," & DBSet(Rs!TipoIVA3, "N", Nulo3) & ",0,"
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!FecRecep, "F") & ",0"
            Cad = Cad & "(" & sql & ")"
            
            'Insertar en la contabilidad
            sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
            sql = sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
            sql = sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
            sql = sql & " VALUES " & Cad
            ConnConta.Execute sql
            
            'añadido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            sql = sql & ",'" & DevNombreSQL(Rs!numfactu) & " @ " & Format(Rs!fecfactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!nomprove) & "'," & Rs!codProve & ")"
            conn.Execute sql
            
            
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactProv = False
        cadErr = Err.Description
    Else
        InsertarCabFactProv = True
    End If
End Function


Private Function InsertarCabFactSoc(cadwhere As String, cadErr As String, ByRef Mc As Contadores, FechaFin As Date, Secci As String, Tipo As Byte, FecRec As Date, TipoM As String) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Concepto As String
Dim LetraSer As String
Dim Socio As String

    On Error GoTo EInsertar
    
    '[Monica]09/05/2013: en el caso de Montifrut cuando contabilizo la fecha de recepcion va a ser la fecha de factura
    If vParamAplic.Cooperativa = 12 Then
        sql = " SELECT codtipom, fecfactu,year(fecfactu) as anofacpr,fecfactu,numfactu,rsocios_seccion.codmacpro,"
    Else
        sql = " SELECT codtipom, fecfactu,year(" & DBSet(FecRec, "F") & ") as anofacpr," & DBSet(FecRec, "F") & ",numfactu,rsocios_seccion.codmacpro,"
    End If
    
    sql = sql & "baseimpo, tipoiva, porc_iva,imporiva,basereten,porc_ret,impreten,baseaport,porc_apo,impapor,totalfac,"
    sql = sql & "rectif_codtipom, rectif_numfactu, rectif_fecfactu,"
    sql = sql & "rsocios.codsocio, rsocios.nomsocio, rsocios.codbanco, rsocios.codsucur, rsocios.digcontr, rsocios.cuentaba, rsocios.iban "
    '[Monica]27/01/2012: Si han introducido el nro de factura recibido es el que hay que llevar a conta
    sql = sql & ", rfactsoc.numfacrec "
    
    sql = sql & " FROM (" & "rfactsoc "
    sql = sql & "INNER JOIN rsocios ON rfactsoc.codsocio=rsocios.codsocio) "
    sql = sql & " INNER JOIN rsocios_seccion ON rfactsoc.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Secci, "N")
    sql = sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
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
                LetraSer = ""
                LetraSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(TipoM, "T"))
            
                FacturaSoc = LetraSer & "-" & DBLet(Rs!numfactu, "N")
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
            
            
            Variedades = VariedadesFactura(cadwhere)
            
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
            
            sql = ""
            sql = Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & AnyoFacPr & "," & DBSet(FecRecep, "F") & "," & DBSet(FacturaSoc, "T") & "," & DBSet(CtaSocio, "T") & "," & DBSet(Concepto, "T") & ","
            sql = sql & DBSet(BaseImp, "N") & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & DBSet(Rs!porc_iva, "N") & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!ImporIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & DBSet(TotalFac, "N") & "," & DBSet(Rs!TipoIVA, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(FecRecep, "F") & ",0"
            Cad = Cad & "(" & sql & ")"
            
            'Insertar en la contabilidad
            sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
            sql = sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
            sql = sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
            sql = sql & " VALUES " & Cad
            ConnConta.Execute sql
            
            'añadido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            sql = sql & ",'" & DevNombreSQL(FacturaSoc) & " @ " & Format(Rs!fecfactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!nomsocio) & "'," & Rs!Codsocio & ")"
            conn.Execute sql

            FacturaSoc = DBLet(Rs!numfactu, "N")
            
        Else
        
            If Mc.ConseguirContador("1", (FecRec <= CDate(FechaFin) - 365), True) = 0 Then
            
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
                    LetraSer = ""
                    LetraSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(TipoM, "T"))
                
                    FacturaSoc = LetraSer & "-" & DBLet(Rs!numfactu, "N")
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
                   sql = "select rsocios.tiporelacion from rsocios where codsocio in (select codsocio from rfactsoc where " & cadwhere & ")"
                   If DevuelveValor(sql) = 1 Then
                       
                       sql = "select codsocio from rfactsoc where " & cadwhere
                       Socio = DevuelveValor(sql)
                       
                       sql = "select replace(codmacpro,mid(codmacpro,1,length(rseccion.raiz_cliente_socio)), rseccion.raiz_cliente_asociado) "
                       sql = sql & " from (rsocios_seccion inner join rseccion on rsocios_seccion.codsecci = rseccion.codsecci) inner join rsocios on rsocios_seccion .codsocio = rsocios.codsocio "
                       sql = sql & " where rsocios_seccion.codsocio = " & DBSet(Socio, "N")
                       sql = sql & " and rseccion.codsecci = " & DBSet(Seccion, "N")
    
                       CtaSocio = DevuelveValor(sql)
                   End If
                End If
                
                
                Variedades = VariedadesFactura(cadwhere)
                
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
                
                sql = ""
                sql = Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & AnyoFacPr & "," & DBSet(FecRecep, "F") & "," & DBSet(FacturaSoc, "T") & "," & DBSet(CtaSocio, "T") & "," & DBSet(Concepto, "T") & ","
                sql = sql & DBSet(BaseImp, "N") & "," & ValorNulo & "," & ValorNulo & ","
                sql = sql & DBSet(Rs!porc_iva, "N") & "," & ValorNulo & "," & ValorNulo & ","
                sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!ImporIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
                sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                sql = sql & DBSet(TotalFac, "N") & "," & DBSet(Rs!TipoIVA, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
                sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(FecRecep, "F") & ",0"
                Cad = Cad & "(" & sql & ")"
                
                'Insertar en la contabilidad
                sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
                sql = sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
                sql = sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
                sql = sql & " VALUES " & Cad
                ConnConta.Execute sql
                
                'añadido como david para saber que numero de registro corresponde a cada factura
                'Para saber el numreo de registro que le asigna a la factrua
                sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
                sql = sql & ",'" & DevNombreSQL(FacturaSoc) & " @ " & Format(Rs!fecfactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!nomsocio) & "'," & Rs!Codsocio & ")"
                conn.Execute sql
    
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



Private Function InsertarAsientoDiario(cadwhere As String, cadErr As String, ByRef Mc As Contadores, FechaFin As Date, Secci As String, Tipo As Byte, FecRec As Date, TipoM As String) As Boolean
' la contabilizacion de las facturas internas de horto, se han de insertar en el diario no en el registro de iva de proveedor
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Concepto As String
Dim LetraSer As String
Dim Obs As String
Dim cadMen As String
Dim b As Boolean
'Dim CtaSocio As String


    On Error GoTo EInsertar
       
    sql = " SELECT fecfactu,year(" & DBSet(FecRec, "F") & ") as anofacpr," & DBSet(FecRec, "F") & ",numfactu,rsocios_seccion.codmacpro,"
    sql = sql & "baseimpo, tipoiva, porc_iva,imporiva,basereten,porc_ret,impreten,baseaport,porc_apo,impapor,totalfac,"
    sql = sql & "rectif_codtipom, rectif_numfactu, rectif_fecfactu,"
    sql = sql & "rsocios.codsocio, rsocios.nomsocio, rsocios.codbanco, rsocios.codsucur, rsocios.digcontr, rsocios.cuentaba "
    sql = sql & " FROM (" & "rfactsoc "
    sql = sql & "INNER JOIN rsocios ON rfactsoc.codsocio=rsocios.codsocio) "
    sql = sql & " INNER JOIN rsocios_seccion ON rfactsoc.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Secci, "N")
    sql = sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        If Mc.ConseguirContador("1", (FecRec <= CDate(FechaFin) - 365), True) = 0 Then
        
            BaseImp = DBLet(Rs!baseimpo, "N")
            TotalFac = BaseImp + DBLet(Rs!ImporIva, "N")
            AnyoFacPr = Rs!anofacpr
            
            ImpReten = DBLet(Rs!ImpReten, "N")
            ImpAport = DBLet(Rs!impapor, "N")
            
            LetraSer = ""
            LetraSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(TipoM, "T"))
            
            FacturaSoc = LetraSer & "-" & DBLet(Rs!numfactu, "N")
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
               sql = "select rsocios.tiporelacion from rsocios where codsocio in (select codsocio from rfactsoc where " & cadwhere & ")"
               If DevuelveValor(sql) = 1 Then
                   
                   sql = "select codsocio from rfactsoc where " & cadwhere
                   Socio = DevuelveValor(sql)
                   
                   sql = "select replace(codmacpro,mid(codmacpro,1,length(rseccion.raiz_cliente_socio)), rseccion.raiz_cliente_asociado) "
                   sql = sql & " from (rsocios_seccion inner join rseccion on rsocios_seccion.codsecci = rseccion.codsecci) inner join rsocios on rsocios_seccion .codsocio = rsocios.codsocio "
                   sql = sql & " where rsocios_seccion.codsocio = " & DBSet(Socio, "N")
                   sql = sql & " and rseccion.codsecci = " & DBSet(Seccion, "N")

                   CtaSocio = DevuelveValor(sql)
               End If
            End If
            
            Variedades = VariedadesFactura(cadwhere)
            
            Obs = "Contabilización Factura Interna de Fecha " & Format(FecRecep, "dd/mm/yyyy")
        
            'Insertar en la conta Cabecera Asiento
            b = InsertarCabAsientoDia(vEmpresa.NumDiarioInt, Mc.Contador, CStr(Format(FecRecep, "dd/mm/yyyy")), Obs, cadMen)
            cadMen = "Insertando Cab. Asiento: " & cadMen
    
            If b Then
                Socio = DevuelveValor("select codsocio from rfactsoc where " & cadwhere)
'                CtaSocio = DevuelveValor("select codmacpro from rsocios_seccion where codsocio = " & Socio & " and codsecci = " & vParamAplic.SeccionHorto)
            
                b = InsertarLinAsientoFactIntProv("rfactsoc", cadwhere, cadMen, Tipo, CtaSocio, Mc.Contador)
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



Private Function InsertarLinAsientoFactIntProv(cadTabla As String, cadwhere As String, cadErr As String, Tipo As Byte, CtaSocio As String, Optional Contador As Long) As Boolean
Dim sql As String
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
Dim b As Boolean
Dim Cad As String
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
        sql = "select mid(rfactsoc.codtipom,1,3) from " & cadTabla & " where " & cadwhere
        TipoFact = DevuelveValor(sql)
    
    End If
    
    FeFact = FecFactuSoc
    NumFact = DevuelveValor("select numfactu from rfactsoc where " & cadwhere)
    
    If Tipo = 2 And TipoFact = "FLI" Then
        sql = "select rfactsoc.codsocio from " & cadTabla & " where " & cadwhere
        Socio = DevuelveValor(sql)
        
        Set vSocio = New cSocio
        If vSocio.LeerDatos(Socio) Then
            If vEmpresa.TieneAnalitica Then
                If vSocio.TipoProd = 1 Then
                    sql = " SELECT 1, variedades.ctacomtercero as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Else
                    sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                End If
            Else
                If vSocio.TipoProd = 1 Then
                    sql = " SELECT 1, variedades.ctacomtercero as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                Else
                    sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                End If
            End If
            sql = sql & " FROM rfactsoc_variedad, variedades "
            sql = sql & " WHERE " & Replace(cadwhere, "rfactsoc", "rfactsoc_variedad") & " and"
            sql = sql & " rfactsoc_variedad.codvarie = variedades.codvarie "
            sql = sql & " group by 1,2 "
            sql = sql & " order by 1,2 "
            
        Else
            InsertarLinAsientoFactIntProv = False
            Exit Function
        End If
    Else
    ' fin de lo añadido
    
        If vEmpresa.TieneAnalitica Then
            Select Case Tipo
                Case 1, 3, 7, 9
                     sql = " SELECT 1, variedades.ctaanticipo as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Case 2, 4, 8, 10
                     sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
                Case 6 ' siniestros
                     sql = " SELECT 1, variedades.ctasiniestros as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
            End Select
            If TipoFact = "FTS" Then
                sql = " SELECT 1, variedades.ctaacarecol as cuenta, sum(rfactsoc_variedad.imporvar) as importe, variedades.codccost "
            End If
        Else
            Select Case Tipo
                Case 1, 3, 7, 9
                     sql = " SELECT 1, variedades.ctaanticipo as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                Case 2, 4, 8, 10
                     sql = " SELECT 1, variedades.ctaliquidacion as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
                Case 6 ' siniestros
                     sql = " SELECT 1, variedades.ctasiniestros as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
            End Select
            If TipoFact = "FTS" Then
                sql = " SELECT 1, variedades.ctaacarecol as cuenta, sum(rfactsoc_variedad.imporvar) as importe "
            End If
        End If
        sql = sql & " FROM rfactsoc_variedad, variedades "
        sql = sql & " WHERE " & Replace(cadwhere, "rfactsoc", "rfactsoc_variedad") & " and"
        sql = sql & " rfactsoc_variedad.codvarie = variedades.codvarie "
        sql = sql & " group by 1,2 "
        sql = sql & " order by 1,2 "

    End If

    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    I = 0
    ImporteD = 0
    ImporteH = 0
    
    numdocum = Format(NumFact, "0000000")
    Ampliacion = FacturaSoc  'TipoFact & "-" & Format(NumFact, "0000000")
    ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoIntProv, "N")) & " " & Ampliacion
    ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoIntProv, "N")) & " " & Ampliacion
    
    If Not Rs.EOF Then Rs.MoveFirst
    
    b = True

    Cad = ""
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = Cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        ImpLinea = Rs!Importe
        
        ' si se trata de una liquidacion hemos de descontar los anticipos por variedad
        ' que no sean anticipo de gastos
        If (Tipo = 2 Or Tipo = 4 Or Tipo = 8 Or Tipo = 10) And TipoFact <> "FTS" Then
            Sql3 = "select sum(baseimpo) from rfactsoc_anticipos, variedades "
            Sql3 = Sql3 & " where " & Replace(cadwhere, "rfactsoc", "rfactsoc_anticipos")
            Sql3 = Sql3 & " and rfactsoc_anticipos.codvarieanti = variedades.codvarie "
            Sql3 = Sql3 & " and variedades.ctaliquidacion = " & DBSet(Rs!cuenta, "N")
            
            ImpAnticipo = DevuelveValor(Sql3)
            
            ImpLinea = ImpLinea - ImpAnticipo
        End If
        '----
        totimp = totimp + ImpLinea
        
        I = I + 1
        
        Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
        Cad = Cad & DBSet(I, "N") & "," & DBSet(Rs!cuenta, "T") & "," & DBSet(numdocum, "T") & ","
        
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If Rs.Fields(2).Value > 0 Then
            ' importe al debe en positivo
            Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Rs.Fields(2).Value, "N") & ","
            Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
        
            ImporteD = ImporteD + (CCur(Rs.Fields(2).Value))
        Else
            ' importe al haber en positivo, cambiamos el signo
            Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
            Cad = Cad & DBSet((Rs.Fields(2).Value) * (-1), "N") & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
        
            ImporteH = ImporteH + (CCur(Rs.Fields(2).Value) * (-1))
        End If
        
        Cad = "(" & Cad & ")"
        
        b = InsertarLinAsientoDia(Cad, cadMen)
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
            sql = "update linapu set timporteD = " & DBSet(totimp, "N")
            sql = sql & " where numdiari = " & DBSet(vEmpresa.NumDiarioInt, "N")
            sql = sql & " and fechaent = " & DBSet(FecRecep, "F")
            sql = sql & " and numasien = " & DBSet(Contador, "N")
            sql = sql & " and linliapu = " & DBSet(I, "N")
            
            ConnConta.Execute sql
        Else
            sql = "update linapu set timporteH = " & DBSet(totimp, "N")
            sql = sql & " where numdiari = " & DBSet(vEmpresa.NumDiarioInt, "N")
            sql = sql & " and fechaent = " & DBSet(FecRecep, "F")
            sql = sql & " and numasien = " & DBSet(Contador, "N")
            sql = sql & " and linliapu = " & DBSet(I, "N")
            
            ConnConta.Execute sql
        End If
    End If

    If b And I > 0 Then
        I = I + 1
        
        ' el Total es sobre la cuenta del socio
        Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
        Cad = Cad & DBSet(I, "N") & ","
        Cad = Cad & DBSet(CtaSocio, "T") & "," & DBSet(numdocum, "T") & ","
            
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If ImporteD - ImporteH < 0 Then
            ' importe al debe en positivo
            Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet((ImporteD - ImporteH) * (-1), "N") & ","
            Cad = Cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
        Else
            ' importe al haber en positivo, cambiamos el signo
            Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
            Cad = Cad & DBSet((ImporteD - ImporteH), "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
        End If
        
        Cad = "(" & Cad & ")"
        
        b = InsertarLinAsientoDia(Cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & I
        
    End If

    If b Then
        ' las retenciones si las hay
        If ImpReten <> 0 Then
            I = I + 1
            
            Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
            Cad = Cad & DBSet(I, "N") & "," & DBSet(CtaSocio, "T") & "," & DBSet(numdocum, "T") & ","
        
            If ImpReten > 0 Then
                ' importe al debe en positivo
                Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(ImpReten, "N") & ","
                Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaReten, "T") & "," & ValorNulo & ",0"
            Else
                ' importe al haber en positivo, cambiamos el signo
                Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                Cad = Cad & DBSet((ImpReten * (-1)), "N") & "," & ValorNulo & "," & DBSet(CtaReten, "T") & "," & ValorNulo & ",0"
            
            End If
            
            Cad = "(" & Cad & ")"
            
            b = InsertarLinAsientoDia(Cad, cadMen)
            cadMen = "Insertando Lin. Asiento: " & I
            
            If b Then
                I = I + 1
                
                Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
                Cad = Cad & DBSet(I, "N") & "," & DBSet(CtaReten, "T") & "," & DBSet(numdocum, "T") & ","
                If ImpReten > 0 Then
                    ' importe al haber en positivo
                    Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & ValorNulo & ","
                    Cad = Cad & DBSet(ImpReten, "N") & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & DBSet((ImpReten * (-1)), "N") & ","
                    Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
                End If
                
                Cad = "(" & Cad & ")"
                
                b = InsertarLinAsientoDia(Cad, cadMen)
                cadMen = "Insertando Lin. Asiento: " & I
            End If
            
        End If
    End If
    
    
    If b Then
        ' las aportaciones de fondo operativo si las hay
        If ImpAport <> 0 Then
            I = I + 1
            
            Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
            Cad = Cad & DBSet(I, "N") & "," & DBSet(CtaSocio, "T") & "," & DBSet(numdocum, "T") & ","
        
            If ImpAport > 0 Then
                ' importe al debe en positivo
                Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(ImpAport, "N") & ","
                Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaAport, "T") & "," & ValorNulo & ",0"
            Else
                ' importe al haber en positivo, cambiamos el signo
                Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                Cad = Cad & DBSet((ImpAport * (-1)), "N") & "," & ValorNulo & "," & DBSet(CtaAport, "T") & "," & ValorNulo & ",0"
            
            End If
            
            Cad = "(" & Cad & ")"
            
            b = InsertarLinAsientoDia(Cad, cadMen)
            cadMen = "Insertando Lin. Asiento: " & I
            
            If b Then
                I = I + 1
                
                Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
                Cad = Cad & DBSet(I, "N") & "," & DBSet(CtaAport, "T") & "," & DBSet(numdocum, "T") & ","
                If ImpAport > 0 Then
                    ' importe al haber en positivo
                    Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & ValorNulo & ","
                    Cad = Cad & DBSet(ImpAport, "N") & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & DBSet((ImpAport * (-1)), "N") & ","
                    Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
                End If
                
                Cad = "(" & Cad & ")"
                
                b = InsertarLinAsientoDia(Cad, cadMen)
                cadMen = "Insertando Lin. Asiento: " & I
            End If
        End If
    End If
    
    '[Monica]09/03/2015: para el caso de Catadau no hay apuntes de gastos, añadida la condicion de cooperativa
    If b And vParamAplic.Cooperativa <> 0 Then
        ' insertamos todos los gastos a pie de factura rfactsoc_gastos
        sql = " SELECT rconcepgasto.codmacta as cuenta, sum(rfactsoc_gastos.importe) as importe "
        sql = sql & " from rconcepgasto INNER JOIN rfactsoc_gastos ON rconcepgasto.codgasto = rfactsoc_gastos.codgasto "
        sql = sql & " where " & Replace(cadwhere, "rfactsoc", "rfactsoc_gastos")
        sql = sql & " group by 1 "
        sql = sql & " order by 1 "
        
        Set Rs = New ADODB.Recordset
        Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        While Not Rs.EOF And b
            I = I + 1
            
            Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
            Cad = Cad & DBSet(I, "N") & "," & DBSet(CtaSocio, "T") & "," & DBSet(numdocum, "T") & ","
        
            If Rs!Importe > 0 Then
                ' importe al debe en positivo
                Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Rs!Importe, "N") & ","
                Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!cuenta, "T") & "," & ValorNulo & ",0"
            Else
                ' importe al haber en positivo, cambiamos el signo
                Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                Cad = Cad & DBSet((Rs!Importe * (-1)), "N") & "," & ValorNulo & "," & DBSet(Rs!cuenta, "T") & "," & ValorNulo & ",0"
            End If
            
            Cad = "(" & Cad & ")"
            
            b = InsertarLinAsientoDia(Cad, cadMen)
            cadMen = "Insertando Lin. Asiento: " & I
            
            If b Then
                I = I + 1
                
                Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
                Cad = Cad & DBSet(I, "N") & "," & DBSet(Rs!cuenta, "T") & "," & DBSet(numdocum, "T") & ","
                If ImpAport > 0 Then
                    ' importe al haber en positivo
                    Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & ValorNulo & ","
                    Cad = Cad & DBSet(Rs!Importe, "N") & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & DBSet((Rs!Importe * (-1)), "N") & ","
                    Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaSocio, "T") & "," & ValorNulo & ",0"
                End If
            
                Cad = "(" & Cad & ")"
                
                b = InsertarLinAsientoDia(Cad, cadMen)
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
    InsertarLinAsientoFactIntProv = b
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

Public Function PasarFacturaTerc(cadwhere As String, CodCCost As String, FechaFin As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariagro.tcafpc --> conta.cabfactprov
' ariagro.tlifpc --> conta.linfactprov
'Actualizar la tabla ariges.scafpc.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim sql As String
Dim Mc As Contadores


    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
        
    
    Set Mc = New Contadores
    
    '---- Insertar en la conta Cabecera Factura
    b = InsertarCabFactTerc(cadwhere, cadMen, Mc, FechaFin)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
        CCoste = CodCCost
        '---- Insertar lineas de Factura en la Conta
        b = InsertarLinFact_new("rcafter", cadwhere, cadMen, Mc.Contador)
        cadMen = "Insertando Lin. Factura: " & cadMen

        If b Then
            '---- Poner intconta=1 en ariges.scafac
            b = ActualizarCabFact("rcafter", cadwhere, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
'    If Not b Then
'        SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
'        SQL = SQL & " Select *," & DBSet(Mid(cadMen, 1, 200), "T") & " as error From tmpFactu "
'        SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
'        Conn.Execute SQL
'    End If
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaTerc = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaTerc = False
        If Not b Then
            InsertarTMPErrFac cadMen, cadwhere
'            SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
'            SQL = SQL & " Select *," & DBSet(Mid(cadMen, 1, 200), "T") & " as error From tmpFactu "
'            SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
'            Conn.Execute SQL
        End If
    End If
End Function

Private Function InsertarCabFactTerc(cadwhere As String, cadErr As String, ByRef Mc As Contadores, FechaFin As String) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Nulo4 As String

    On Error GoTo EInsertar


    sql = " SELECT fecfactu,year(fecrecep) as anofacpr,fecrecep,numfactu,rsocios_seccion.codmacpro as codmacta,"
    sql = sql & "rcafter.dtoppago,rcafter.dtognral,baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    sql = sql & "totalfac,tipoiva1,tipoiva2,tipoiva3, retfacpr, trefacpr, rsocios_seccion.codsocio, rsocios.nomsocio, rcafter.intracom "
    sql = sql & " FROM (" & "rcafter "
    sql = sql & "INNER JOIN " & "rsocios ON rcafter.codsocio=rsocios.codsocio )"
    sql = sql & " INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.Seccionhorto, "N")
    sql = sql & " WHERE " & cadwhere

    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    If Not Rs.EOF Then

        If Mc.ConseguirContador("1", (Rs!FecRecep <= CDate(FechaFin) - 365), True) = 0 Then

            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            DtoPPago = Rs!DtoPPago
            DtoGnral = Rs!DtoGnral
            BaseImp = Rs!BaseIVA1 + CCur(DBLet(Rs!BaseIVA2, "N")) + CCur(DBLet(Rs!BaseIVA3, "N"))
            TotalFac = Rs!TotalFac
            AnyoFacPr = Rs!anofacpr

            Nulo2 = "N"
            Nulo3 = "N"
            Nulo4 = "N"
            If DBLet(Rs!BaseIVA2, "N") = "0" Then Nulo2 = "S"
            If DBLet(Rs!BaseIVA3, "N") = "0" Then Nulo3 = "S"
            If DBLet(Rs!trefacpr, "N") = "0" Then Nulo4 = "S"
            sql = ""
            sql = Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & Rs!anofacpr & "," & DBSet(Rs!FecRecep, "F") & "," & DBSet(Rs!numfactu, "T") & "," & DBSet(Rs!Codmacta, "T") & "," & ValorNulo & ","
            sql = sql & DBSet(Rs!BaseIVA1, "N") & "," & DBSet(Rs!BaseIVA2, "N", "S") & "," & DBSet(Rs!BaseIVA3, "N", "S") & ","
            sql = sql & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", Nulo2) & "," & DBSet(Rs!porciva3, "N", Nulo3) & ","
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!ImpoIva1, "N") & "," & DBSet(Rs!impoIVA2, "N", Nulo2) & "," & DBSet(Rs!impoIVA3, "N", Nulo3) & ","
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!TipoIVA1, "N") & "," & DBSet(Rs!TipoIVA2, "N", Nulo2) & "," & DBSet(Rs!TipoIVA3, "N", Nulo3) & ","
            sql = sql & DBSet(Rs!intracom, "N") & ","
            sql = sql & DBSet(Rs!retfacpr, "N", Nulo4) & "," & DBSet(Rs!trefacpr, "N", Nulo4) & ","
            If Nulo4 = "S" Then
                sql = sql & ValorNulo & ","
            Else
                sql = sql & DBSet(vParamAplic.CtaTerReten, "T") & ","
            End If
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!FecRecep, "F") & ",0"
            Cad = Cad & "(" & sql & ")"

            'Insertar en la contabilidad
            sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
            sql = sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
            sql = sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
            sql = sql & " VALUES " & Cad
            ConnConta.Execute sql

            'añadido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            sql = sql & ",'" & DevNombreSQL(Rs!numfactu) & " @ " & Format(Rs!fecfactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!nomsocio) & "'," & Rs!Codsocio & ")"
            conn.Execute sql

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

' ### [Monica] 16/01/2008
Public Function InsertarEnTesoreriaNewFac(cadwhere As String, CtaBan As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
Dim b As Boolean
Dim sql As String, Text33csb As String, Text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim DigConta As String
Dim CC As String
Dim vrefer As String
Dim LetraSer As String
Dim Rsx As ADODB.Recordset


    On Error GoTo EInsertarTesoreriaNewFac

    b = False
    InsertarEnTesoreriaNewFac = False
    CadValues = ""
    CadValues2 = ""
    
    Set Rsx = New ADODB.Recordset
    sql = "select * from facturas where " & cadwhere
    Rsx.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    If Not Rsx.EOF Then
    
        Sql4 = "select codbanco, codsucur, digcontr, cuentaba, codmacta, iban from clientes where codclien = " & DBLet(Rsx!CodClien, "N")
        Set Rs4 = New ADODB.Recordset
        
        Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs4.EOF Then
            LetraSer = ""
            LetraSer = DevuelveDesdeBDNew(cAgro, "stipom", "letraser", "codtipom", Rsx!CodTipom, "T")
            
            Text33csb = "'Factura:" & DBLet(LetraSer, "T") & "-" & DBLet(Rsx!numfactu, "T") & " " & Format(DBLet(Rsx!fecfactu, "F"), "dd/mm/yy") & "'"
            Text41csb = "de " & DBSet(Rsx!TotalFac, "N")
                  
            CC = DBLet(Rs4!digcontr, "T")
            If DBLet(Rs4!digcontr, "T") = "**" Then CC = "00"
        
            '[Monica]03/07/2013: añado trim(codmacta)
            CadValuesAux2 = "(" & DBSet(LetraSer, "T") & "," & DBSet(Rsx!numfactu, "N") & "," & DBSet(Rsx!fecfactu, "F") & ", 1," & DBSet(Trim(Rs4!Codmacta), "T") & ","
            CadValues2 = CadValuesAux2 & DBSet(Rsx!codforpa, "N") & "," & DBSet(Rsx!fecfactu, "F") & "," & DBSet(Rsx!TotalFac, "N") & ","
            CadValues2 = CadValues2 & DBSet(CtaBan, "T") & "," & DBSet(Rs4!CodBanco, "N", "S") & "," & DBSet(Rs4!CodSucur, "N", "S") & ","
            CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(Rs4!CuentaBa, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & ",1" ')"
            '[Monica]22/11/2013: Tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                CadValues2 = CadValues2 & "," & DBSet(Rs4!Iban, "T", "S") & ") "
            Else
                CadValues2 = CadValues2 & ") "
            End If
            
            
            'Insertamos en la tabla scobro de la CONTA
            sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
            sql = sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
            sql = sql & " text33csb, text41csb, agente" ') "
            '[Monica]22/11/2013: Tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                sql = sql & ", iban) "
            Else
                sql = sql & ") "
            End If
            
            
            sql = sql & " VALUES " & CadValues2
            ConnConta.Execute sql
    
        End If
    
        b = True
    End If
    
EInsertarTesoreriaNewFac:
    If Err.Number <> 0 Then
        b = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaNewFac = b
End Function

Public Function InsertarEnTesoreriaSoc(cadwhere As String, MenError As String, numfactu As String, fecfactu As Date) As Boolean
'Guarda datos de Tesoreria en tablas: spagop o scobro dependiendo del signo de la factura
Dim b As Boolean
Dim sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim CC As String
Dim LetraSer As String
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

    On Error GoTo EInsertarTesoreriaSoc

    InsertarEnTesoreriaSoc = False
    
    '[Monica] 21/01/2010 tenemos que descontar del totaltesor los gastos a pie de factura
    sql = "select sum(importe) from rfactsoc_gastos where " & Replace(cadwhere, "rfactsoc", "rfactsoc_gastos")
    GastosPie = DevuelveValor(sql)
    '[Monica]29/11/2013: si es Montifrut los gastos a pie no se descuentan del importe
    If vParamAplic.Cooperativa = 12 Then GastosPie = 0
    
    
    '[Monica] 13/06/2013 tenemos que descontar las facturas varias que se insertaron
    sql = "select sum(totalfac) from fvarcabfact where (codsecci, codtipom, numfactu, fecfactu) in (select codsecci, codtipomfvar, numfactufvar, fecfactufvar from rfactsoc_fvarias where " & Replace(cadwhere, "rfactsoc", "rfactsoc_fvarias") & ")"
    GastosVarias = DevuelveValor(sql)
    
    
    '[Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay
    FactuRec = DevuelveValor("select numfacrec from rfactsoc where " & cadwhere)
    If FactuRec = "0" Then
        LetraSer = ""
        LetraSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(tipoMov, "T"))
    
        FactuRec = LetraSer & "-" & numfactu
    End If
    
    vBancoSoc = ""
    If BancoSoc <> 0 Then vBancoSoc = BancoSoc
    
    vSucurSoc = ""
    If SucurSoc <> 0 Then vSucurSoc = SucurSoc
    
    
    TotalTesor = TotalTesor - GastosPie - GastosVarias
    
    
    'si hay porcentaje de corredor hemos de descontarlo tb. Este porcentaje lo cargaba Montifrut
    sql = "select porccorredor from rfactsoc where " & cadwhere
    PorcCorredor = DevuelveValor(sql)
    
    TotalTesor1 = Round2(TotalTesor * PorcCorredor / 100, 2)
    TotalTesor = TotalTesor - Round2(TotalTesor * PorcCorredor / 100, 2)
    
    If TotalTesor > 0 Then ' se insertara en la cartera de pagos (spagop)
        
        '[Monica]09/05/2013: Añadido el nro de vencimientos
        CadValues2 = ""
        
        sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & ForpaPosi
        Set rsVenci = New ADODB.Recordset
        rsVenci.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not rsVenci.EOF Then
            If rsVenci!numerove > 0 Then
                
                'vamos creando la cadena para insertar en spagosp de la CONTA
                LetraSer = ""
                LetraSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(tipoMov, "T"))
                
                'Obtener los dias de pago de la tabla de parametros: spara1
                    
                    '[Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
                    '[Monica]03/07/2013: añado trim(codmacta)
                    CadValuesAux2 = "('" & Trim(CtaSocio) & "', " & DBSet(FactuRec, "T") & ", '" & Format(fecfactu, FormatoFecha) & "', "
                    
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
                
                      'David. Para que ponga la cuenta bancaria (SI LA tiene)
                      CadValues2 = CadValues2 & DBSet(vBancoSoc, "T", "S") & "," & DBSet(vSucurSoc, "T", "S") & ","
                      CadValues2 = CadValues2 & DBSet(DigcoSoc, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & ","
                
                      'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
                      Select Case TipoFact
                        Case 1, 3, 7, 9 ' anticipo y anticipo venta campo
            '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
            '                Sql = "Anticipo num.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                            sql = "Anticipo num.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                        Case 2, 4, 8, 10 ' liquidacion y liquidacion venta campo
            '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
            '                Sql = "Liquidacion num.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                            sql = "Liquidacion num.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                        Case Else
            '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
            '                Sql = "Fact.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                            sql = "Fact.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                      End Select
                        
                      CadValues2 = CadValues2 & "'" & DevNombreSQL(sql) & "',"
                    
                      sql = "Variedades: " & Variedades
                      CadValues2 = CadValues2 & "'" & DevNombreSQL(sql) & "'" '),"
                      '[Monica]22/11/2013: Tema iban
                      If vEmpresa.HayNorma19_34Nueva = 1 Then
                          CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & "),"
                      Else
                          CadValues2 = CadValues2 & "),"
                      End If
        
                      'Resto Vencimientos
                      '--------------------------------------------------------------------
                      UltimoVto = 1
                      For I = 2 To rsVenci!numerove
                          UltimoVto = I
                         'FECHA Resto Vencimientos
                          '==== Modificado: Laura 23/01/2007
                          'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                          FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                          '==================================================
        
                          CadValues2 = CadValues2 & CadValuesAux2 & I
                          CadValues2 = CadValues2 & ", " & ForpaPosi & ", '" & Format(FecVenci, FormatoFecha) & "', "
        
                          'IMPORTE Resto de Vendimientos
                          ImpVenci = Round(TotalTesor / rsVenci!numerove, 2)
        
                          CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & "," & DBSet(CtaBanco, "T") & ","
                          
                          'David. Para que ponga la cuenta bancaria (SI LA tiene)
                          CadValues2 = CadValues2 & DBSet(vBancoSoc, "T", "S") & "," & DBSet(vSucurSoc, "T", "S") & ","
                          CadValues2 = CadValues2 & DBSet(DigcoSoc, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & ","
        
                          'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
                          Select Case TipoFact
                            Case 1, 3, 7, 9 ' anticipo y anticipo venta campo
                '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
                '                Sql = "Anticipo num.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                                sql = "Anticipo num.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                            Case 2, 4, 8, 10 ' liquidacion y liquidacion venta campo
                '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
                '                Sql = "Liquidacion num.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                                sql = "Liquidacion num.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                            Case Else
                '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
                '                Sql = "Fact.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                                sql = "Fact.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                          End Select
                            
                          CadValues2 = CadValues2 & "'" & DevNombreSQL(sql) & "',"
                        
                          sql = "Variedades: " & Variedades
                          CadValues2 = CadValues2 & "'" & DevNombreSQL(sql) & "'" '),"
                          '[Monica]22/11/2013: Tema iban
                          If vEmpresa.HayNorma19_34Nueva = 1 Then
                              CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & "),"
                          Else
                              CadValues2 = CadValues2 & "),"
                          End If
                      Next I
                      
                      
                      'Ultimo Vencimiento es si lo hay la parte de descuento
                      If TotalTesor1 <> 0 Then ' For i = 2 To rsVenci!numerove
                          I = UltimoVto + 1
                          
'                         'FECHA Resto Vencimientos
'                          FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
'                          '==================================================
'                          'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
'                          FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
        
                          'Comprobar si tiene mes a no girar
'                          FecVenci1 = FecVenci

'                          If DBSet(RS!mesnogir, "N") <> 0 Then
'                                FecVenci1 = ComprobarMesNoGira(FecVenci1, DBSet(RS!mesnogir, "N"), DBSet(0, "N"), RS!DiaPago1, RS!DiaPago2, RS!DiaPago3)
'                          End If
        
                          CadValues2 = CadValues2 & CadValuesAux2 & I & ", " & ForpaPosi & ", '" & Format(FecVenci1, FormatoFecha) & "', "
        
                          'IMPORTE Resto de Vendimientos
                          ImpVenci = TotalTesor1  'Round2(TotalTesor / rsVenci!numerove, 2)
        
                          CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBanco, "T") & ","
                          
                          'David. Para que ponga la cuenta bancaria (SI LA tiene)
                          CadValues2 = CadValues2 & DBSet(vBancoSoc, "T", "S") & "," & DBSet(vSucurSoc, "T", "S") & ","
                          CadValues2 = CadValues2 & DBSet(DigcoSoc, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & ","
        
                          'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
                          Select Case TipoFact
                            Case 1, 3, 7, 9 ' anticipo y anticipo venta campo
                '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
                '                Sql = "Anticipo num.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                                sql = "Anticipo num.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                            Case 2, 4, 8, 10 ' liquidacion y liquidacion venta campo
                '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
                '                Sql = "Liquidacion num.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                                sql = "Liquidacion num.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                            Case Else
                '                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
                '                Sql = "Fact.: " & letraser & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                                sql = "Fact.: " & FactuRec & "-" & Format(fecfactu, "dd/mm/yyyy")
                          End Select
                            
                          CadValues2 = CadValues2 & "'" & DevNombreSQL(sql) & "',"
                          sql = "Variedades: " & Variedades
                          CadValues2 = CadValues2 & "'" & DevNombreSQL(sql) & "'" '),"
                          
                          '[Monica]22/11/2013: Tema iban
                          If vEmpresa.HayNorma19_34Nueva = 1 Then
                              CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & "),"
                          Else
                              CadValues2 = CadValues2 & "),"
                          End If
                          
                      
                      End If
'                      Next i
                      
                    If CadValues2 <> "" Then
                        CadValues2 = Mid(CadValues2, 1, Len(CadValues2) - 1)
                    
                        'Insertamos en la tabla spagop de la CONTA
                        'David. Cuenta bancaria y descripcion textos
                        sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            sql = sql & ", iban) "
                        Else
                            sql = sql & ") "
                        End If
                        
                        sql = sql & " VALUES " & CadValues2
                        ConnConta.Execute sql
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
        LetraSer = ""
        LetraSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(tipoMov, "T"))

'                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
'        Text33csb = "'Factura:" & DBLet(letraser, "T") & "-" & DBLet(numfactu, "T") & " " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
        Text33csb = "'Factura:" & DBLet(FactuRec, "T") & " " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
        Text41csb = "de " & DBSet(TotalTesor * (-1), "N")
        Text42csb = "Variedades: " & Variedades

        sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & DBSet(ForpaNega, "N")
        Set rsVenci = New ADODB.Recordset
        rsVenci.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        If Not rsVenci.EOF Then
            If DBLet(rsVenci!numerove, "N") > 0 Then
                
                CadValuesAux2 = "(" & DBSet(LetraSer, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & ","

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
                
                'IMPORTE del Vencimiento
                ImpVenci = TotalTesor * (-1)

                CC = DBLet(DigcoSoc, "T")
                If DBLet(DigcoSoc, "T") = "**" Then CC = "00"
        
                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ","
                CadValues2 = CadValues2 & DBSet(CtaBanco, "T") & "," & DBSet(vBancoSoc, "T", "S") & "," & DBSet(vSucurSoc, "T", "S") & ","
                CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & "," & DBSet(Text42csb, "T") & ",1" '),"
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & "),"
                Else
                    CadValues2 = CadValues2 & "),"
                End If

                'Resto Vencimientos
                '--------------------------------------------------------------------
                If TotalTesor1 <> 0 Then 'For i = 2 To rsVenci!numerove
                   'FECHA Resto Vencimientos
                    '=== Laura 23/01/2007
                    'FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                    '===
                    I = 2
                    
                    CadValues2 = CadValues2 & CadValuesAux2 & I & ", " & DBSet(Trim(CtaSocio), "T") & ", " & DBSet(ForpaNega, "N") & ", '" & Format(FecVenci, FormatoFecha) & "', "
                    
                    'IMPORTE Resto de Vendimientos
                    'ImpVenci = Round2(TotalTesor * (-1) / rsVenci!numerove, 2)
                    ImpVenci = TotalTesor1 * (-1)
                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ","
                    CadValues2 = CadValues2 & DBSet(CtaBanco, "T") & "," & DBSet(vBancoSoc, "N", "S") & "," & DBSet(vSucurSoc, "N", "S") & ","
                    CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                    CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & "," & DBSet(Text42csb, "T") & ",1" '),"
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & "),"
                    Else
                        CadValues2 = CadValues2 & "),"
                    End If
                    
                End If
                'Next i
                ' quitamos la ultima coma
                CadValues2 = Mid(CadValues2, 1, Len(CadValues2) - 1)

                'Insertamos en la tabla scobro de la CONTA
                sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                sql = sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                sql = sql & " text33csb, text41csb, text42csb, agente" ') "
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    sql = sql & ", iban) "
                Else
                    sql = sql & ") "
                End If
                
                
                sql = sql & " VALUES " & CadValues2
                ConnConta.Execute sql
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

    b = True

EInsertarTesoreriaSoc:
    If Err.Number <> 0 Then
        b = False
        MenError = "Error al insertar en Tesoreria: " & Err.Description
    End If
    InsertarEnTesoreriaSoc = b
End Function

' ### [Monica] 16/01/2008
Public Function InsertarEnTesoreriaNewADV(cadwhere As String, CtaBan As String, FecVen As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
Dim b As Boolean
Dim sql As String, Text33csb As String, Text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim DigConta As String
Dim CC As String
Dim vrefer As String
Dim LetraSer As String
Dim Rsx As ADODB.Recordset
Dim vSocio As cSocio

    On Error GoTo EInsertarTesoreriaNew

    b = False
    InsertarEnTesoreriaNewADV = False
    CadValues = ""
    CadValues2 = ""
    
    Set Rsx = New ADODB.Recordset
    sql = "select * from advfacturas where " & cadwhere
    Rsx.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    If Not Rsx.EOF Then
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(DBLet(Rsx!Codsocio, "N")) Then
            If vSocio.LeerDatosSeccion(DBLet(Rsx!Codsocio, "N"), vParamAplic.SeccionADV) Then
'[Monica]27/09/2011: tanto si el importe es positivo o negativo se introduce en la scobro
'                If DBLet(Rsx!TotalFac, "N") >= 0 Then
                    ' si el importe de la factura es positiva o cero
                    LetraSer = ""
                    LetraSer = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "letraser", "codtipom", Rsx!CodTipom, "T")
        
                    Text33csb = "'Factura:" & DBLet(LetraSer, "T") & "-" & DBLet(Rsx!numfactu, "T") & " " & Format(DBLet(Rsx!fecfactu, "F"), "dd/mm/yy") & "'"
                    Text41csb = "de " & DBSet(Rsx!TotalFac, "N")
        
                    CC = DBLet(vSocio.Digcontrol, "T")
                    If DBLet(vSocio.Digcontrol, "T") = "**" Then CC = "00"
                    
                    '[Monica]03/07/2013: añado trim(codmacta)
                    CadValuesAux2 = "(" & DBSet(LetraSer, "T") & "," & DBSet(Rsx!numfactu, "N") & "," & DBSet(Rsx!fecfactu, "F") & ", 1," & DBSet(Trim(vSocio.CtaClien), "T") & ","
                    CadValues2 = CadValuesAux2 & DBSet(Rsx!codforpa, "N") & "," & DBSet(FecVen, "F") & "," & DBSet(Rsx!TotalFac, "N") & ","
                    CadValues2 = CadValues2 & DBSet(CtaBan, "T") & "," & DBSet(vSocio.Banco, "N", "S") & "," & DBSet(vSocio.Sucursal, "N", "S") & ","
                    CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(vSocio.CuentaBan, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                    CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & ",1" ')"
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        CadValues2 = CadValues2 & "," & DBSet(vSocio.Iban, "T", "S") & ") "
                    Else
                        CadValues2 = CadValues2 & ") "
                    End If
                    
                    
        
                    'Insertamos en la tabla scobro de la CONTA
                    sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                    sql = sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                    sql = sql & " text33csb, text41csb, agente" ') "
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        sql = sql & ", iban) "
                    Else
                        sql = sql & ") "
                    End If
                    
                    sql = sql & " VALUES " & CadValues2
                    ConnConta.Execute sql
                    
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
    
        b = True
    End If
    
EInsertarTesoreriaNew:
    If Err.Number <> 0 Then
        b = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaNewADV = b
End Function



' ### [Monica] 16/01/2008
Public Function InsertarEnTesoreriaNewBOD(cadwhere As String, CtaBan As String, FecVen As String, MenError As String, Tipo As Byte) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
' Tipo: 0 = almazara
'       1 = bodega

Dim b As Boolean
Dim sql As String, Text33csb As String, Text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim DigConta As String
Dim CC As String
Dim vrefer As String
Dim LetraSer As String
Dim Rsx As ADODB.Recordset
Dim vSocio As cSocio
Dim Seccion As Integer
    On Error GoTo EInsertarTesoreriaNew

    b = False
    InsertarEnTesoreriaNewBOD = False
    CadValues = ""
    CadValues2 = ""
    
    Set Rsx = New ADODB.Recordset
    sql = "select * from rbodfacturas where " & cadwhere
    Rsx.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
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
                    LetraSer = ""
                    LetraSer = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "letraser", "codtipom", Rsx!CodTipom, "T")
        
                    Text33csb = "'Factura:" & DBLet(LetraSer, "T") & "-" & DBLet(Rsx!numfactu, "T") & " " & Format(DBLet(Rsx!fecfactu, "F"), "dd/mm/yy") & "'"
                    Text41csb = "de " & DBSet(Rsx!TotalFac, "N")
        
                    CC = DBLet(vSocio.Digcontrol, "T")
                    If DBLet(vSocio.Digcontrol, "T") = "**" Then CC = "00"
        
                    '[Monica]03/07/2013: añado trim(codmacta)
                    CadValuesAux2 = "(" & DBSet(LetraSer, "T") & "," & DBSet(Rsx!numfactu, "N") & "," & DBSet(Rsx!fecfactu, "F") & ", 1," & DBSet(Trim(vSocio.CtaClien), "T") & ","
                    CadValues2 = CadValuesAux2 & DBSet(Rsx!codforpa, "N") & "," & DBSet(FecVen, "F") & "," & DBSet(Rsx!TotalFac, "N") & ","
                    CadValues2 = CadValues2 & DBSet(CtaBan, "T") & "," & DBSet(vSocio.Banco, "N", "S") & "," & DBSet(vSocio.Sucursal, "N", "S") & ","
                    CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(vSocio.CuentaBan, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                    CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & ",1" ')"
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        CadValues2 = CadValues2 & ", " & DBSet(vSocio.Iban, "T", "S") & ") "
                    Else
                        CadValues2 = CadValues2 & ") "
                    End If
                    
        
                    'Insertamos en la tabla scobro de la CONTA
                    sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                    sql = sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                    sql = sql & " text33csb, text41csb, agente" ') "
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        sql = sql & ", iban) "
                    Else
                        sql = sql & ") "
                    End If
                    
                    sql = sql & " VALUES " & CadValues2
                    ConnConta.Execute sql
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
    
        b = True
    End If
    
EInsertarTesoreriaNew:
    If Err.Number <> 0 Then
        b = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaNewBOD = b
End Function





Private Function VariedadesFactura(cadenawhere As String) As String
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String

    On Error Resume Next
    

    sql = "select distinct  nomvarie from rfactsoc_variedad INNER JOIN variedades ON rfactsoc_variedad.codvarie = variedades.codvarie "
    sql = sql & " where (rfactsoc_variedad.codtipom, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu) "
    sql = sql & " in (select codtipom, numfactu, fecfactu from rfactsoc where " & cadenawhere & ")"
     
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    Cad = ""
    While Not Rs.EOF
        Cad = Cad & DBLet(Rs.Fields(0).Value, "T") & ","
    
        Rs.MoveNext
    Wend
    
    If Cad <> "" Then ' quitamos la ultima coma
        Cad = Mid(Cad, 1, Len(Cad) - 1)
    End If
    
    Set Rs = Nothing
    
    VariedadesFactura = Cad
    
End Function


'----------------------------------------------------------------------
' FACTURAS ALMAZARA SOCIOS
'----------------------------------------------------------------------

Public Function PasarFacturaAlmzSoc(cadwhere As String, FechaFin As String, FecRecep As Date, CtaRete As String, TotalFactura As Currency) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura Socio
' ariagro.rcabfactalmz --> conta.cabfactprov
' ariagro.rlinfactalmz --> conta.linfactprov
'Actualizar la tabla ariagro.rcabfactalmz.contabilizada=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim sql As String
Dim Mc As Contadores


    On Error GoTo EContab

    
    Set Mc = New Contadores
    
    '---- Insertar en la conta Cabecera Factura
    b = InsertarCabFactAlmzSoc(cadwhere, cadMen, Mc, CDate(FechaFin), FecRecep, TotalFactura)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
        CtaReten = CtaRete
        
        If b Then
            '---- Insertar lineas de Factura en la Conta
            b = InsertarLinFactAlmzSoc("rcabfactalmz", cadwhere, cadMen, Mc.Contador)
            cadMen = "Insertando Lin. Factura Almazara Socio: " & cadMen
    
            If b Then
                '---- Poner intconta=1 en ariges.scafac
                b = ActualizarCabFactAlmz("rcabfactalmz", cadwhere, cadMen)
                cadMen = "Actualizando Factura Almazara Socio: " & cadMen
            End If
        End If
    End If
    
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura Socio", Err.Description
    End If
    If b Then
        PasarFacturaAlmzSoc = True
    Else
        PasarFacturaAlmzSoc = False
        If Not b Then
            sql = "Insert into tmpErrFac(tipofichero,numfactu,fecfactu,codsocio,error) "
            sql = sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
            sql = sql & " WHERE " & Replace(cadwhere, "rcabfactalmz", "tmpFactu")
            conn.Execute sql
        End If
    End If
End Function


Private Function InsertarCabFactAlmzSoc(cadwhere As String, cadErr As String, ByRef Mc As Contadores, FechaFin As Date, FecRec As Date, TotalFactura As Currency) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Concepto As String

    On Error GoTo EInsertar
       
    sql = " SELECT fecfactu,year(" & DBSet(FecRec, "F") & ") as anofacpr," & DBSet(FecRec, "F") & ",numfactu,rsocios_seccion.codmacpro,"
    sql = sql & "baseimpo, tipoiva, porc_iva,imporiva,basereten,porc_ret,impreten,totalfac, rcabfactalmz.codsocio, rsocios.nomsocio "
    sql = sql & " FROM (" & "rcabfactalmz "
    sql = sql & "INNER JOIN rsocios ON rcabfactalmz.codsocio=rsocios.codsocio) "
    sql = sql & " INNER JOIN rsocios_seccion ON rcabfactalmz.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N")
    sql = sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
    
        If Mc.ConseguirContador("1", (FecRec <= CDate(FechaFin) - 365), True) = 0 Then
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
               sql = "select rsocios.tiporelacion from rsocios where codsocio in (select codsocio from rcabfactalmz where " & cadwhere & ")"
               If DevuelveValor(sql) = 1 Then
                   
                   sql = "select nroasociado from rsocios where codsocio in (select codsocio from rcabfactalmz where " & cadwhere & ")"
                   Socio = DevuelveValor(sql)
                   
                   sql = "select raiz_cliente_asociado from rseccion where codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N")
                   CtaSocio = DevuelveValor(sql) & Format(Socio, "00000")
               End If
            End If
            
            FecRecep = FecRec
            
            Concepto = "ALMAZARA ACEITE"
            
            sql = ""
            sql = Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & AnyoFacPr & "," & DBSet(FecRecep, "F") & "," & DBSet(FacturaSoc, "T") & "," & DBSet(CtaSocio, "T") & "," & DBSet(Concepto, "T") & ","
            sql = sql & DBSet(BaseImp, "N") & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & DBSet(Rs!porc_iva, "N") & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!ImporIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & DBSet(TotalFac, "N") & "," & DBSet(Rs!TipoIVA, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(FecRecep, "F") & ",0"
            Cad = Cad & "(" & sql & ")"
            
            'Insertar en la contabilidad
            sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
            sql = sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
            sql = sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
            sql = sql & " VALUES " & Cad
            ConnConta.Execute sql
            
            'añadido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            sql = sql & ",'" & DevNombreSQL(FacturaSoc) & " @ " & Format(Rs!fecfactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!nomsocio) & "'," & Rs!Codsocio & ")"
            conn.Execute sql
            
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
Dim b As Boolean
Dim sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim CC As String
Dim LetraSer As String
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
    b = False
    
    sql = "select rsocios.nomsocio, rsocios.codbanco, rsocios.codsucur, rsocios.digcontr, rsocios.cuentaba, rsocios_seccion.codmacpro, rsocios.iban "
    sql = sql & " from rsocios inner join rsocios_seccion on rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & vParamAplic.SeccionAlmaz
    sql = sql & " where rsocios.codsocio = " & DBSet(Socio, "N")

    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

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
           sql = "select rsocios.tiporelacion from rsocios where codsocio = " & DBSet(Socio, "N")
           If DevuelveValor(sql) = 1 Then
               
               sql = "select nroasociado from rsocios where codsocio = " & DBSet(Socio, "N")
               Socio2 = DevuelveValor(sql)
               
               sql = "select raiz_cliente_asociado from rseccion where codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N")
               CtaSocio = DevuelveValor(sql) & Format(Socio2, "00000")
           End If
        End If

'lo dejamos como estaba
'[Monica]27/09/2011: tanto si es positivo como si no se almacena en la cartera de cobros
        If TotalTesor > 0 Then ' se insertara en la cartera de pagos (spagop)
            CadValues2 = ""
        
            UltimaFactura = Mid(numfactu, Len(numfactu) - 6, 8)
        
            CadValuesAux2 = "('" & CtaSocio & "', " & DBSet(UltimaFactura, "T") & ", '" & Format(fecfactu, FormatoFecha) & "', "
        
            '------------------------------------------------------------
            I = 1
            CadValues2 = CadValuesAux2 & I
            CadValues2 = CadValues2 & ", " & ForpaPosi & ", '" & Format(FecVenci, FormatoFecha) & "', "
            CadValues2 = CadValues2 & DBSet(TotalTesor, "N") & ", " & DBSet(CtaBanco, "T") & ","
        
            'David. Para que ponga la cuenta bancaria (SI LA tiene)
            CadValues2 = CadValues2 & DBSet(BancoSoc, "T", "S") & "," & DBSet(SucurSoc, "T", "S") & ","
            CadValues2 = CadValues2 & DBSet(DigcoSoc, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & ","
        
            'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
            sql = "Almz.Nros:" & numfactu
                
            CadValues2 = CadValues2 & "'" & DevNombreSQL(sql) & "',"
            
            sql = " de " & Format(fecfactu, "dd/mm/yyyy")
            CadValues2 = CadValues2 & "'" & DevNombreSQL(sql) & "'" ')"
            '[Monica]22/11/2013: Tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & ") "
            Else
                CadValues2 = CadValues2 & ") "
            End If
            
        
            'Grabar tabla spagop de la CONTABILIDAD
            '-------------------------------------------------
            If CadValues2 <> "" Then
                'Insertamos en la tabla spagop de la CONTA
                'David. Cuenta bancaria y descripcion textos
                sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    sql = sql & ", iban) "
                Else
                    sql = sql & ") "
                End If
                
                sql = sql & " VALUES " & CadValues2
                ConnConta.Execute sql
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
            CadValues2 = CadValues2 & DBSet(CtaBanco, "T") & "," & DBSet(BancoSoc, "N", "S") & "," & DBSet(SucurSoc, "N", "S") & ","
            CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & "," & DBSet(Text42csb, "T") & ",1" ')"
            '[Monica]22/11/2013: Tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & ") "
            Else
                CadValues2 = CadValues2 & ") "
            End If
            

            'Insertamos en la tabla scobro de la CONTA
            sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
            sql = sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
            sql = sql & " text33csb, text41csb, text42csb, agente" ') "
            '[Monica]22/11/2013: Tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                sql = sql & ", iban) "
            Else
                sql = sql & ") "
            End If
            sql = sql & " VALUES " & CadValues2
            ConnConta.Execute sql
        End If

        b = True
    End If
    
    
EInsertarTesoreriaAlmz:
    
    If Err.Number <> 0 Then
        b = False
        MenError = "Error al insertar en Tesoreria de Almazara: " & Err.Description
    End If
    InsertarEnTesoreriaAlmz = b
End Function



Private Function InsertarLinFactAlmzSoc(cadTabla As String, cadwhere As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
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
    

    sql = " SELECT sum(rlinfactalmz.importel) as importe "
    sql = sql & " FROM rlinfactalmz "
    sql = sql & " WHERE " & Replace(cadwhere, "rcabfactalmz", "rlinfactalmz")

    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    If Not Rs.EOF Then
        SQLaux = Cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        ImpLinea = Rs!Importe
        '----
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        sql = ""
        Sql2 = ""
        
        sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        sql = sql & DBSet(vParamAplic.CtaGastosAlmz, "T")
        
        Sql2 = sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        sql = sql & "," & DBSet(ImpLinea, "N") & ","
        sql = sql & ValorNulo ' centro de coste
        
        Cad = Cad & "(" & sql & ")" & ","
        I = I + 1
    End If
    
    Rs.Close
    Set Rs = Nothing
    
    ' las retenciones si las hay
    If ImpReten <> 0 Then
        sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        sql = sql & DBSet(CtaSocio, "T")
        sql = sql & "," & DBSet(ImpReten, "N") & ","
        sql = sql & ValorNulo ' no llevan centro de coste
        
        Cad = Cad & "(" & sql & ")" & ","
        I = I + 1
        
        sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        sql = sql & DBSet(CtaReten, "T")
        sql = sql & "," & DBSet(ImpReten * (-1), "N") & ","
        sql = sql & ValorNulo ' no llevan centro de coste
        
        Cad = Cad & "(" & sql & ")" & ","
        I = I + 1
    End If
    
    
    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        sql = sql & " VALUES " & Cad
        ConnConta.Execute sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactAlmzSoc = False
        cadErr = Err.Description
    Else
        InsertarLinFactAlmzSoc = True
    End If
End Function

Private Function ActualizarCabFactAlmz(cadTabla As String, cadwhere As String, cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim sql As String

    On Error GoTo EActualizar
    
    sql = "UPDATE " & cadTabla & " SET contabilizado=1 "
    sql = sql & " WHERE " & cadwhere

    conn.Execute sql
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCabFactAlmz = False
        cadErr = Err.Description
    Else
        ActualizarCabFactAlmz = True
    End If
End Function


Public Function PasarFacturaAlmzCli(cadwhere As String, CodCCost As String, LetraSerie As String, TotalFactura As Currency) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariagro.rcabfactalmz --> conta.cabfact
' ariagro.rlinfactalmz --> conta.linfact
'Actualizar la tabla ariagro.rcabfactalmz.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim sql As String

    On Error GoTo EContab

    
    'Insertar en la conta Cabecera Factura
    b = InsertarCabFactAlmzCli(cadwhere, cadMen, LetraSerie, TotalFactura)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
        CCoste = CodCCost
        'Insertar lineas de Factura en la Conta
        b = InsertarLinFactAlmzCli("rcabfactalmz", cadwhere, cadMen, LetraSerie)
        cadMen = "Insertando Lin. Factura Almazara Cliente: " & cadMen

        If b Then
            'Poner intconta=1 en ariagro.facturas
            b = ActualizarCabFactAlmz("rcabfactalmz", cadwhere, cadMen)
            cadMen = "Actualizando Factura Almazara: " & cadMen
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
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        PasarFacturaAlmzCli = True
    Else
        PasarFacturaAlmzCli = False
        
        sql = "Insert into tmpErrFac(tipofichero,numfactu,fecfactu,codsocio,error) "
        sql = sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
        sql = sql & " WHERE " & Replace(cadwhere, "rcabfactalmz", "tmpFactu")
        conn.Execute sql
    End If
End Function


Private Function InsertarCabFactAlmzCli(cadwhere As String, cadErr As String, LetraSerie As String, TotalFactura As Currency) As Boolean
'Insertando en tabla conta.cabfact
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Concepto As String
Dim Cad As String


    On Error GoTo EInsertar
    
    sql = sql & " SELECT " & DBSet(LetraSerie, "T") & ",tipofichero,numfactu,fecfactu,rsocios_seccion.codmacpro,year(fecfactu) as anofaccl,"
    sql = sql & "baseimpo,tipoiva,porc_iva,imporiva,basereten, porc_ret, impreten, totalfac, tipoiva "
    sql = sql & " FROM (" & "rcabfactalmz inner join " & "rsocios_seccion on rcabfactalmz.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & vParamAplic.SeccionAlmaz & ") "
    sql = sql & "INNER JOIN " & "rsocios ON rsocios_seccion.codsocio=rsocios.codsocio "
    sql = sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
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
           sql = "select rsocios.tiporelacion from rsocios where codsocio in (select codsocio from rcabfactalmz where " & cadwhere & ")"
           If DevuelveValor(sql) = 1 Then
               
               sql = "select nroasociado from rsocios where codsocio in (select codsocio from rcabfactalmz where " & cadwhere & ")"
               Socio = DevuelveValor(sql)
               
               sql = "select raiz_cliente_asociado from rseccion where codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N")
               CtaSocio = DevuelveValor(sql) & Format(Socio, "00000")
           End If
        End If
        
        sql = ""
        sql = "'" & LetraSerie & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(CtaSocio, "T") & "," & Year(Rs!fecfactu) & "," & DBSet(Concepto, "T") & ","
        sql = sql & DBSet(Rs!baseimpo, "N") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!porc_iva, "N") & "," & ValorNulo & "," & ValorNulo & ","
        sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!ImporIva, "N", "N") & ","
        sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        sql = sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!TipoIVA, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
        sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        sql = sql & DBSet(Rs!fecfactu, "F")
        Cad = Cad & "(" & sql & ")"
'        RS.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    'Insertar en la contabilidad
    sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
    sql = sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
    sql = sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
    sql = sql & " VALUES " & Cad
    ConnConta.Execute sql
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactAlmzCli = False
        cadErr = Err.Description
    Else
        InsertarCabFactAlmzCli = True
    End If
End Function


Private Function InsertarLinFactAlmzCli(cadTabla As String, cadwhere As String, cadErr As String, LetraSerie As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
Dim I As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String

    On Error GoTo EInLinea
    

    sql = " SELECT " & DBSet(LetraSerie, "T") & ",rlinfactalmz.numfactu,rlinfactalmz.fecfactu,rlinfactalmz.codsocio," & vParamAplic.CtaVentasAlmz & ",sum(importel) as importe "
    sql = sql & " FROM rlinfactalmz "
    sql = sql & " WHERE " & Replace(cadwhere, "rcabfactalmz", "rlinfactalmz")
    sql = sql & " GROUP BY 1,2,3,4,5 "
    sql = sql & " order by 1,2,3,4,5 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = Cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        totimp = totimp + DBLet(Rs!Importe, "N")
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        sql = ""
        Sql2 = ""
        
        sql = "'" & LetraSerie & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
        sql = sql & DBSet(vParamAplic.CtaVentasAlmz, "T")
        
        Sql2 = sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        sql = sql & "," & DBSet(Rs!Importe, "N") & ","
        
        sql = sql & ValorNulo ' centro de coste
        
        
        Cad = Cad & "(" & sql & ")" & ","
        
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
            Cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            Cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If


    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        sql = sql & " VALUES " & Cad
        ConnConta.Execute sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactAlmzCli = False
        cadErr = Err.Description
    Else
        InsertarLinFactAlmzCli = True
    End If
End Function


'??????????????
'?????????????? POZOS
'??????????????

Public Function InsertarEnTesoreriaPOZOS(MenError As String, ByRef RS1 As ADODB.Recordset, FecVenci As Date, Forpa As String, CtaBanco As String) As Boolean
'Guarda datos de Tesoreria en tablas: spagop o scobro dependiendo del signo de la factura
Dim b As Boolean
Dim sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim CC As String
Dim LetraSer As String

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
Dim Cad As String
            
    On Error GoTo EInsertarTesoreriaPOZ

    InsertarEnTesoreriaPOZOS = False
    b = False
    
    Text71csb = ""
    Text72csb = ""
    Text82csb = ""
    
    sql = "select rsocios.nomsocio, rsocios.codbanco, rsocios.codsucur, rsocios.digcontr, rsocios.cuentaba, rsocios_seccion.codmaccli, rsocios.nifsocio, "
    '[Monica]03/08/2012: añadimos los datos fiscales a la scobro
    sql = sql & " rsocios.dirsocio, rsocios.pobsocio, rsocios.prosocio, rsocios.codpostal, rsocios.iban "
    sql = sql & " from rsocios inner join rsocios_seccion on rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & vParamAplic.SeccionPOZOS
    sql = sql & " where rsocios.codsocio = " & DBSet(RS1!Codsocio, "N")

    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

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
        Sql5 = "select sum(totalfact+coalesce(imprecargo,0)) from rrecibpozos where codtipom = " & DBSet(RS1!CodTipom, "T")
        Sql5 = Sql5 & " and numfactu = " & DBSet(RS1!numfactu, "N")
        Sql5 = Sql5 & " and fecfactu = " & DBSet(RS1!fecfactu, "F")
        
        TotalFact = DevuelveValor(Sql5)
        
        Select Case DBLet(RS1!CodTipom, "T")
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
                        Cad = Mid(Year(DBLet(RS1!fecfactu, "F")), 3, 2) & "RCP" & Format(DBLet(RS1!numfactu, "N"), "0000000") & " Cont:" & Format(CLng(DBLet(Rs6!Hidrante, "T")), "00000")
                        Cad = Cad & " Soc:" & Format(DBLet(RS1!Codsocio, "N"), "000000") & " " & Mid(DBLet(Rs6!Concepto), 1, 15) & " Pol/Par:" & Mid(Trim(DBLet(Rs6!Poligono, "T")), 1, 2) & "/" & DBLet(Rs6!parcelas)

                        If Len(Cad) > 80 Then Cad = Mid(Cad, 1, 78) & ".."
                        Text33csb = Cad

                        Cad = "Lec:" & Format(DBLet(Rs6!fech_act, "F"), "dd-mm-yy") & " " & Format(DBLet(Rs6!Consumo1, "N"), "000000") & " m³ Pr:" & Format(DBLet(Rs6!Precio1, "N"), "0.00") & " €/m³ Total: " & Format(DBLet(Rs6!TotalFact, "N"), "###,##0.00")
                        Text41csb = Cad

                        '[Monica]20/02/2014: en utxera tb grabamos el codigo de socio
                        'Referencia = DBLet(Rs6!Hidrante, "T")
                        Referencia = Format(DBLet(RS1!Codsocio, "N"), "000000")
                    Else ' Escalona
                       '[Monica]29/01/2014: cambios text33csb(80) text41csb(60)
                       
                       
                       '[Monica]20/06/2014: cambiamos lo que imprimimos en los textos (quitamos socio y añadimos fecha de lectura anterior
                       '                    los mismos cambios para utxera
                       
                        Text33csb = ""
                        Text41csb = ""
                        Cad = Mid(Year(DBLet(RS1!fecfactu, "F")), 3, 2) & "RCP" & Format(DBLet(RS1!numfactu, "N"), "0000000") & " Cont:" & Format(CLng(DBLet(Rs6!Hidrante, "T")), "00000")
                        Cad = Cad & " Pol/Par:" & Mid(Trim(DBLet(Rs6!Poligono, "T")), 1, 2) & "/" & Mid(Trim(DBLet(Rs6!parcelas)), 1, 20) & " Lec.ant:" & Format(DBLet(Rs6!lect_ant, "N"), "000000000")
                        
'                        If Len(Cad) > 80 Then Cad = Mid(Cad, 1, 78) & ".."

                        Text33csb = Cad
                        
                        Dim longitud As Integer
                        longitud = Len(Cad)
                        
                        Text33csb = Text33csb & Space(80 - longitud)
                        
                        Cad = "Le.ac:" & Format(DBLet(Rs6!lect_act, "N"), "000000000") & " Con:" & Format(DBLet(Rs6!Consumo1, "N"), "000000") & " Pr:" & Format(DBLet(Rs6!Precio1, "N"), "#0.00") & "€/m³ Tot:" & Format(DBLet(Rs6!TotalFact, "N"), "#####0.00")
                        '[Monica]15/01/2016: si hay recargo lo especifico
                        If DBLet(Rs6!imprecargo, "N") <> 0 Then
                            Cad = Cad & "+" & Format(DBLet(Rs6!imprecargo, "N"), "##0.00")
                        End If
                        Text41csb = Cad
                        
                        longitud = Len(Cad)
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
                        Cad = Mid(Year(DBLet(RS1!fecfactu, "F")), 3, 2) & "RMP" & Format(DBLet(RS1!numfactu, "N"), "0000000")
                        
                        '[Monica]29/04/2014: grabamos las hanegadas y las brazas en lugar de "Precios según información enviada"
                        Cad = Cad & " Soc:" & Format(DBLet(RS1!Codsocio, "N"), "000000") & " " & Mid(DBLet(Rs6!Concepto), 1, 20) & " " & Format(v_hanegada, "#####0") & "hg " & Format(v_brazas, "#####0") & "br a " & DBSet(Rs6!Precio, "N") & "Eur" ' " Precios según información enviada"
                         
                        Text33csb = Cad
                         
                        Cad = "Tot:" & Format(DBLet(Rs6!TotalFact, "N"), "####0.00") & " "
                        Cad = Cad & DBLet(Rs6!Hidrante, "T")
                        
                        Text41csb = Cad
                        
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
                            Cad = Mid(Year(DBLet(RS1!fecfactu, "F")), 3, 2) & "RMP" & Format(DBLet(RS1!numfactu, "N"), "0000000")
                            Cad = Cad & " Soc:" & Format(DBLet(RS1!Codsocio, "N"), "000000") & " " & Mid(DBLet(Rs6!Concepto), 1, 20) & " Precios según información enviada"
                             
                            Text33csb = Cad
                             
                            Cad = "Tot:" & Format(DBLet(Rs6!TotalFact, "N"), "####0.00")
                            '[Monica]15/01/2016: metemos el recargo
                            If DBLet(Rs6!imprecargo, "N") <> 0 Then
                                Cad = Cad & " Rec:" & Format(DBLet(Rs6!imprecargo, "N"), "###0.00")
                            End If
                            Cad = Cad & " "
                            
                            
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

                                If I > 1 Then Cad = Cad & "/"

                                Cad = Cad & Format(CLng(DBLet(Rs6!Hidrante, "T")), "00000")
                                
                                Rs4.MoveNext
                            Wend
                            Text41csb = Cad
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
                            Text63csb = "N.I.F.: " & DBLet(Rs!nifsocio, "N")
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
                     Cad = Mid(Year(DBLet(RS1!fecfactu, "F")), 3, 2) & "TAL" & Format(DBLet(RS1!numfactu, "N"), "0000000")
                     Cad = Cad & " Soc:" & Format(DBLet(RS1!Codsocio, "N"), "000000") & " " & Mid(DBLet(Rs6!Concepto), 1, 15) & " Precios según información enviada"
                     
                     Text33csb = Cad
                     
                     Cad = "Tot:" & Format(DBLet(Rs6!TotalFact, "N"), "####0.00") & " "
                    '[Monica]15/01/2016: metemos el recargo
                    If DBLet(Rs6!imprecargo, "N") <> 0 Then
                        Cad = Cad & " Rec:" & Format(DBLet(Rs6!imprecargo, "N"), "###0.00")
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
                        Cad = Cad & Format(DBLet(Rs4!Poligono, "N"), "00") & "/" & Format(DBLet(Rs4!Parcela, "N"), "000")
                        If DBLet(Rs4!SubParce, "T") = "" Then
                            Cad = Cad & "  "
                        Else
                            Cad = Cad & Mid(DBLet(Rs4!SubParce, "T"), 1, 1) & " "
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
                    Text41csb = Cad
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
                        Cad = Mid(Year(DBLet(RS1!fecfactu, "F")), 3, 2) & "RVP" & Format(DBLet(RS1!numfactu, "N"), "0000000")
                        Cad = Cad & " Soc:" & Format(DBLet(RS1!Codsocio, "N"), "000000") & " " & Mid(DBLet(Rs6!importemo), 1, 30) & " Tot:" & Format(DBLet(Rs6!TotalFact, "N"), "####0.00")
                        
                        '[Monica]15/01/2016: metemos el recargo
                        If DBLet(Rs6!imprecargo, "N") <> 0 Then
                            Cad = Cad & " Rec:" & Format(DBLet(Rs6!imprecargo, "N"), "###0.00")
                        End If
                        
                         
                        Text33csb = Cad
                         
                        Cad = DBLet(Rs6!Conceptoar1, "T") & "/" & DBLet(Rs6!Conceptoar2, "T")
                        
                        Text41csb = Cad
                        
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
                        Cad = Mid(Year(DBLet(RS1!fecfactu, "F")), 3, 2) & "RMT" & Format(DBLet(RS1!numfactu, "N"), "0000000")
                        Cad = Cad & " " & DBLet(Rs6!Concepto)
                         
                        Text33csb = Cad
                         
                        Cad = "Tot:" & Format(DBLet(Rs6!TotalFact, "N"), "####0.00") & " "
                        '[Monica]15/01/2016: metemos el recargo
                        If DBLet(Rs6!imprecargo, "N") <> 0 Then
                            Cad = Cad & " Rec:" & Format(DBLet(Rs6!imprecargo, "N"), "###0.00")
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
                        
                            Cad = Cad & " " & Mid(DBLet(Rs4!nomparti, "T"), 1, 15) & " " & DBLet(Rs4!Poligono, "N") & " " & DBLet(Rs4!Parcela, "N") & " " & Format(v_hanegada, "###0") & "H " & Format(v_brazas, "###0") & "B " & Format(DBLet(Rs4!Precio1, "N"), "#,##0.0000")
                        End If
                        Text41csb = Cad
                        
                        Set Rs4 = Nothing
                        
                        Referencia = Format(DBLet(RS1!Codsocio, "N"), "000000")
                    
                
                End If
                
                Set Rs6 = Nothing
            
            '[Monica]15/01/2016: todas las facturas rectificativas de escalona
            Case "RRC", "RRM", "RRT", "RRV" ', "RTA"
                 Text33csb = ""
                 Text41csb = ""
                 
                 Cad = Mid(Year(DBLet(RS1!fecfactu, "F")), 3, 2) & DBLet(RS1!CodTipom, "T") & Format(DBLet(RS1!numfactu, "N"), "0000000")
                 Cad = Cad & " Rectifica la factura: " & DBLet(RS1!CodTipomrec, "T") & "-" & Format(DBLet(RS1!numfacturec, "N"), "0000000") & " de fecha " & Format(DBLet(RS1!fecfacturec, "F"), "dd/mm/yyyy")
                 
                 Text33csb = Cad
                 
                 Cad = "Tot:" & Format(DBLet(RS1!TotalFact, "N"), "####0.00") & " "
                '[Monica]15/01/2016: metemos el recargo
                If DBLet(RS1!imprecargo, "N") <> 0 Then
                    Cad = Cad & " Rec:" & Format(DBLet(RS1!imprecargo, "N"), "###0.00")
                End If
                 
                
                Text41csb = Cad
                Referencia = Format(DBLet(RS1!Codsocio, "T"), "000000")
                 
            
            
        
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
        CadValues2 = CadValuesAux2 & DBSet(Forpa, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet((TotalFact), "N") & ","
        CadValues2 = CadValues2 & DBSet(CtaBanco, "T") & "," & DBSet(BancoSoc, "N", "S") & "," & DBSet(SucurSoc, "N", "S") & ","
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
        sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
        sql = sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
        '[Monica] 16/07/2010: hemos añadido todo lo que debe llevar impreso el recibo de banco ( desde agente )
        sql = sql & " text33csb, text41csb, text42csb, agente, text43csb, text51csb, text52csb, text53csb,"
        sql = sql & " text61csb, text62csb, text63csb, text71csb,text72csb,text73csb, text81csb, text82csb, text83csb, referencia, "
        sql = sql & " nomclien, domclien, pobclien, cpclien, proclien" ') "
        '[Monica]22/11/2013: Tema iban
        If vEmpresa.HayNorma19_34Nueva = 1 Then
            sql = sql & ", iban) "
        Else
            sql = sql & ") "
        End If
        sql = sql & " VALUES " & CadValues2
        ConnConta.Execute sql


        '[Monica]15/01/2016: para el caso de Escalona cuando la factura es rectificativa actualizamos su cobro y el de la factura que rectifica
        '                    con el importe de vencimiento + gastos
        If vParamAplic.Cooperativa = 10 Then
            If DBLet(RS1!CodTipom, "T") = "RRC" Or DBLet(RS1!CodTipom, "T") = "RRM" Or DBLet(RS1!CodTipom, "T") = "RRT" Or _
               DBLet(RS1!CodTipom, "T") = "RRV" Or DBLet(RS1!CodTipom, "T") = "RTA" Then
                         
                 Dim LSer As String
                 LSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(RS1!CodTipom, "T"))
                
                 sql = "update scobro set impcobro = impvenci + gastos, fecultco = " & DBSet(FecVenci, "F")
                 sql = sql & " where numserie = " & DBSet(LSer, "T") & " and codfaccl = " & DBSet(RS1!numfactu, "N")
                 sql = sql & " and fecfaccl = " & DBSet(RS1!fecfactu, "F")
                 
                 ConnConta.Execute sql
                 
                 LSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(RS1!CodTipomrec, "T"))
                 
                 sql = "update scobro set impcobro = impvenci + gastos, fecultco = " & DBSet(FecVenci, "F")
                 sql = sql & " where numserie = " & DBSet(LSer, "T") & " and codfaccl = " & DBSet(RS1!numfacturec, "N")
                 sql = sql & " and fecfaccl = " & DBSet(RS1!fecfacturec, "F")
                 
                 ConnConta.Execute sql
            End If
        End If

        b = True
        
    Else
        MenError = "No se ha encontrado socio " & DBLet(RS1!Codsocio, "N") & " o no tiene seccion asignada. Revise. "
    End If
    
    
EInsertarTesoreriaPOZ:
    
    If Err.Number <> 0 Then
        b = False
        MenError = "Error al insertar en Tesoreria de POZOS: " & Err.Description
    End If
    InsertarEnTesoreriaPOZOS = b
End Function


Private Function ParteCadena(Origen As String, NroSubcadenas As Integer, longitud) As String
Dim J As Integer
Dim I As Integer
Dim k As Integer
Dim Cad As String

    On Error Resume Next

    ParteCadena = ""

    J = 1
    Cad = ""
    For k = 1 To NroSubcadenas
        I = J + longitud - 1
        If Len(Origen) - J > 0 Then
            If Len(Mid(Origen, J + 1, Len(Origen) - J)) > longitud Then
                While Mid(Origen, I + 1, 1) <> " "
                    I = I - 1
                Wend
            End If
            If J > 0 And I - J + 1 > 0 Then
                Cad = Cad & Mid(Origen, J, I - J + 1) & "|"
            End If
            J = I + 1
        End If
    Next k
    
    ParteCadena = Cad
    

End Function


'----------------------------------------------------------------------
' FACTURAS TRANSPORTISTAS
'----------------------------------------------------------------------

Public Function PasarFacturaTra(cadwhere As String, CodCCost As String, FechaFin As String, Seccion As String, TipoFact As Byte, FecRecep As Date, FecVto As Date, ForpaPos As String, ForpaNeg As String, CtaBanc As String, CtaRete As String, CtaApor As String, TipoM As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariagro.rfactsoc --> conta.cabfactprov
' ariagro.rfactsoc_variedad --> conta.linfactprov
'Actualizar la tabla ariagro.rfactsoc.contabilizada=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim sql As String
Dim Mc As Contadores


    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
        
    
    Set Mc = New Contadores
    
    '[Monica]17/10/2011: FACTURAS INTERNAS
    If EsFacturaInternaTrans(cadwhere) Then
        CtaReten = CtaRete
        ' Insertamos en el diario
        b = InsertarAsientoDiarioTRANS(cadwhere, cadMen, Mc, CDate(FechaFin), Seccion, TipoFact, FecRecep, TipoM)
        cadMen = "Insertando Factura en Diario: " & cadMen
    Else
    
        '---- Insertar en la conta Cabecera Factura
        b = InsertarCabFactTra(cadwhere, cadMen, Mc, CDate(FechaFin), Seccion, TipoFact, FecRecep, TipoM)
        cadMen = "Insertando Cab. Factura: " & cadMen
        
    End If
    
    If b Then
        FecVenci = FecVto
        ForpaPosi = ForpaPos
        ForpaNega = ForpaNeg
        CtaBanco = CtaBanc
        CtaReten = CtaRete
        CtaAport = CtaApor
        tipoMov = TipoM    ' codtipom de la factura de socio
        
'01-06-2009
        b = InsertarEnTesoreriaTra(cadwhere, cadMen, FacturaTRA, FecFactuTRA)
        cadMen = "Insertando en Tesoreria: " & cadMen

        If b Then
            CCoste = CodCCost
            '[Monica]17/10/2011: INTERNAS
            If Not EsFacturaInternaTrans(cadwhere) Then
                '---- Insertar lineas de Factura en la Conta
                b = InsertarLinFactTra("rfacttra", cadwhere, cadMen, TipoFact, Mc.Contador)
                cadMen = "Insertando Lin. Factura: " & cadMen
            End If
    
            If b Then
                '---- Poner intconta=1 en ariagro.rfacttra
                b = ActualizarCabFactSoc("rfacttra", cadwhere, cadMen)
                cadMen = "Actualizando Factura Transporte: " & cadMen
            End If
        End If
    End If
    
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura Transporte", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaTra = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaTra = False
        If Not b Then
            InsertarTMPErrFacSoc cadMen, cadwhere
        End If
    End If
End Function


Private Function InsertarCabFactTra(cadwhere As String, cadErr As String, ByRef Mc As Contadores, FechaFin As Date, Secci As String, Tipo As Byte, FecRec As Date, TipoM As String) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Concepto As String
Dim LetraSer As String

    On Error GoTo EInsertar
       
    
    sql = " SELECT fecfactu,year(" & DBSet(FecRec, "F") & ") as anofacpr," & DBSet(FecRec, "F") & ",numfactu,rtransporte.codmacpro,"
    sql = sql & "baseimpo, tipoiva, porc_iva,imporiva,basereten,porc_ret,impreten,baseaport,porc_apo,impapor,totalfac,"
    sql = sql & "rectif_codtipom, rectif_numfactu, rectif_fecfactu,"
    sql = sql & "rtransporte.codtrans, rtransporte.nomtrans, rtransporte.codbanco, rtransporte.codsucur, rtransporte.digcontr, rtransporte.cuentaba "
    sql = sql & ",rtransporte.iban "
    sql = sql & " FROM (" & "rfacttra "
    sql = sql & "INNER JOIN rtransporte ON rfacttra.codtrans=rtransporte.codtrans) "
    sql = sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
    
        If Mc.ConseguirContador("1", (FecRec <= CDate(FechaFin) - 365), True) = 0 Then
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            BaseImp = DBLet(Rs!baseimpo, "N")
            TotalFac = BaseImp + DBLet(Rs!ImporIva, "N")
            AnyoFacPr = Rs!anofacpr
            
            ImpReten = DBLet(Rs!ImpReten, "N")
            ImpAport = DBLet(Rs!impapor, "N")
            
            LetraSer = ""
            LetraSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(TipoM, "T"))
            
            FacturaTRA = LetraSer & "-" & DBLet(Rs!numfactu, "N")
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
            
            sql = ""
            sql = Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & AnyoFacPr & "," & DBSet(FecRecep, "F") & "," & DBSet(FacturaTRA, "T") & "," & DBSet(CtaTransporte, "T") & "," & DBSet(Concepto, "T") & ","
            sql = sql & DBSet(BaseImp, "N") & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & DBSet(Rs!porc_iva, "N") & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!ImporIva, "N") & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & DBSet(TotalFac, "N") & "," & DBSet(Rs!TipoIVA, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(FecRecep, "F") & ",0"
            Cad = Cad & "(" & sql & ")"
            
            'Insertar en la contabilidad
            sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
            sql = sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
            sql = sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
            sql = sql & " VALUES " & Cad
            ConnConta.Execute sql
            
            'añadido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            sql = sql & ",'" & DevNombreSQL(FacturaTRA) & " @ " & Format(Rs!fecfactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!codTrans) & "')"
            conn.Execute sql
            
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


Public Function InsertarEnTesoreriaTra(cadwhere As String, MenError As String, numfactu As String, fecfactu As Date) As Boolean
'Guarda datos de Tesoreria en tablas: spagop o scobro dependiendo del signo de la factura
Dim b As Boolean
Dim sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim CC As String
Dim LetraSer As String
Dim Text33csb As String
Dim Text41csb As String
Dim Text42csb As String
Dim GastosPie As Currency

    On Error GoTo EInsertarTesoreriaTra

    InsertarEnTesoreriaTra = False
    
    
    If TotalTesor > 0 Then ' se insertara en la cartera de pagos (spagop)
        CadValues2 = ""
    
        'vamos creando la cadena para insertar en spagosp de la CONTA
        LetraSer = ""
        LetraSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(tipoMov, "T"))
        
        '[Monica]03/07/2013: añado trim(codmacta)
        CadValuesAux2 = "('" & Trim(CtaTransporte) & "', " & DBSet(LetraSer & "-" & numfactu, "T") & ", '" & Format(fecfactu, FormatoFecha) & "', "
    
        '------------------------------------------------------------
        I = 1
        CadValues2 = CadValuesAux2 & I
        CadValues2 = CadValues2 & ", " & ForpaPosi & ", '" & Format(FecVenci, FormatoFecha) & "', "
        CadValues2 = CadValues2 & DBSet(TotalTesor, "N") & ", " & DBSet(CtaBanco, "T") & ","
    
        'David. Para que ponga la cuenta bancaria (SI LA tiene)
        CadValues2 = CadValues2 & DBSet(BancoTRA, "T", "S") & "," & DBSet(SucurTRA, "T", "S") & ","
        CadValues2 = CadValues2 & DBSet(DigcoTRA, "T", "S") & "," & DBSet(CtaBaTRA, "T", "S") & ","
    
        'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
        sql = "Factura num.: " & LetraSer & "-" & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
            
        CadValues2 = CadValues2 & "'" & DevNombreSQL(sql) & "',"
        
        'SQL = "Variedades: " & Variedades
        sql = ""
        CadValues2 = CadValues2 & "'" & DevNombreSQL(sql) & "'" ')"
        
        '[Monica]22/11/2013: Tema iban
        If vEmpresa.HayNorma19_34Nueva = 1 Then
            CadValues2 = CadValues2 & ", " & DBSet(IbanTRA, "T", "S") & ") "
        Else
            CadValues2 = CadValues2 & ") "
        End If
        
    
        'Grabar tabla spagop de la CONTABILIDAD
        '-------------------------------------------------
        If CadValues2 <> "" Then
            'Insertamos en la tabla spagop de la CONTA
            'David. Cuenta bancaria y descripcion textos
            sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
            
            '[Monica]22/11/2013: Tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                sql = sql & ", iban) "
            Else
                sql = sql & ") "
            End If
            sql = sql & " VALUES " & CadValues2
            ConnConta.Execute sql
        End If
    Else
        ' si es negativo se inserta en positivo en la cartera de cobros (scobro)
        
        LetraSer = ""
        LetraSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(tipoMov, "T"))
        
        Text33csb = "'Factura:" & DBLet(LetraSer, "T") & "-" & DBLet(numfactu, "T") & " " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
        Text41csb = "de " & DBSet(TotalTesor, "N")
'        text42csb = "Variedades: " & Variedades
        
        CC = DBLet(DigcoTRA, "T")
        If DBLet(DigcoTRA, "T") = "**" Then CC = "00"
    
        '[Monica]03/07/2013: añado trim(codmacta)
        CadValuesAux2 = "(" & DBSet(LetraSer, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & ", 1," & DBSet(Trim(CtaTransporte), "T") & ","
        CadValues2 = CadValuesAux2 & DBSet(ForpaNega, "N") & "," & DBSet(fecfactu, "F") & "," & DBSet(TotalTesor * (-1), "N") & ","
        CadValues2 = CadValues2 & DBSet(CtaBanco, "T") & "," & DBSet(BancoTRA, "N", "S") & "," & DBSet(SucurTRA, "N", "S") & ","
        CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(CtaBaTRA, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & "," & DBSet(Text42csb, "T") & ",1" ')"
        '[Monica]22/11/2013: Tema iban
        If vEmpresa.HayNorma19_34Nueva = 1 Then
            CadValues2 = CadValues2 & ", " & DBSet(IbanTRA, "T", "S") & ") "
        Else
            CadValues2 = CadValues2 & ") "
        End If
        
        
        'Insertamos en la tabla scobro de la CONTA
        sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
        sql = sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
        sql = sql & " text33csb, text41csb, text42csb, agente" ') "
        '[Monica]22/11/2013: Tema iban
        If vEmpresa.HayNorma19_34Nueva = 1 Then
            sql = sql & ", iban) "
        Else
            sql = sql & ") "
        End If
        
        sql = sql & " VALUES " & CadValues2
        ConnConta.Execute sql
    
    End If

    b = True

EInsertarTesoreriaTra:
    If Err.Number <> 0 Then
        b = False
        MenError = "Error al insertar en Tesoreria: " & Err.Description
    End If
    InsertarEnTesoreriaTra = b
End Function


Private Function InsertarLinFactTra(cadTabla As String, cadwhere As String, cadErr As String, Tipo As Byte, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
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
        sql = " SELECT 1, variedades.ctatransporte as cuenta, sum(rfacttra_albaran.importe) as importe, variedades.codccost "
    Else
        sql = " SELECT 1, variedades.ctatransporte as cuenta, sum(rfacttra_albaran.importe) as importe "
    End If
    sql = sql & " FROM rfacttra_albaran, variedades "
    sql = sql & " WHERE " & Replace(cadwhere, "rfacttra", "rfacttra_albaran") & " and"
    sql = sql & " rfacttra_albaran.codvarie = variedades.codvarie "
    sql = sql & " group by 1,2 "
    sql = sql & " order by 1,2 "

    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = Cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        ImpLinea = Rs!Importe
        
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        sql = ""
        Sql2 = ""
        
        sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        sql = sql & DBSet(Rs!cuenta, "T")
        
        Sql2 = sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        sql = sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
                If DBLet(Rs!CodCCost, "T") = "----" Then
                    sql = sql & DBSet(CCoste, "T")
                Else
                    sql = sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBLet(Rs!CodCCost, "T")
                End If
        Else
            sql = sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        Cad = Cad & "(" & sql & ")" & ","
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
            Cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            Cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If

    ' las retenciones si las hay
    If ImpReten <> 0 Then
        sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        sql = sql & DBSet(CtaTransporte, "T")
        sql = sql & "," & DBSet(ImpReten, "N") & ","
        sql = sql & ValorNulo ' no llevan centro de coste
        
        Cad = Cad & "(" & sql & ")" & ","
        I = I + 1
        
        sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        sql = sql & DBSet(CtaReten, "T")
        sql = sql & "," & DBSet(ImpReten * (-1), "N") & ","
        sql = sql & ValorNulo ' no llevan centro de coste
        
        Cad = Cad & "(" & sql & ")" & ","
        I = I + 1
    End If
    
    ' las aportaciones de fondo operativo si las hay
    If ImpAport <> 0 Then
        sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        sql = sql & DBSet(CtaTransporte, "T")
        sql = sql & "," & DBSet(ImpAport, "N") & ","
        sql = sql & ValorNulo ' no llevan centro de coste
        
        Cad = Cad & "(" & sql & ")" & ","
        I = I + 1
    
        sql = NumRegis & "," & AnyoFacPr & "," & I & ","
        sql = sql & DBSet(CtaAport, "T")
        sql = sql & "," & DBSet(ImpAport * (-1), "N") & ","
        sql = sql & ValorNulo ' no llevan centro de coste
        
        Cad = Cad & "(" & sql & ")" & ","
        I = I + 1
    End If
    
    
    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        sql = sql & " VALUES " & Cad
        ConnConta.Execute sql
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


Public Function EsFacturaInterna(cWhere As String) As Boolean
Dim sql As String


    On Error GoTo eEsFacturaInterna
    
    sql = "select rsocios.esfactadvinterna from rfactsoc inner join rsocios on rfactsoc.codsocio = rsocios.codsocio "
    sql = sql & " where " & cWhere
    
    EsFacturaInterna = (DevuelveValor(sql) = 1)
    Exit Function
    
eEsFacturaInterna:
    MuestraError Err.Number, "Es Factura Interna", Err.Description
End Function

Public Function EsFacturaInternaTrans(cWhere As String) As Boolean
Dim sql As String


    On Error GoTo eEsFacturaInternaTrans
    
    sql = "select rtransporte.esfacttrainterna from rfacttra inner join rtransporte on rfacttra.codtrans = rtransporte.codtrans "
    sql = sql & " where " & cWhere
    
    EsFacturaInternaTrans = (DevuelveValor(sql) = 1)
    Exit Function
    
eEsFacturaInternaTrans:
    MuestraError Err.Number, "Es Factura Interna de Transporte", Err.Description
End Function


Private Function EsContado(vCadena As String) As Boolean
Dim sql As String

    sql = "select escontado from rrecibpozos where " & vCadena
    EsContado = (DevuelveValor(sql) = "1")

End Function

Public Function PasarFacturaPOZOS(cadwhere As String, CodCCost As String, CtaBan As String, FecVen As String, TipoM As String, FecFac As Date, Observac As String, Forpa As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
Dim b As Boolean
Dim cadMen As String
Dim sql As String
Dim Mc As Contadores
Dim Obs As String
Dim RS1 As ADODB.Recordset


    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    
    'Insertar en la conta Cabecera Factura
    b = InsertarCabFactPOZ(cadwhere, Observac, cadMen)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
        CCoste = CodCCost
        'Insertar lineas de Factura en la Conta
        b = InsertarLinFactPOZ("rrecibpozos", cadwhere, cadMen, TipoM)
        cadMen = "Insertando Lin. Factura: " & cadMen
        
        '++monica:añadida la parte de insertar en tesoreria
        If b Then
            '[Monica]30/09/2011: como tenia hecha la funcion de insertar en tesoreria para todos,
            '                    la aprovecho y le paso como parametro el recordset Rs1
            sql = "select * from rrecibpozos where " & cadwhere
            Set RS1 = New ADODB.Recordset
            RS1.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            '[Monica]18/07/2014: añadida la funcion de si es contado
            If TipoM = "RMT" And EsContado(cadwhere) Then
                b = InsertarAsientoCobroPOZOS(cadMen, RS1, CDate(FecVen), CtaBan)
            Else
                b = InsertarEnTesoreriaPOZOS(cadMen, RS1, CDate(FecVen), Forpa, CtaBan)
            End If
            cadMen = "Insertando en Tesoreria: " & cadMen
            
            Set RS1 = Nothing
        End If
    End If
        '++

    If b Then
        'Poner intconta=1 en ariagro.facturas
        b = ActualizarCabFact("rrecibpozos", cadwhere, cadMen)
        cadMen = "Actualizando Factura: " & cadMen
    End If
    
    
EContab:
    
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Recibos Pozos", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaPOZOS = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaPOZOS = False
        
        sql = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
        sql = sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
        sql = sql & " WHERE " & Replace(cadwhere, "rrecibpozos", "tmpFactu")
        conn.Execute sql
    End If
End Function

Private Function InsertarCabFactPOZ(cadwhere As String, Observac As String, cadErr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim ImporIva As Currency

    On Error GoTo EInsertar
    
    sql = "SELECT stipom.letraser,numfactu,fecfactu, rsocios_seccion.codmaccli as codmacta,year(fecfactu) as anofaccl,tipoiva,porc_iva,"
    sql = sql & "sum(baseimpo + coalesce(imprecargo,0)) baseimpo "
    sql = sql & " FROM ((" & "rrecibpozos inner join " & "usuarios.stipom on rrecibpozos.codtipom=usuarios.stipom.codtipom) "
    sql = sql & "INNER JOIN rsocios ON rrecibpozos.codsocio=rsocios.codsocio) "
    sql = sql & "INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionPOZOS, "N")
    sql = sql & " WHERE " & cadwhere
    sql = sql & " group by 1,2,3,4,5,6,7 "
    sql = sql & " order by 1,2,3,4,5,6,7 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = 0
        DtoGnral = 0
        BaseImp = Rs!baseimpo
        '---- Laura 10/10/2006:  añadir el totalfac para utilizarlo en insertar lineas
        ImporIva = Round2((DBLet(BaseImp, "N") * DBLet(Rs!porc_iva, "N") / 100), 2)
        TotalFac = BaseImp + ImporIva
        '----
        
        sql = ""
        sql = "'" & Rs!LetraSer & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!fecfactu) & "," & DBSet(Observac, "T") & ","
        sql = sql & DBSet(Rs!baseimpo, "N") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!porc_iva, "N") & "," & ValorNulo & "," & ValorNulo & ","
        sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(ImporIva, "N", "N") & "," & ValorNulo & "," & ValorNulo & ","
        sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        sql = sql & DBSet(TotalFac, "N") & "," & DBSet(Rs!TipoIVA, "N") & "," & ValorNulo & "," & ValorNulo & ",0,"
        sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        sql = sql & DBSet(Rs!fecfactu, "F")
        Cad = Cad & "(" & sql & ")"
'        RS.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    'Insertar en la contabilidad
    sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
    sql = sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
    sql = sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
    sql = sql & " VALUES " & Cad
    ConnConta.Execute sql
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactPOZ = False
        cadErr = Err.Description
    Else
        InsertarCabFactPOZ = True
    End If
End Function



Private Function InsertarLinFactPOZ(cadTabla As String, cadwhere As String, cadErr As String, tipoMov As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
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
            sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu,sum(round(precio1*consumo1,2)) as importeconsumo,sum(round(precio2*consumo2,2) + impcuota) as importecuota, " & DBSet(vParamAplic.CodCCostPOZ, "T") & " as codccost "
        Else
            sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu,sum(round(precio1*consumo1,2)) as importeconsumo,sum(round(precio2*consumo2,2) + impcuota) as importecuota "
        End If
        
        sql = sql & " FROM rrecibpozos inner join usuarios.stipom on rrecibpozos.codtipom=usuarios.stipom.codtipom "
        sql = sql & " WHERE " & cadwhere
        
        If vEmpresa.TieneAnalitica Then
            sql = sql & " GROUP BY 1,2,3,4,7 " '& cadCampo, codccost
        Else
            sql = sql & " GROUP BY 1,2,3,4" '& cadCampo
        End If
        
        Set Rs = New ADODB.Recordset
        Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        Cad = ""
        I = 1
        totimp = 0
        SQLaux = ""
        If Not Rs.EOF Then
            SQLaux = Cad
            
            ImpConsumo = DBLet(Rs!Importeconsumo, "N")
            ImpCuota = DBLet(Rs!importecuota, "N")
            totimp = totimp + ImpConsumo + ImpCuota
    
            'concatenamos linea para insertar en la tabla de conta.linfact
            sql = ""
            Sql2 = ""
            
            
            If ImpConsumo <> 0 Then
                sql = "'" & Rs!LetraSer & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & DBSet(I, "N") & "," & DBSet(vParamAplic.CtaVentasConsPOZ, "T") & ","
                
                Sql2 = Cad & sql  'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
                sql = sql & DBSet(ImpConsumo, "N") & ","
                
                If vEmpresa.TieneAnalitica Then
                    sql = sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBSet(Rs!CodCCost, "T")
                Else
                    sql = sql & ValorNulo
                    CCoste = ValorNulo
                End If
                
                Cad = "(" & sql & ")" & ","
                
                SQLaux = SQLaux & Cad
                
                I = I + 1
            End If
            
            
            If ImpCuota <> 0 Then
                sql = "('" & Rs!LetraSer & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & DBSet(I, "N") & "," & DBSet(vParamAplic.CtaVentasCuoPOZ, "T") & ","
                
                Sql2 = Cad & sql   'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
                sql = sql & DBSet(ImpCuota, "N") & ","
                
                If vEmpresa.TieneAnalitica Then
                    sql = sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBSet(Rs!CodCCost, "T")
                Else
                    sql = sql & ValorNulo
                    CCoste = ValorNulo
                End If
                
                Cad = Cad & sql & ")" & ","
                
                SQLaux = SQLaux & Cad
            End If
            
            Rs.MoveNext
        End If
        
        Rs.Close
        Set Rs = Nothing
        
        BaseImp = DevuelveValor("select sum(baseimpo) from rrecibpozos where " & cadwhere)
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
            Cad = Sql2 & "),"
        End If
    
    
        'Insertar en la contabilidad
        If Cad <> "" Then
            Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
            sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
            sql = sql & " VALUES " & Cad
            ConnConta.Execute sql
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
            sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(baseimpo+coalesce(imprecargo,0)) as importe, " & DBSet(vParamAplic.CodCCostPOZ, "T")
        Else
            sql = " SELECT usuarios.stipom.letraser,rrecibpozos.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(baseimpo+coalesce(imprecargo,0)) as importe "
        End If
        
        sql = sql & " FROM rrecibpozos inner join usuarios.stipom on rrecibpozos.codtipom=usuarios.stipom.codtipom "
        sql = sql & " WHERE " & cadwhere
        If vEmpresa.TieneAnalitica Then
            sql = sql & " GROUP BY 1,2,3,4,6 " '& cadCampo, codccost
        Else
            sql = sql & " GROUP BY 1,2,3,4 " '& cadCampo
        End If
        
        
        Set Rs = New ADODB.Recordset
        Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        Cad = ""
        I = 1
        totimp = 0
        SQLaux = ""
        While Not Rs.EOF
            SQLaux = Cad
            'calculamos la Base Imp del total del importe para cada cta cble ventas
            ImpLinea = DBLet(Rs!Importe, "N")
            totimp = totimp + DBLet(Rs!Importe, "N")
    
            'concatenamos linea para insertar en la tabla de conta.linfact
            sql = ""
            Sql2 = ""
            
            sql = "'" & Rs!LetraSer & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
            sql = sql & DBSet(Rs!cuenta, "T")
            
            Sql2 = sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
            sql = sql & "," & DBSet(ImpLinea, "N") & ","
            
            If vEmpresa.TieneAnalitica Then
                sql = sql & DBSet(Rs!CodCCost, "T")
                CCoste = DBSet(Rs!CodCCost, "T")
            Else
                sql = sql & ValorNulo
                CCoste = ValorNulo
            End If
            
            Cad = Cad & "(" & sql & ")" & ","
            
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
                Cad = SQLaux & "(" & Sql2 & ")" & ","
            Else 'solo una linea
                Cad = "(" & Sql2 & ")" & ","
            End If
            
    '        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
    '        cad = Replace(cad, SQL, Aux)
        End If
    
    
        'Insertar en la contabilidad
        If Cad <> "" Then
            Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
            sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
            sql = sql & " VALUES " & Cad
            ConnConta.Execute sql
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

'###########################CONTABILIZACION DE FACTURAS DE TRANSPORTE INTERNAS


Private Function InsertarAsientoDiarioTRANS(cadwhere As String, cadErr As String, ByRef Mc As Contadores, FechaFin As Date, Secci As String, Tipo As Byte, FecRec As Date, TipoM As String) As Boolean
' la contabilizacion de las facturas internas de horto, se han de insertar en el diario no en el registro de iva de proveedor
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Concepto As String
Dim LetraSer As String
Dim Obs As String
Dim cadMen As String
Dim b As Boolean
'Dim CtaSocio As String


    On Error GoTo EInsertar
       
    sql = " SELECT fecfactu,year(" & DBSet(FecRec, "F") & ") as anofacpr," & DBSet(FecRec, "F") & ",numfactu,rtransporte.codmacpro,"
    sql = sql & "baseimpo, tipoiva, porc_iva,imporiva,basereten,porc_ret,impreten,baseaport,porc_apo,impapor,totalfac,"
    sql = sql & "rectif_codtipom, rectif_numfactu, rectif_fecfactu,"
    sql = sql & "rtransporte.codtrans, rtransporte.nomtrans, rtransporte.codbanco, rtransporte.codsucur, rtransporte.digcontr, rtransporte.cuentaba "
    sql = sql & ",rtransporte.iban "
    sql = sql & " FROM (" & "rfacttra "
    sql = sql & "INNER JOIN rtransporte ON rfacttra.codtrans=rtransporte.codtrans) "
    sql = sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        If Mc.ConseguirContador("1", (FecRec <= CDate(FechaFin) - 365), True) = 0 Then
        
            BaseImp = DBLet(Rs!baseimpo, "N")
            TotalFac = BaseImp + DBLet(Rs!ImporIva, "N")
            AnyoFacPr = Rs!anofacpr
            
            ImpReten = DBLet(Rs!ImpReten, "N")
            
            LetraSer = ""
            LetraSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(TipoM, "T"))
            
            FacturaTRA = LetraSer & "-" & DBLet(Rs!numfactu, "N")
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
            b = InsertarCabAsientoDia(vEmpresa.NumDiarioInt, Mc.Contador, CStr(Format(FecRecep, "dd/mm/yyyy")), Obs, cadMen)
            cadMen = "Insertando Cab. Asiento: " & cadMen
    
            If b Then
                b = InsertarLinAsientoFactIntTRA("rfacttra", cadwhere, cadMen, Tipo, CtaTransporte, Mc.Contador)
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





Private Function InsertarLinAsientoFactIntTRA(cadTabla As String, cadwhere As String, cadErr As String, Tipo As Byte, CtaSocio As String, Optional Contador As Long) As Boolean
Dim sql As String
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
Dim b As Boolean
Dim Cad As String
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
    NumFact = DevuelveValor("select numfactu from rfacttra where " & cadwhere)
    
    If vEmpresa.TieneAnalitica Then
        sql = " SELECT 1, variedades.ctatransporte as cuenta, sum(rfacttra_albaran.importe) as importe, variedades.codccost "
    Else
        sql = " SELECT 1, variedades.ctatransporte as cuenta, sum(rfacttra_albaran.importe) as importe "
    End If
    sql = sql & " FROM rfacttra_albaran, variedades "
    sql = sql & " WHERE " & Replace(cadwhere, "rfacttra", "rfacttra_albaran") & " and"
    sql = sql & " rfacttra_albaran.codvarie = variedades.codvarie "
    sql = sql & " group by 1,2 "
    sql = sql & " order by 1,2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    I = 0
    ImporteD = 0
    ImporteH = 0
    
    numdocum = Format(NumFact, "0000000")
    Ampliacion = FacturaTRA
    ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoIntProv, "N")) & " " & Ampliacion
    ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoIntProv, "N")) & " " & Ampliacion
    
    If Not Rs.EOF Then Rs.MoveFirst
    
    b = True

    Cad = ""
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = Cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        ImpLinea = Rs!Importe
        
        totimp = totimp + ImpLinea
        
        I = I + 1
        
        Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
        Cad = Cad & DBSet(I, "N") & "," & DBSet(Rs!cuenta, "T") & "," & DBSet(numdocum, "T") & ","
        
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If Rs.Fields(2).Value > 0 Then
            ' importe al debe en positivo
            Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Rs.Fields(2).Value, "N") & ","
            Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaTransporte, "T") & "," & ValorNulo & ",0"
        
            ImporteD = ImporteD + (CCur(Rs.Fields(2).Value))
        Else
            ' importe al haber en positivo, cambiamos el signo
            Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
            Cad = Cad & DBSet((Rs.Fields(2).Value) * (-1), "N") & "," & ValorNulo & "," & DBSet(CtaTransporte, "T") & "," & ValorNulo & ",0"
        
            ImporteH = ImporteH + (CCur(Rs.Fields(2).Value) * (-1))
        End If
        
        Cad = "(" & Cad & ")"
        
        b = InsertarLinAsientoDia(Cad, cadMen)
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
            sql = "update linapu set timporteD = " & DBSet(totimp, "N")
            sql = sql & " where numdiari = " & DBSet(vEmpresa.NumDiarioInt, "N")
            sql = sql & " and fechaent = " & DBSet(FecRecep, "F")
            sql = sql & " and numasien = " & DBSet(Contador, "N")
            sql = sql & " and linliapu = " & DBSet(I, "N")
            
            ConnConta.Execute sql
        Else
            sql = "update linapu set timporteH = " & DBSet(totimp, "N")
            sql = sql & " where numdiari = " & DBSet(vEmpresa.NumDiarioInt, "N")
            sql = sql & " and fechaent = " & DBSet(FecRecep, "F")
            sql = sql & " and numasien = " & DBSet(Contador, "N")
            sql = sql & " and linliapu = " & DBSet(I, "N")
            
            ConnConta.Execute sql
        End If
    End If

    If b And I > 0 Then
        I = I + 1
        
        ' el Total es sobre la cuenta del transportista
        Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
        Cad = Cad & DBSet(I, "N") & ","
        Cad = Cad & DBSet(CtaTransporte, "T") & "," & DBSet(numdocum, "T") & ","
            
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If ImporteD - ImporteH < 0 Then
            ' importe al debe en positivo
            Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet((ImporteD - ImporteH) * (-1), "N") & ","
            Cad = Cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
        Else
            'importe al haber en positivo, cambiamos el signo
            Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
            Cad = Cad & DBSet((ImporteD - ImporteH), "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
        End If
        
        Cad = "(" & Cad & ")"
        
        b = InsertarLinAsientoDia(Cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & I
        
    End If

    If b Then
        ' las retenciones si las hay
        If ImpReten <> 0 Then
            I = I + 1
            
            Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
            Cad = Cad & DBSet(I, "N") & "," & DBSet(CtaTransporte, "T") & "," & DBSet(numdocum, "T") & ","
        
            If ImpReten > 0 Then
                ' importe al debe en positivo
                Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(ImpReten, "N") & ","
                Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaReten, "T") & "," & ValorNulo & ",0"
            Else
                ' importe al haber en positivo, cambiamos el signo
                Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                Cad = Cad & DBSet((ImpReten * (-1)), "N") & "," & ValorNulo & "," & DBSet(CtaReten, "T") & "," & ValorNulo & ",0"
            
            End If
            
            Cad = "(" & Cad & ")"
            
            b = InsertarLinAsientoDia(Cad, cadMen)
            cadMen = "Insertando Lin. Asiento: " & I
            
            If b Then
                I = I + 1
                
                Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
                Cad = Cad & DBSet(I, "N") & "," & DBSet(CtaReten, "T") & "," & DBSet(numdocum, "T") & ","
                If ImpReten > 0 Then
                    ' importe al haber en positivo
                    Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & ValorNulo & ","
                    Cad = Cad & DBSet(ImpReten, "N") & "," & ValorNulo & "," & DBSet(CtaTransporte, "T") & "," & ValorNulo & ",0"
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & DBSet((ImpReten * (-1)), "N") & ","
                    Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaTransporte, "T") & "," & ValorNulo & ",0"
                End If
                
                Cad = "(" & Cad & ")"
                
                b = InsertarLinAsientoDia(Cad, cadMen)
                cadMen = "Insertando Lin. Asiento: " & I
            End If
            
        End If
    End If
    
    
    If b Then
        ' las aportaciones de fondo operativo si las hay
        If ImpAport <> 0 Then
            I = I + 1
            
            Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
            Cad = Cad & DBSet(I, "N") & "," & DBSet(CtaTransporte, "T") & "," & DBSet(numdocum, "T") & ","
        
            If ImpAport > 0 Then
                ' importe al debe en positivo
                Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(ImpAport, "N") & ","
                Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaAport, "T") & "," & ValorNulo & ",0"
            Else
                ' importe al haber en positivo, cambiamos el signo
                Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                Cad = Cad & DBSet((ImpAport * (-1)), "N") & "," & ValorNulo & "," & DBSet(CtaAport, "T") & "," & ValorNulo & ",0"
            
            End If
            
            Cad = "(" & Cad & ")"
            
            b = InsertarLinAsientoDia(Cad, cadMen)
            cadMen = "Insertando Lin. Asiento: " & I
            
            If b Then
                I = I + 1
                
                Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FecRecep, "F") & "," & DBSet(Contador, "N") & ","
                Cad = Cad & DBSet(I, "N") & "," & DBSet(CtaAport, "T") & "," & DBSet(numdocum, "T") & ","
                If ImpAport > 0 Then
                    ' importe al haber en positivo
                    Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliaciond, "T") & "," & ValorNulo & ","
                    Cad = Cad & DBSet(ImpAport, "N") & "," & ValorNulo & "," & DBSet(CtaTransporte, "T") & "," & ValorNulo & ",0"
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    Cad = Cad & DBSet(vEmpresa.ConceptoIntProv, "N") & "," & DBSet(ampliacionh, "T") & "," & DBSet((ImpAport * (-1)), "N") & ","
                    Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaTransporte, "T") & "," & ValorNulo & ",0"
                End If
                
                Cad = "(" & Cad & ")"
                
                b = InsertarLinAsientoDia(Cad, cadMen)
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
    InsertarLinAsientoFactIntTRA = b
    Exit Function
End Function




Public Function PasarAsientoGastoCampo(cadwhere As String, FechaFin As String, FecRecep As String, CtaContra As String, Concep As String, Amplia As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariagro.rfactsoc --> conta.cabfactprov
' ariagro.rfactsoc_variedad --> conta.linfactprov
'Actualizar la tabla ariagro.rfactsoc.contabilizada=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim sql As String
Dim Mc As Contadores


    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    Set Mc = New Contadores
    
    ' Insertamos en el diario
    b = InsertarAsientoGastoCampo(cadwhere, cadMen, Mc, CDate(FechaFin), CDate(FecRecep), CtaContra, Concep, Amplia)
    cadMen = "Insertando Asiento de Gasto en Diario: " & cadMen
    
    If b Then
        '---- Poner contabilizado=1 en rcampos_gastos
        b = ActualizarCabFactSoc("rcampos_gastos", cadwhere, cadMen)
        cadMen = "Actualizando Concepto Gasto Campo: " & cadMen
    End If
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Asiento Gasto de Campo", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarAsientoGastoCampo = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarAsientoGastoCampo = False
        If Not b Then
            InsertarTMPErrFacSoc cadMen, cadwhere
        End If
    End If
End Function



Private Function InsertarAsientoGastoCampo(cadwhere As String, cadErr As String, ByRef Mc As Contadores, FechaFin As Date, FecRec As Date, CtaContra As String, Concep As String, Amplia As String) As Boolean
' la contabilizacion de las facturas internas de horto, se han de insertar en el diario no en el registro de iva de proveedor
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Concepto As String
Dim LetraSer As String
Dim Obs As String
Dim cadMen As String
Dim b As Boolean
'Dim CtaSocio As String


    On Error GoTo EInsertar
       
    sql = " SELECT rcampos_gastos.codgasto, rcampos_gastos.fecha, rcampos_gastos.importe, rconcepgasto.codmacgto "
    sql = sql & " FROM (rcampos_gastos "
    sql = sql & "INNER JOIN rconcepgasto ON rcampos_gastos.codgasto=rconcepgasto.codgasto) "
    sql = sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        If Mc.ConseguirContador("1", (FecRec <= CDate(FechaFin) - 365), True) = 0 Then
        
            Obs = "Contabilización Gasto de Campo de Fecha " & Format(FecRec, "dd/mm/yyyy")
        
            'Insertar en la conta Cabecera Asiento
            b = InsertarCabAsientoDia(vEmpresa.NumDiarioInt, Mc.Contador, CStr(Format(FecRec, "dd/mm/yyyy")), Obs, cadMen)
            cadMen = "Insertando Cab. Asiento: " & cadMen
    
            If b Then
                b = InsertarLinAsientoDiaGastos("rcampos_gastos", cadwhere, cadMen, CtaContra, Mc.Contador, Concep, Amplia)
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


Private Function InsertarLinAsientoDiaGastos(cadTabla As String, cadwhere As String, cadErr As String, CtaContra As String, Contador As Long, Concep As String, Amplia As String) As Boolean
Dim sql As String
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
Dim b As Boolean
Dim Cad As String
Dim cadMen As String
Dim FeFact As Date

Dim cadCampo As String

    On Error GoTo eInsertarLinAsientoDiaGastos

    InsertarLinAsientoDiaGastos = False

    sql = " SELECT rcampos_gastos.fecha, rcampos_gastos.codcampo, rconcepgasto.codmacgto cuenta, rcampos_gastos.importe as importe "
    sql = sql & " FROM rcampos_gastos Inner JOIN rconcepgasto ON rcampos_gastos.codgasto = rconcepgasto.codgasto "
    sql = sql & " where " & cadwhere
    sql = sql & " order by 1, 2 " '& cadCampo

    
    Set Rs = New ADODB.Recordset
    
    Rs.Open sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
            
    I = 0
    ImporteD = 0
    ImporteH = 0
    
    numdocum = Format(Rs!codcampo, "00000000")
'    Ampliacion = Format(Rs!codcampo, "00000000")
    ampliaciond = Amplia 'Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoInt, "N")) & " " & Ampliacion
    ampliacionh = Amplia 'Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoInt, "N")) & " " & Ampliacion
    
    If Not Rs.EOF Then Rs.MoveFirst
    
    b = True
    
    If Not Rs.EOF Then
        I = I + 1
        
        FeFact = Rs!Fecha
        
        Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(Rs!Fecha, "F") & "," & DBSet(Contador, "N") & ","
        Cad = Cad & DBSet(I, "N") & "," & DBSet(Rs!cuenta, "T") & "," & DBSet(numdocum, "T") & ","
        
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If DBLet(Rs!Importe, "N") > 0 Then
            ' importe al debe en positivo
            Cad = Cad & DBSet(Concep, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Rs!Importe, "N") & ","
            Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaContra, "T") & "," & ValorNulo & ",0"
        
            ImporteD = ImporteD + CCur(DBLet(Rs!Importe, "N"))
        Else
            ' importe al haber en positivo, cambiamos el signo
            Cad = Cad & DBSet(Concep, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
            Cad = Cad & DBSet(Rs!Importe, "N") & "," & ValorNulo & "," & DBSet(CtaContra, "T") & "," & ValorNulo & ",0"
        
            ImporteH = ImporteH + CCur(DBLet(Rs!Importe, "N"))
        End If
        
        Cad = "(" & Cad & ")"
        
        b = InsertarLinAsientoDia(Cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & I

        I = I + 1
                
        ' el Total es sobre la cuenta del cliente
        Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(Rs!Fecha, "F") & "," & DBSet(Contador, "N") & ","
        Cad = Cad & DBSet(I, "N") & ","
        Cad = Cad & DBSet(CtaContra, "T") & "," & DBSet(numdocum, "T") & ","
            
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If DBLet(Rs!Importe, "N") > 0 Then
            ' importe al haber en positivo, cambiamos el signo
            Cad = Cad & DBSet(Concep, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
            Cad = Cad & DBSet(Rs!Importe, "N") & "," & ValorNulo & "," & DBSet(Rs!cuenta, "N") & "," & ValorNulo & ",0"
        Else
            ' importe al debe en positivo
            Cad = Cad & DBSet(Concep, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(DBLet(Rs!Importe, "N") * (-1), "N") & ","
            Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!cuenta, "N") & "," & ValorNulo & ",0"
        
        End If
        
        Cad = "(" & Cad & ")"
        
        b = InsertarLinAsientoDia(Cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & I

    End If
    
    Set Rs = Nothing
    InsertarLinAsientoDiaGastos = b
    Exit Function
    
eInsertarLinAsientoDiaGastos:
    cadErr = "Insertar Linea Asiento Gastos: " & Err.Description
    cadErr = cadErr & cadMen
    InsertarLinAsientoDiaGastos = False
End Function


'----------------------------------------------------------------------
' FACTURAS VARIAS REGISTRO CLIENTE
'----------------------------------------------------------------------
Public Function PasarFacturaFVAR(cadwhere As String, CodCCost As String, FechaFin As String, Seccion As String, TipoFact As Byte, FecVto As Date, ForpaPos As String, ForpaNeg As String, CtaBanc As String, TipoM As String, Optional FecRecep As Date) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariagro.rfactsoc --> conta.cabfactprov
' ariagro.rfactsoc_variedad --> conta.linfactprov
'Actualizar la tabla ariagro.rfactsoc.contabilizada=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim sql As String
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
    
    ImpReten = 0
    CtaReten = ""
        
    If TipoM = "FVG" Then
        b = True
        ' tendriamos que insertar en el diario FALTA
    Else
        If TipoM = "FVP" Then 'registro de iva de proveedor
            b = InsertarCabFactFVARPro(cadwhere, cadMen, Mc, CDate(FechaFin), Seccion, CStr(FecRecep))
            cadMen = "Insertando Cab. Factura Proveedor: " & cadMen
        Else ' registro de iva de cliente
            '---- Insertar en la conta Cabecera Factura
            b = InsertarCabFactFVAR(cadwhere, cadMen, TipoFact, Seccion)
            cadMen = "Insertando Cab. Factura: " & cadMen
        End If
    End If
    
    If b Then
        FecVenci = FecVto
        ForpaPosi = ForpaPos
        ForpaNega = ForpaNeg
        CtaBanco = CtaBanc
        tipoMov = TipoM    ' codtipom de la factura de socio
        
        If TipoM = "FVP" Then ' registro de iva de proveedor
            b = InsertarEnTesoreriaNewFVARPro(cadwhere, cadMen, CtaBanco, CStr(FecVenci))
            cadMen = "Insertando en Tesoreria: " & cadMen
        Else
            'si la factura es a un cliente o de socio a no descontar en liquidacion , se inserta en tesoreria
            If TipoFact = 1 Or (TipoFact = 0 And Not FraADescontarEnLiquidacion(cadwhere)) Then
                b = InsertarEnTesoreriaNewFVAR(cadwhere, CtaBanco, CStr(FecVenci), cadMen, TipoFact, Seccion)
                cadMen = "Insertando en Tesoreria: " & cadMen
            End If
        End If
        If b Then
            If TipoM = "FVP" Then ' registro de iva de proveedores
                CCoste = CodCCost
                '---- Insertar lineas de Factura en la Conta
                b = InsertarLinFactFVAR("fvarcabfactpro", cadwhere, cadMen, Mc.Contador)
                cadMen = "Insertando Lin. Factura: " & cadMen
            Else
                If TipoM <> "FVG" Then
                    CCoste = CodCCost
                    '---- Insertar lineas de Factura en la Conta
                    b = InsertarLinFactFVAR("fvarcabfact", cadwhere, cadMen)
                    cadMen = "Insertando Lin. Factura: " & cadMen
                End If
            End If
            
            If b Then
                '---- Poner intconta=1 en ariges.scafac
                If TipoM = "FVP" Then ' registro de iva de proveedores
                    b = ActualizarCabFact("fvarcabfactpro", cadwhere, cadMen)
                Else
                    b = ActualizarCabFact("fvarcabfact", cadwhere, cadMen)
                End If
                cadMen = "Actualizando Factura Varia: " & cadMen
            End If
        End If
    End If
    
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Facturas Varias", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaFVAR = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaFVAR = False
        If Not b Then
            InsertarTMPErrFacFVAR cadMen, cadwhere
        End If
    End If
End Function


Private Function InsertarCabFactFVAR(cadwhere As String, cadErr As String, Tipo As Byte, vSeccion As String) As Boolean
'Insertando en tabla conta.cabfact
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Seccion As Integer

    On Error GoTo EInsertar
    
    ' factura de cliente (socio)
    If Tipo = 0 Then
        sql = "SELECT stipom.letraser,numfactu,fecfactu, rsocios_seccion.codmaccli as codmacta,year(fecfactu) as anofaccl,"
        sql = sql & "baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
        sql = sql & "totalfac,tipoiva1,tipoiva2,tipoiva3, porcrec1, porcrec2, porcrec3, imporec1, imporec2, imporec3, "
        sql = sql & "retfaccl, trefaccl, cuereten "
        sql = sql & " FROM ((" & "fvarcabfact inner join " & "usuarios.stipom on fvarcabfact.codtipom=usuarios.stipom.codtipom) "
        sql = sql & "INNER JOIN rsocios ON fvarcabfact.codsocio=rsocios.codsocio) "
        sql = sql & "INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(vSeccion, "N")
        sql = sql & " WHERE " & cadwhere
    Else
    ' factura de cliente (cliente)
        sql = "SELECT stipom.letraser,numfactu,fecfactu, clientes.codmacta as codmacta,year(fecfactu) as anofaccl,"
        sql = sql & "baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
        sql = sql & "totalfac,tipoiva1,tipoiva2,tipoiva3, porcrec1, porcrec2, porcrec3, imporec1, imporec2, imporec3, "
        sql = sql & "retfaccl, trefaccl, cuereten "
        sql = sql & " FROM ((" & "fvarcabfact inner join " & "usuarios.stipom on fvarcabfact.codtipom=usuarios.stipom.codtipom) "
        sql = sql & "INNER JOIN clientes ON fvarcabfact.codclien=clientes.codclien) "
        sql = sql & " WHERE " & cadwhere
    End If
        
        
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = 0
        DtoGnral = 0
        BaseImp = Rs!BaseIVA1 + CCur(DBLet(Rs!BaseIVA2, "N")) + CCur(DBLet(Rs!BaseIVA3, "N"))
        '---- Laura 10/10/2006:  añadir el totalfac para utilizarlo en insertar lineas
        TotalFac = Rs!TotalFac
        '----
        
        sql = ""
        sql = "'" & Rs!LetraSer & "'," & Rs!numfactu & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!fecfactu) & "," & ValorNulo & ","
        sql = sql & DBSet(Rs!BaseIVA1, "N") & "," & DBSet(Rs!BaseIVA2, "N", "S") & "," & DBSet(Rs!BaseIVA3, "N", "S") & "," & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", "S") & "," & DBSet(Rs!porciva3, "N", "S") & ","
        sql = sql & DBSet(Rs!porcrec1, "N", "S") & "," & DBSet(Rs!porcrec2, "N", "S") & "," & DBSet(Rs!porcrec3, "N", "S") & "," & DBSet(Rs!ImpoIva1, "N", "N") & "," & DBSet(Rs!impoIVA2, "N", "S") & "," & DBSet(Rs!impoIVA3, "N", "S") & ","
        sql = sql & DBSet(Rs!imporec1, "N", "S") & "," & DBSet(Rs!imporec2, "N", "S") & "," & DBSet(Rs!imporec3, "N", "S") & ","
        sql = sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!TipoIVA1, "N") & "," & DBSet(Rs!TipoIVA2, "N", "S") & "," & DBSet(Rs!TipoIVA3, "N", "S") & ",0,"
        sql = sql & DBSet(Rs!retfaccl, "N", "S") & "," & DBSet(Rs!trefaccl, "N", "S") & "," & DBSet(Rs!cuereten, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        sql = sql & DBSet(Rs!fecfactu, "F")
        Cad = Cad & "(" & sql & ")"
'        RS.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
        
    'Insertar en la contabilidad
    sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
    sql = sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
    sql = sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
    sql = sql & " VALUES " & Cad
    ConnConta.Execute sql
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactFVAR = False
        cadErr = Err.Description
    Else
        InsertarCabFactFVAR = True
    End If
End Function



Public Function InsertarEnTesoreriaNewFVAR(cadwhere As String, CtaBan As String, FecVen As String, MenError As String, Tipo As Byte, vSeccion As String) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
Dim b As Boolean
Dim sql As String, Text33csb As String, Text41csb As String
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
Dim LetraSer As String
Dim Rsx As ADODB.Recordset
Dim FecVenci As Date
Dim ImpVenci As Currency
Dim ImpVenci1 As Currency
Dim AcumIva As Currency
Dim PorcIva As String

Dim Rsx7 As ADODB.Recordset
Dim Sql7 As String
Dim Cadena As String

Dim CadRegistro As String
Dim CadRegistro1 As String

Dim vSocio As cSocio

    On Error GoTo EInsertarTesoreriaNewFac

    b = False
    InsertarEnTesoreriaNewFVAR = False
    
    CadValues = ""
    CadValues2 = ""
    
    Set Rsx = New ADODB.Recordset
    sql = "select * from fvarcabfact where " & cadwhere
    Rsx.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    If Not Rsx.EOF Then
        LetraSer = ""
        LetraSer = ObtenerLetraSerie2(Rsx!CodTipom)
    
        If Tipo = 0 Then
            ' socio
            
            Dim vSoc As cSocio
            Set vSoc = New cSocio
            
            
            If vSoc.LeerDatos(DBLet(Rsx!Codsocio, "N")) Then
                If vSoc.LeerDatosSeccion(DBLet(Rsx!Codsocio, "N"), CStr(vSeccion)) Then
                    b = True
                            
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
            Sql4 = "select codbanco, codsucur, digcontr, cuentaba, codmacta, iban from clientes where codclien = " & DBLet(Rsx!CodClien, "N")
            Set Rs4 = New ADODB.Recordset
            
            Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not Rs4.EOF Then
                b = True
                
                CC = DBLet(Rs4!digcontr, "T")
                If DBLet(Rs4!digcontr, "T") = "**" Then CC = "00"
                
                Iban = DBLet(Rs4!Iban, "T")
                CodBanco = DBLet(Rs4!CodBanco, "N")
                CodSucur = DBLet(Rs4!CodSucur, "N")
                CuentaBa = DBLet(Rs4!CuentaBa, "T")
                Codmacta = DBLet(Rs4!Codmacta, "T")
            End If
        End If
            
        If b Then
            Text33csb = "'Factura:" & DBLet(LetraSer, "T") & "-" & DBLet(Rsx!numfactu, "T") & " " & Format(DBLet(Rsx!fecfactu, "F"), "dd/mm/yy") & "'"
            Text41csb = "de " & DBSet(Rsx!TotalFac, "N")
            
            'Obtener el Nº de Vencimientos de la forma de pago
            sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & DBSet(Rsx!codforpa, "N")
            Set rsVenci = New ADODB.Recordset
            rsVenci.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

            If Not rsVenci.EOF Then
                If DBLet(rsVenci!numerove, "N") > 0 Then
            
                    CadValuesAux2 = "('" & Trim(LetraSer) & "', " & DBSet(Rsx!numfactu, "N") & ", " & DBSet(Rsx!fecfactu, "F") & ", "
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
                    CadValues2 = CadValues2 & DBSet(Trim(Codmacta), "T") & ", " & DBSet(Rsx!codforpa, "N") & ", " & DBSet(FecVenci, "F") & ", "
                    
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
                    
                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(CodBanco, "N", "S") & ", " & DBSet(CodSucur, "N", "S") & ", " & DBSet(CC, "T", "S") & ", " & DBSet(CuentaBa, "T", "S") & ", "
                    CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & Text33csb & "," & DBSet(Text41csb, "T") & ",1" '),"
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        CadValues2 = CadValues2 & ", " & DBSet(Iban, "T", "S") & "),"
                    Else
                        CadValues2 = CadValues2 & "),"
                    End If
                    
                
                    'Resto Vencimientos
                    '--------------------------------------------------------------------
                    For I = 2 To rsVenci!numerove
                       'FECHA Resto Vencimientos
                        '=== Laura 23/01/2007
                        'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                        FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                        '===
                            
                        CadValues2 = CadValues2 & CadValuesAux2 & I & ", " & DBSet(Trim(Rs4!Codmacta), "T") & ", " & DBSet(Rsx!codforpa, "N") & ", '" & Format(FecVenci, FormatoFecha) & "', "
                        
                        'IMPORTE Resto de Vendimientos
                        ImpVenci = Round2(TotalFac / rsVenci!numerove, 2)
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CtaBan, "T") & ", " & DBSet(Rs4!CodBanco, "N", "S") & ", " & DBSet(Rs4!CodSucur, "N", "S") & ", " & DBSet(CC, "T", "S") & ", " & DBSet(Rs4!CuentaBa, "T", "S") & ", "
                        CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & Text33csb & "," & DBSet(Text41csb, "T") & ",1" '),"
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & ", " & DBSet(Iban, "T", "S") & "),"
                        Else
                            CadValues2 = CadValues2 & "),"
                        End If
                        
                    Next I
                    ' quitamos la ultima coma
                    CadValues2 = Mid(CadValues2, 1, Len(CadValues2) - 1)
                        
                    'Insertamos en la tabla scobro de la CONTA
                    sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                    sql = sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                    sql = sql & " text33csb, text41csb, agente" ') "
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        sql = sql & ", iban) "
                    Else
                        sql = sql & ") "
                    End If
                    sql = sql & " VALUES " & CadValues2
                    ConnConta.Execute sql
                
                End If
            End If
        
            b = True

        End If
    
    End If
    
EInsertarTesoreriaNewFac:
    If Err.Number <> 0 Then
        b = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaNewFVAR = b
End Function





Private Function InsertarLinFactFVAR(cadTabla As String, cadwhere As String, cadErr As String, Optional NumRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim sql As String
Dim SQLaux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Cad As String, Aux As String
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
            sql = " SELECT usuarios.stipom.letraser,fvarlinfact.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importe) as importe, fvarconce.codccost "
        Else
            sql = " SELECT usuarios.stipom.letraser,fvarlinfact.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importe) as importe "
        End If
        
        sql = sql & " FROM (fvarlinfact inner join usuarios.stipom on fvarlinfact.codtipom=usuarios.stipom.codtipom) "
        sql = sql & " inner join fvarconce on fvarlinfact.codconce=fvarconce.codconce "
        sql = sql & " WHERE " & Replace(cadwhere, "fvarcabfact", "fvarlinfact")
    Else
        If vEmpresa.TieneAnalitica Then
            sql = " SELECT usuarios.stipom.letraser,fvarlinfactpro.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importe) as importe, fvarconce.codccost "
        Else
            sql = " SELECT usuarios.stipom.letraser,fvarlinfactpro.codtipom,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importe) as importe "
        End If
        
        sql = sql & " FROM (fvarlinfactpro inner join usuarios.stipom on fvarlinfactpro.codtipom=usuarios.stipom.codtipom) "
        sql = sql & " inner join fvarconce on fvarlinfactpro.codconce=fvarconce.codconce "
        sql = sql & " WHERE " & Replace(cadwhere, "fvarcabfactpro", "fvarlinfactpro")
    End If
    
    If vEmpresa.TieneAnalitica Then
        sql = sql & " GROUP BY 5,7 " '& cadCampo, codccost
    Else
        sql = sql & " GROUP BY 5 " '& cadCampo
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    I = 1
    totimp = 0
    SQLaux = ""
    While Not Rs.EOF
        SQLaux = Cad
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
        sql = ""
        Sql2 = ""
        
        If cadTabla = "fvarcabfact" Then
            sql = "'" & Trim(Rs!LetraSer) & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & I & ","
            sql = sql & DBSet(Trim(Rs!cuenta), "T")
            
        Else
            sql = NumRegis & "," & Year(Rs!fecfactu) & "," & I & ","
            sql = sql & DBSet(Trim(Rs!cuenta), "T")
        
        End If
        
        Sql2 = sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        
        sql = sql & "," & DBSet(ImpLinea, "N") & ","
        
        If vEmpresa.TieneAnalitica Then
            sql = sql & DBSet(Rs!CodCCost, "T")
            CCoste = DBSet(Rs!CodCCost, "T")
        Else
            sql = sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        Cad = Cad & "(" & sql & ")" & ","
        
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
            Cad = SQLaux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            Cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If

    If cadTabla = "fvarcabfactpro" Then
        ' las retenciones si las hay
        If ImpReten <> 0 Then
            sql = NumRegis & "," & AnyoFacPr & "," & I & ","
            sql = sql & DBSet(Trim(CtaSocio), "T")
            sql = sql & "," & DBSet(ImpReten, "N") & ","
            sql = sql & ValorNulo ' no llevan centro de coste
            
            Cad = Cad & "(" & sql & ")" & ","
            I = I + 1
            
            sql = NumRegis & "," & AnyoFacPr & "," & I & ","
            sql = sql & DBSet(Trim(CtaReten), "T")
            sql = sql & "," & DBSet(ImpReten * (-1), "N") & ","
            sql = sql & ValorNulo ' no llevan centro de coste
            
            Cad = Cad & "(" & sql & ")" & ","
            I = I + 1
        End If
    End If


    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        If cadTabla = "fvarcabfact" Then
            sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Else
            sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        End If
        sql = sql & " VALUES " & Cad
        ConnConta.Execute sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactFVAR = False
        cadErr = Err.Description
    Else
        InsertarLinFactFVAR = True
    End If
End Function


Private Function FraADescontarEnLiquidacion(cWhere As String) As Boolean
Dim sql As String
Dim Rs As ADODB.Recordset

    sql = "select enliquidacion from fvarcabfact where " & cWhere
    
    FraADescontarEnLiquidacion = (DevuelveValor(sql) > 0)

End Function




Private Function InsertarCabFactFVARPro(cadwhere As String, cadErr As String, ByRef Mc As Contadores, FechaFin As String, Seccion As String, FecRecep As String) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String

    On Error GoTo EInsertar
       
    
    sql = " SELECT fecfactu," & Year(CDate(FecRecep)) & " as anofacpr,numfactu,rsocios_seccion.codmacpro codmacta,"
    sql = sql & "baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    sql = sql & "totalfac,tipoiva1,tipoiva2,tipoiva3,rsocios_seccion.codsocio, rsocios.nomsocio, "
    sql = sql & "retfaccl, trefaccl, cuereten, rsocios.codbanco, rsocios.codsucur, rsocios.digcontr, rsocios.cuentaba  "
    sql = sql & " FROM (fvarcabfactpro "
    sql = sql & " INNER JOIN rsocios_seccion ON fvarcabfactpro.codsocio=rsocios_seccion.codsocio) "
    sql = sql & " INNER JOIN rsocios ON fvarcabfactpro.codsocio = rsocios.codsocio"
    sql = sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not Rs.EOF Then
    
        If Mc.ConseguirContador("1", (CDate(FecRecep) <= CDate(FechaFin) - 365), True) = 0 Then
        
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            DtoPPago = 0
            DtoGnral = 0
            BaseImp = Rs!BaseIVA1 + CCur(DBLet(Rs!BaseIVA2, "N")) + CCur(DBLet(Rs!BaseIVA3, "N"))
            TotalFac = BaseImp + Rs!ImpoIva1 + CCur(DBLet(Rs!impoIVA2, "N")) + CCur(DBLet(Rs!impoIVA3, "N"))
            AnyoFacPr = Rs!anofacpr
            
            Nulo2 = "N"
            Nulo3 = "N"
            If DBLet(Rs!BaseIVA2, "N") = "0" Then Nulo2 = "S"
            If DBLet(Rs!BaseIVA3, "N") = "0" Then Nulo3 = "S"
            sql = ""
            sql = Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & Rs!anofacpr & "," & DBSet(FecRecep, "F") & "," & DBSet(Rs!numfactu, "T") & "," & DBSet(Trim(Rs!Codmacta), "T") & "," & ValorNulo & ","
            sql = sql & DBSet(Rs!BaseIVA1, "N") & "," & DBSet(Rs!BaseIVA2, "N", "S") & "," & DBSet(Rs!BaseIVA3, "N", "S") & ","
            sql = sql & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", Nulo2) & "," & DBSet(Rs!porciva3, "N", Nulo3) & ","
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!ImpoIva1, "N") & "," & DBSet(Rs!impoIVA2, "N", Nulo2) & "," & DBSet(Rs!impoIVA3, "N", Nulo3) & ","
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            sql = sql & DBSet(TotalFac, "N") & "," & DBSet(Rs!TipoIVA1, "N") & "," & DBSet(Rs!TipoIVA2, "N", Nulo2) & "," & DBSet(Rs!TipoIVA3, "N", Nulo3) & ",0,"
            sql = sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(FecRecep, "F") & ",0"
            Cad = Cad & "(" & sql & ")"
            
            'Insertar en la contabilidad
            sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
            sql = sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
            sql = sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
            sql = sql & " VALUES " & Cad
            ConnConta.Execute sql
            
            'añadido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.Codigo & "," & Mc.Contador
            sql = sql & ",'" & DevNombreSQL(Rs!numfactu) & " @ " & Format(Rs!fecfactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!nomsocio) & "'," & Rs!Codsocio & ")"
            conn.Execute sql
            
            CtaSocio = DBLet(Rs!Codmacta, "T")
            FacturaSoc = DBLet(Rs!numfactu, "N")
            FecFactuSoc = DBLet(Rs!fecfactu)
            
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



Private Function InsertarEnTesoreriaNewFVARPro(cadwhere As String, MenError As String, CtaBanco As String, FecVenci As Date) As Boolean
'Guarda datos de Tesoreria en tablas: spagop o scobro dependiendo del signo de la factura
Dim b As Boolean
Dim sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim CC As String
Dim LetraSer As String
Dim Text33csb As String
Dim Text41csb As String
Dim Text42csb As String
Dim GastosPie As Currency
Dim FactuRec As String

    On Error GoTo EInsertarTesoreria

    InsertarEnTesoreriaNewFVARPro = False
    
    If TotalFac > 0 Then ' se insertara en la cartera de pagos (spagop)
        CadValues2 = ""
    
        'vamos creando la cadena para insertar en spagosp de la CONTA
        LetraSer = ""
        LetraSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(tipoMov, "T"))
        
        '[Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
        
        '[Monica]03/07/2013: añado trim(codmacta)
        CadValuesAux2 = "('" & Trim(CtaSocio) & "', " & DBSet(FacturaSoc, "T") & ", '" & Format(FecFactuSoc, FormatoFecha) & "', "
    
        '------------------------------------------------------------
        I = 1
        CadValues2 = CadValuesAux2 & I
        
        CadValues2 = CadValues2 & ", " & ForpaPosi & ", '" & Format(FecVenci, FormatoFecha) & "', "
        CadValues2 = CadValues2 & DBSet(TotalFac, "N") & ", " & DBSet(CtaBanco, "T") & ","
    
        'David. Para que ponga la cuenta bancaria (SI LA tiene)
        CadValues2 = CadValues2 & DBSet(BancoSoc, "T", "S") & "," & DBSet(SucurSoc, "T", "S") & ","
        CadValues2 = CadValues2 & DBSet(DigcoSoc, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & ","
    
        'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
        sql = "Fact.: " & LetraSer & "-" & FacturaSoc & "-" & Format(FecFactuSoc, "dd/mm/yyyy")
            
        CadValues2 = CadValues2 & "'" & DevNombreSQL(sql) & "',"
        
        sql = ""
        CadValues2 = CadValues2 & "'" & DevNombreSQL(sql) & "'" ')"
        '[Monica]22/11/2013: Tema iban
        If vEmpresa.HayNorma19_34Nueva = 1 Then
            CadValues2 = CadValues2 & ", " & DBSet(IbanSoc, "T", "S") & ") "
        Else
            CadValues2 = CadValues2 & ") "
        End If
        
    
        'Grabar tabla spagop de la CONTABILIDAD
        '-------------------------------------------------
        If CadValues2 <> "" Then
            'Insertamos en la tabla spagop de la CONTA
            'David. Cuenta bancaria y descripcion textos
            sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
            '[Monica]22/11/2013: Tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                sql = sql & ", iban) "
            Else
                sql = sql & ") "
            End If

            sql = sql & " VALUES " & CadValues2
            ConnConta.Execute sql
        End If
        
    Else
        ' si es negativo se inserta en positivo en la cartera de cobros (scobro)

        LetraSer = ""
        LetraSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(tipoMov, "T"))

'                [Monica]27/01/2012: Cogemos el nro de factura recibido si lo hay, antes: letraser & "-" & numfactu
'        Text33csb = "'Factura:" & DBLet(letraser, "T") & "-" & DBLet(numfactu, "T") & " " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
        Text33csb = "'Factura:" & DBLet(FacturaSoc, "T") & " " & Format(DBLet(FecFactuSoc, "F"), "dd/mm/yy") & "'"
        Text41csb = "de " & DBSet(TotalFac, "N")
        Text42csb = ""

        CC = DBLet(DigcoSoc, "T")
        If DBLet(DigcoSoc, "T") = "**" Then CC = "00"
            
        '[Monica]03/07/2013: añado trim(codmacta)
        CadValuesAux2 = "(" & DBSet(LetraSer, "T") & "," & DBSet(FacturaSoc, "N") & "," & DBSet(FecFactuSoc, "F") & ", 1," & DBSet(Trim(CtaSocio), "T") & ","
        CadValues2 = CadValuesAux2 & DBSet(ForpaNega, "N") & "," & DBSet(FecFactuSoc, "F") & "," & DBSet(TotalFac * (-1), "N") & ","
        CadValues2 = CadValues2 & DBSet(CtaBanco, "T") & "," & DBSet(BancoSoc, "N", "S") & "," & DBSet(SucurSoc, "N", "S") & ","
        CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(CtaBaSoc, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & "," & DBSet(Text42csb, "T") & ",1" ')"
        '[Monica]22/11/2013: Tema iban
        If vEmpresa.HayNorma19_34Nueva = 1 Then
            CadValues2 = CadValues2 & "," & DBSet(IbanSoc, "T", "S") & ") "
        Else
            CadValues2 = CadValues2 & ") "
        End If
        

        'Insertamos en la tabla scobro de la CONTA
        sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
        sql = sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
        sql = sql & " text33csb, text41csb, text42csb, agente" ') "
        '[Monica]22/11/2013: Tema iban
        If vEmpresa.HayNorma19_34Nueva = 1 Then
            sql = sql & ", iban) "
        Else
            sql = sql & ") "
        End If
        
        sql = sql & " VALUES " & CadValues2
        ConnConta.Execute sql

    End If

    b = True
    
    
EInsertarTesoreria:
    If Err.Number <> 0 Then
        b = False
        MenError = "Error al insertar en Tesoreria: " & Err.Description
    End If
    InsertarEnTesoreriaNewFVARPro = b
End Function



'############################################################################
'################ INSERTAR EN DIARIO EL ASIENTO DE COBRO DE RMT
'############################################################################

Private Function InsertarAsientoCobroPOZOS(cadMen As String, ByRef Rs As ADODB.Recordset, FecRec As Date, CtaContra As String) As Boolean

' la contabilizacion de las facturas internas de horto, se han de insertar en el diario no en el registro de iva de proveedor
Dim sql As String
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Concepto As String
Dim LetraSer As String
Dim Obs As String
Dim b As Boolean
'Dim CtaSocio As String

Dim Mc As Contadores
    On Error GoTo EInsertar
       
    Cad = ""
    Set Mc = New Contadores

    If Mc.ConseguirContador("0", (DBLet(Rs!fecfactu) <= CDate(vEmpresa.FechaFin)), True) = 0 Then
    
        sql = "select codmaccli from rsocios_seccion where codsecci = " & vParamAplic.SeccionPOZOS & " and codsocio = " & DBSet(Rs!Codsocio, "N")
        CtaSocio = DevuelveValor(sql)
        
        '[Monica]18/06/2014: antes poniamos la fecha de factura, ahora la fecha de hoy
        Obs = "Contabilización Cobro Rec.Manta " & Format(Now, "dd/mm/yyyy")
    
        'Insertar en la conta Cabecera Asiento
        cadMen = ""
        b = InsertarCabAsientoDia(1, Mc.Contador, CStr(Format(Rs!fecfactu, "dd/mm/yyyy")), Obs, cadMen)
        cadMen = "Insertando Cab. Asiento: " & cadMen

        If b Then
            cadMen = ""
            b = InsertarLinAsientoCobroPOZOS(Rs, cadMen, CtaSocio, CtaContra, Mc.Contador)
            cadMen = "Insertando Lin. Asiento Diario: " & cadMen
        
        End If
        
        If b Then
        
            ProcesoCorrecto = False
        
            frmActualizar2.Numasiento = Mc.Contador
            frmActualizar2.FechaAsiento = Rs!fecfactu
            frmActualizar2.numdiari = vEmpresa.NumDiarioInt
            frmActualizar2.OpcionActualizar = 1
            frmActualizar2.Show vbModal
            
            b = ProcesoCorrecto
        End If
            
        Set Mc = Nothing
        
        
    End If
    
EInsertar:
    If Err.Number <> 0 Or Not b Then
        InsertarAsientoCobroPOZOS = False
        cadMen = cadMen & Err.Description
    Else
        InsertarAsientoCobroPOZOS = b And ProcesoCorrecto
    End If
End Function


Private Function InsertarLinAsientoCobroPOZOS(ByRef Rs As ADODB.Recordset, cadErr As String, CtaSocio As String, CtaContra As String, Contador As Long) As Boolean
Dim sql As String
Dim numdocum As String
Dim Ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim I As Long
Dim b As Boolean
Dim Cad As String
Dim FeFact As Date
Dim cadMen As String

Dim LetraSer As String
Dim Concep As Integer
Dim Amplia As String

    On Error GoTo eInsertarLinAsientoCobroPOZOS

    InsertarLinAsientoCobroPOZOS = False
        
        
    I = 0
    ImporteD = 0
    ImporteH = 0
    
    LetraSer = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(Rs!CodTipom, "T"))
    
    numdocum = LetraSer & Format(Rs!numfactu, "0000000")
    
    Concep = 3
    
    Amplia = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", CStr(Concep), "N"))
    
    ampliaciond = Amplia & " " & LetraSer & "/" & DBLet(Rs!numfactu, "N")
    ampliacionh = Amplia & " " & LetraSer & "/" & DBLet(Rs!numfactu, "N")
    
    b = True
    
    I = I + 1
    
    FeFact = Rs!fecfactu
    
    Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Contador, "N") & ","
    Cad = Cad & DBSet(I, "N") & "," & DBSet(CtaSocio, "T") & "," & DBSet(numdocum, "T") & ","
    
    ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
    If DBLet(Rs!TotalFact, "N") > 0 Then
        ' importe al haber en positivo
        Cad = Cad & DBSet(Concep, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
        Cad = Cad & DBSet(Rs!TotalFact, "N") & "," & ValorNulo & "," & DBSet(CtaContra, "T") & "," & ValorNulo & ",0"
    
        ImporteH = ImporteH + CCur(DBLet(Rs!TotalFact, "N"))
        
    Else
        ' importe al debe en positivo cambiamos signo
        Cad = Cad & DBSet(Concep, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(DBLet(Rs!TotalFact, "N") * (-1), "N") & ","
        Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaContra, "T") & "," & ValorNulo & ",0"
    
        ImporteD = ImporteD + CCur(DBLet(Rs!TotalFact, "N") * (-1))
    
    End If
    
    Cad = "(" & Cad & ")"
    
    b = InsertarLinAsientoDia(Cad, cadMen)
    cadMen = "Insertando Lin. Asiento: " & I

    I = I + 1
            
    ' el Total es sobre la cuenta del cliente
    Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Contador, "N") & ","
    Cad = Cad & DBSet(I, "N") & ","
    Cad = Cad & DBSet(CtaContra, "T") & "," & DBSet(numdocum, "T") & ","
        
    ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
    If DBLet(Rs!TotalFact, "N") > 0 Then
        ' importe al debe en positivo
        Cad = Cad & DBSet(Concep, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(DBLet(Rs!TotalFact, "N"), "N") & ","
        Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(CtaSocio, "N") & "," & ValorNulo & ",0"
    Else
        ' importe al haber en positivo, cambiamos el signo
        Cad = Cad & DBSet(Concep, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
        Cad = Cad & DBSet(DBLet(Rs!TotalFact, "N") * (-1), "N") & "," & ValorNulo & "," & DBSet(CtaSocio, "N") & "," & ValorNulo & ",0"
    End If
    
    Cad = "(" & Cad & ")"
    
    b = InsertarLinAsientoDia(Cad, cadMen)
    cadMen = "Insertando Lin. Asiento: " & I

    InsertarLinAsientoCobroPOZOS = b
    Exit Function
    
eInsertarLinAsientoCobroPOZOS:
    cadErr = "Insertar Linea Asiento Cobro Pozos: " & Err.Description
    cadErr = cadErr & cadMen
    InsertarLinAsientoCobroPOZOS = False
End Function



Public Function ComprobarSociosSeccion(cadTabla As String, Seccion As Integer) As Boolean
'Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
'vamos a contabilizar
Dim sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECompFactu

    ComprobarSociosSeccion = False
    
    If cadTabla = "rrecibpozos" Then
        sql = "SELECT DISTINCT rrecibpozos.codsocio "
        sql = sql & " FROM (rrecibpozos LEFT JOIN rsocios_seccion ON rrecibpozos.codsocio=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N") & ") "
        sql = sql & " INNER JOIN tmpFactu ON rrecibpozos.codtipom=tmpFactu.codtipom AND rrecibpozos.numfactu=tmpFactu.numfactu AND rrecibpozos.fecfactu=tmpFactu.fecfactu "

        Set Rs = New ADODB.Recordset
        Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF And b
            Sql2 = "select * from rsocios_seccion where (codsocio= " & DBSet(Rs!Codsocio, "N") & " and rsocios_seccion.codsecci = " & DBSet(vParamAplic.SeccionPOZOS, "N") & ")"
            If RegistrosAListar(Sql2, cAgro) = 0 Then
                b = False
                
                Select Case cadTabla
                    Case "rrecibpozos"
                        sql = "Socio no existente en la sección de pozos: " & DBSet(Rs!Codsocio, "N") & vbCrLf
                End Select
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            sql = "Comprobando Socios en Sección ...       " & vbCrLf & vbCrLf & sql
            
            MsgBox sql, vbExclamation
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







'####################################################################################
'################## FUNCIONES PARA ACTUALIZAR UN ASIENTO DEL DIARIO EN EL HCO
'####################################################################################

'Private Function ActualizaElASiento(ByRef A_Donde As String) As Boolean
'
'
'
'    ActualizaElASiento = False
'
'    'Insertamos en cabeceras
'    A_Donde = "Insertando datos en historico cabeceras asiento"
'    If Not InsertarCabecera Then Exit Function
'
'    'Insertamos en lineas
'    A_Donde = "Insertando datos en historico lineas asiento"
'    If Not InsertarLineas Then Exit Function
'
'
'
'    'Modificar saldos
'    A_Donde = "Calculando Lineas y saldos "
'    If Not CalcularLineasYSaldos(False) Then Exit Function
'
'
'    'Borramos cabeceras y lineas del asiento
'    A_Donde = "Borrar cabeceras y lineas en asientos"
'    If Not BorrarASiento(False) Then Exit Function
'    ActualizaElASiento = True
'End Function
'
'
'Private Function InsertarCabecera() As Boolean
'On Error Resume Next
'
'    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari) SELECT numdiari,fechaent,numasien,obsdiari from cabapu where "
'    SQL = SQL & " numdiari =" & numdiari
'    SQL = SQL & " AND fechaent='" & Fecha & "'"
'    SQL = SQL & " AND numasien=" & Numasiento
'
'    conn.Execute SQL
'
'    If Err.Number <> 0 Then
'         'Hay error , almacenamos y salimos
'        InsertarCabecera = False
'    Else
'        InsertarCabecera = True
'    End If
'End Function
'
'Private Function InsertarLineas() As Boolean
'On Error Resume Next
'    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab,punteada)"
'    SQL = SQL & " SELECT numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab,punteada From linapu"
'    SQL = SQL & " WHERE numasien = " & Numasiento
'    SQL = SQL & " AND numdiari = " & numdiari
'    SQL = SQL & " AND fechaent='" & Fecha & "'"
'    conn.Execute SQL
'    If Err.Number <> 0 Then
'        'Hay error , almacenamos y salimos
'        InsertarLineas = False
'    Else
'        InsertarLineas = True
'    End If
'End Function
'
'
'Private Function CalcularLineasYSaldos(EsDesdeRecalcular As Boolean) As Boolean
'Dim Reparto As Boolean
'Dim T As String
'
'    Dim RL As Recordset
'    Set RL = New ADODB.Recordset
'
'
'    'Ahora
'    SQL = "SELECT timporteD AS SD, timporteH AS SH, codmacta"
'    SQL = SQL & "  FROM"
'    If EsDesdeRecalcular Then
'        SQL = SQL & " hlinapu"
'    Else
'        SQL = SQL & " linapu"
'    End If
'    'SQL = SQL & " GROUP BY codmacta, numdiari, fechaent, numasien"
'    SQL = SQL & " WHERE (((numdiari)= " & numdiari
'    SQL = SQL & ") AND ((fechaent)='" & Fecha & "'"
'    SQL = SQL & ") AND ((numasien)=" & Numasiento
'    SQL = SQL & "));"
'
'    Set RL = New ADODB.Recordset
'    RL.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    While Not RL.EOF
'        Cuenta = RL!Codmacta
'        If IsNull(RL!sD) Then
'            ImporteD = 0
'        Else
'            'ImporteD = RL!tImporteD
'            ImporteD = RL!sD
'        End If
'        If IsNull(RL!sH) Then
'            ImporteH = 0
'        Else
'            'ImporteH = RL!tImporteH
'            ImporteH = RL!sH
'        End If
'
'        If Not CalcularSaldos Then
'            RL.Close
'            Exit Function
'        End If
'
'        'Sig
'        RL.MoveNext
'    Wend
'    RL.Close
'    IncrementaProgres 3
'    If Not vParam.Autocoste Then
'        'NO tiene analitica
'        CalcularLineasYSaldos = True
'        Exit Function
'    End If
'
'
'    '------------------------------------------
'    '       ANALITICA     -> Modificado para 2 de Julio, para subcentros de reparto
'
'    If EsDesdeRecalcular Then
'        T = "h"
'    Else
'        T = ""
'    End If
'
'
'    SQL = "SELECT timporteD AS SD, timporteH AS SH, codmacta,"
'    SQL = SQL & " fechaent, numdiari, numasien, " & T & "linapu.codccost, idsubcos"
'    SQL = SQL & " FROM " & T & "linapu,cabccost WHERE cabccost.codccost=" & T & "linapu.codccost"
'    'SQL = SQL & " GROUP BY codmacta, fechaent, numdiari, numasien, codccost"
'    SQL = SQL & " AND numdiari=" & numdiari
'    SQL = SQL & " AND fechaent='" & Fecha & "'"
'    SQL = SQL & " AND numasien=" & Numasiento
'    SQL = SQL & " AND " & T & "linapu.codccost Is Not Null;"
'
'
'
'
'
'    RL.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    While Not RL.EOF
'        Cuenta = RL!Codmacta
'        CCost = RL!CodCCost
'        ImporteD = DBLet(RL!sD, "N")
'        ImporteH = DBLet(RL!sH, "N")
'        Reparto = (RL!idsubcos = 1)
'        If Not CalcularSaldosAnal Then
'            RL.Close
'            Exit Function
'        End If
'        If Reparto Then
'            If Not HacerReparto(True) Then
'                RL.Close
'                Exit Function
'            End If
'        End If
'        'Sig
'        RL.MoveNext
'    Wend
'    RL.Close
'    IncrementaProgres 2
'    CalcularLineasYSaldos = True
'End Function
'
'
'
'Private Function BorrarASiento(EnHistorico As Boolean) As Boolean
'    BorrarASiento = False
'
'    'Borramos las lineas
'    SQL = "Delete from "
'    If EnHistorico Then
'        SQL = SQL & "hlinapu"
'    Else
'        SQL = SQL & "linapu"
'    End If
'    SQL = SQL & " WHERE numasien = " & Numasiento
'    SQL = SQL & " AND numdiari = " & numdiari
'    SQL = SQL & " AND fechaent='" & Fecha & "'"
'    conn.Execute SQL
'
'
'    'La cabecera
'    SQL = "Delete from "
'    If EnHistorico Then
'        SQL = SQL & "hcabapu"
'    Else
'        SQL = SQL & "cabapu"
'    End If
'    SQL = SQL & " WHERE numdiari =" & numdiari
'    SQL = SQL & " AND fechaent='" & Fecha & "'"
'    SQL = SQL & " AND numasien=" & Numasiento
'    conn.Execute SQL
'
'    BorrarASiento = True
'End Function
'
'

'####################################################################################
'   HASTA AQUI  ---> FUNCIONES PARA ACTUALIZAR UN ASIENTO DEL DIARIO EN EL HCO
'####################################################################################

