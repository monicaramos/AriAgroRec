Attribute VB_Name = "ModConta"
Option Explicit

'=============================================================================
'   MODULO PARA ACCEDER A LOS DATOS DE LA CONTABILIDAD
'=============================================================================


'=============================================================================
'==========     CUENTAS
'=============================================================================
'LAURA
Public Function PonerNombreCuenta(ByRef Txt As TextBox, Modo As Byte, Optional clien As String) As String
'Obtener el nombre de una cuenta
Dim DevfrmCCtas As String
Dim Cad As String

' ### [Monica] 07/09/2006 añadida la linea siguiente condicion vParamAplic.NumeroConta = 0
' para que no saque nada si no hay contabilidad
    If Not vParamAplic Is Nothing Then
        If vParamAplic.NumeroConta = 0 Then
            PonerNombreCuenta = ""
            Exit Function
        End If
    End If
    If Txt.Text = "" Then
         PonerNombreCuenta = ""
         Exit Function
    End If
    DevfrmCCtas = Txt.Text
    If CuentaCorrectaUltimoNivel(DevfrmCCtas, Cad) Then
        ' ### [Monica] 07/09/2006
        If InStr(Cad, "No existe la cuenta") > 0 Then
            Txt.Text = DevfrmCCtas
'            If (Modo = 4) And clien <> "" Then 'si insertar antes estaba lo de abajo
            If (Modo = 3 Or Modo = 4) And clien <> "" Then 'si insertar o modificar
                Cad = Cad & "  ¿Desea crearla?"
                If MsgBox(Cad, vbYesNo) = vbYes Then
                    If InStr(1, Txt.Tag, "rsocio") > 0 Then
                        InsertarCuentaCble DevfrmCCtas, clien
                    ElseIf InStr(1, Txt.Tag, "sprove") > 0 Then
                        InsertarCuentaCble DevfrmCCtas, "", clien
                    End If
                    PonerNombreCuenta = clien
                End If
            Else
                MsgBox Cad, vbExclamation
            End If
        Else
            Txt.Text = DevfrmCCtas
            PonerNombreCuenta = Cad
        End If
    Else
        MsgBox Cad, vbExclamation
'        Txt.Text = ""
        PonerNombreCuenta = ""
'        PonerFoco Txt
    End If
    DevfrmCCtas = ""

End Function




'DAVID: Cuentas del la Contabilidad
Public Function CuentaCorrectaUltimoNivel(ByRef cuenta As String, ByRef devuelve As String) As Boolean
    'Comprueba si es numerica
    Dim Sql As String
    Dim OtroCampo As String
    
    CuentaCorrectaUltimoNivel = False
    If cuenta = "" Then
        devuelve = "Cuenta vacia"
        Exit Function
    End If

    If Not IsNumeric(cuenta) Then
        devuelve = "La cuenta debe de ser numérica: " & cuenta
        Exit Function
    End If

    'Rellenamos si procede
    cuenta = RellenaCodigoCuenta(cuenta)

    '==========
    If Not EsCuentaUltimoNivel(cuenta) Then
        devuelve = "No es cuenta de último nivel: " & cuenta
        Exit Function
    End If
    '==================

    OtroCampo = "apudirec"
    'BD 2: conexion a BD Conta
    Sql = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", cuenta, "T", OtroCampo)
    If Sql = "" Then
        devuelve = "No existe la cuenta : " & cuenta
        CuentaCorrectaUltimoNivel = True
        Exit Function
    End If

    'Llegados aqui, si que existe la cuenta
    If OtroCampo = "S" Then 'Si es apunte directo
        CuentaCorrectaUltimoNivel = True
        devuelve = Sql
    Else
        devuelve = "No es apunte directo: " & cuenta
    End If
End Function


'DAVID
Public Function RellenaCodigoCuenta(vCodigo As String) As String
'Rellena con ceros hasta poner una cuenta.
'Ejemplo: 43.1 --> 430000001
Dim i As Integer
Dim J As Integer
Dim cont As Integer
Dim Cad As String

    RellenaCodigoCuenta = vCodigo
    If Len(vCodigo) > vEmpresa.DigitosUltimoNivel Then Exit Function
    
    i = 0: cont = 0
    Do
        i = i + 1
        i = InStr(i, vCodigo, ".")
        If i > 0 Then
            If cont > 0 Then cont = 1000
            cont = cont + i
        End If
    Loop Until i = 0

    'Habia mas de un punto
    If cont > 1000 Or cont = 0 Then Exit Function

    'Cambiamos el punto por 0's  .-Utilizo la variable maximocaracteres, para no tener k definir mas
    i = Len(vCodigo) - 1 'el punto lo quito
    J = vEmpresa.DigitosUltimoNivel - i
    Cad = ""
    For i = 1 To J
        Cad = Cad & "0"
    Next i

    Cad = Mid(vCodigo, 1, cont - 1) & Cad
    Cad = Cad & Mid(vCodigo, cont + 1)
    RellenaCodigoCuenta = Cad
End Function

'DAVID
Public Function EsCuentaUltimoNivel(cuenta As String) As Boolean
    EsCuentaUltimoNivel = (Len(cuenta) = vEmpresa.DigitosUltimoNivel)
End Function

' ### [Monica] 07/09/2006
' copia de la gestion
Private Function InsertarCuentaCble(cuenta As String, cadSocio As String, Optional cadProve As String) As Boolean
Dim Sql As String
Dim vSocio As cSocio
Dim b As Boolean
Dim vIban As String

    On Error GoTo EInsCta
    
    If Not vParamAplic.ContabilidadNueva Then
        Sql = "INSERT INTO cuentas (codmacta,nommacta,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos,maidatos,obsdatos,pais, entidad, oficina, cc, cuentaba"
        '[Monica]22/11/2013: tema iban
        If vEmpresa.HayNorma19_34Nueva = 1 Then
            Sql = Sql & ", iban) "
        Else
            Sql = Sql & ") "
        End If
    Else
        Sql = "INSERT INTO cuentas (codmacta,nommacta,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos,maidatos,obsdatos,codpais"
        Sql = Sql & ", iban) "
    End If
    
    Sql = Sql & " VALUES (" & DBSet(cuenta, "T") & ","
    If cadSocio <> "" Then
        Set vSocio = New cSocio
        If vSocio.LeerDatos(cadSocio) Then                          ' antes cuenta
            Sql = Sql & DBSet(vSocio.Nombre, "T") & ",'S',1," & DBSet(vSocio.Nombre, "T") & "," & DBSet(vSocio.Direccion, "T") & ","
            Sql = Sql & DBSet(vSocio.CPostal, "T") & "," & DBSet(vSocio.Poblacion, "T") & "," & DBSet(vSocio.Provincia, "T") & "," & DBSet(vSocio.nif, "T") & "," & DBSet(vSocio.EMail, "T") & "," & ValorNulo
            If Not vParamAplic.ContabilidadNueva Then
                Sql = Sql & ",'ESPAÑA',"
                Sql = Sql & DBSet(vSocio.Banco, "T", "S") & "," & DBSet(vSocio.Sucursal, "T", "S") & "," & DBSet(vSocio.Digcontrol, "T", "S") & "," & DBSet(vSocio.CuentaBan, "T", "S")
                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    Sql = Sql & "," & DBSet(vSocio.Iban, "T", "S") & ")"
                Else
                    Sql = Sql & ")"
                End If
            Else
                vIban = MiFormat(vSocio.Iban, "") & MiFormat(vSocio.Banco, "0000") & MiFormat(vSocio.Sucursal, "0000") & MiFormat(vSocio.Digcontrol, "00") & MiFormat(vSocio.CuentaBan, "0000000000")
                
                Sql = Sql & ",'ES',"
                Sql = Sql & DBSet(vIban, "T") & ")"
            End If
            ConnConta.Execute Sql
            cadSocio = vSocio.Nombre
            b = True
        Else
            b = False
        End If
        Set vSocio = Nothing
    End If
    
    
EInsCta:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Description, "Insertando cuenta contable", Err.Description
    End If
    InsertarCuentaCble = b
End Function


'=============================================================================
'==========     CENTROS DE COSTE
'=============================================================================
'LAURA
Public Function PonerNombreCCoste(Empresa As String, ByRef Txt As TextBox) As String
'Obtener el nombre de un centro de coste
Dim codCCoste As String
Dim Cad As String

     If Txt.Text = "" Then
         PonerNombreCCoste = ""
         Exit Function
    End If
    codCCoste = Txt.Text
    If CCosteCorrecto(Empresa, codCCoste, Cad) Then
        Txt.Text = codCCoste
        PonerNombreCCoste = Cad
    Else
        MsgBox Cad, vbExclamation
'        Txt.Text = ""
        PonerNombreCCoste = ""
        PonerFoco Txt
    End If
'    codCCoste = ""
End Function

'LAURA
Public Function CCosteCorrecto(Empresa As String, ByRef Centro As String, ByRef devuelve As String) As Boolean
    Dim Sql As String
    
    CCosteCorrecto = False
 
    'BD 2: conexion a BD Conta
    If Not vParamAplic.ContabilidadNueva Then
        If Val(Empresa) <> Val(vEmpresa.codempre) Then
            Sql = DevuelveDesdeBDNew(3, "cabccost", "nomccost", "codccost", Centro, "T")
        Else
            Sql = DevuelveDesdeBDNew(cConta, "cabccost", "nomccost", "codccost", Centro, "T")
        End If
    Else
        If Val(Empresa) <> Val(vEmpresa.codempre) Then
            Sql = DevuelveDesdeBDNew(3, "ccoste", "nomccost", "codccost", Centro, "T")
        Else
            Sql = DevuelveDesdeBDNew(cConta, "ccoste", "nomccost", "codccost", Centro, "T")
        End If
    End If
        
    If Sql = "" Then
        devuelve = "No existe el Centro de coste : " & Centro
        Exit Function
    Else
        devuelve = Sql
        CCosteCorrecto = True
    End If
    
End Function




'=============================================================================
'==========     CONCEPTOS
'=============================================================================
'LAURA
Public Function PonerNombreConcepto(ByRef Txt As TextBox) As String
'Obtener el nombre de un concepto
Dim codConce As String
Dim Cad As String

     If Txt.Text = "" Then
         PonerNombreConcepto = ""
         Exit Function
    End If
    codConce = Txt.Text
    If ConceptoCorrecto(codConce, Cad) Then
        Txt.Text = Format(codConce, "000")
        PonerNombreConcepto = Cad
    Else
        MsgBox Cad, vbExclamation
        Txt.Text = ""
        PonerNombreConcepto = ""
        PonerFoco Txt
    End If
End Function


'LAURA
Public Function ConceptoCorrecto(ByRef Concep As String, ByRef devuelve As String) As Boolean
    Dim Sql As String
    
    ConceptoCorrecto = False
 
    'BD 2: conexion a BD Conta
    Sql = DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", Concep, "N")
    If Sql = "" Then
        devuelve = "No existe el concepto : " & Concep
        Exit Function
    Else
        devuelve = Sql
        ConceptoCorrecto = True
    End If
End Function

' ### [Monica] 27/09/2006
Public Function FacturaContabilizada(numserie As String, numfactu As String, Anofactu As String) As Boolean
Dim Sql As String
Dim NumAsi As Currency

    FacturaContabilizada = False
    Sql = ""
    If Not vParamAplic.ContabilidadNueva Then
        Sql = DevuelveDesdeBDNew(cConta, "cabfact", "numasien", "numserie", numserie, "T", , "codfaccl", numfactu, "N", "anofaccl", Anofactu, "N")
    Else
        Sql = DevuelveDesdeBDNew(cConta, "factcli", "numasien", "numserie", numserie, "T", , "numfactu", numfactu, "N", "anofactu", Anofactu, "N")
    End If
    
    If Sql = "" Then Exit Function
    
    NumAsi = DBLet(Sql, "N")
    
    If NumAsi <> 0 Then FacturaContabilizada = True

End Function

' ### [Monica] 27/09/2006
Public Function FacturaRemesada(numserie As String, numfactu As String, fecfactu As String) As Boolean
Dim Sql As String
Dim NumRem As Currency

    FacturaRemesada = False
    
    Sql = ""
    If Not vParamAplic.ContabilidadNueva Then
        Sql = DevuelveDesdeBDNew(cConta, "scobro", "codrem", "numserie", numserie, "T", , "codfaccl", numfactu, "N", "fecfaccl", fecfactu, "F")
    Else
        Sql = DevuelveDesdeBDNew(cConta, "cobros", "codrem", "numserie", numserie, "T", , "numfactu", numfactu, "N", "fecfactu", fecfactu, "F")
    End If
    
    If Sql = "" Then Exit Function
    
    NumRem = DBLet(Sql, "N")
    
    If NumRem <> 0 Then FacturaRemesada = True
    
End Function

' ### [Monica] 27/09/2006
Public Function FacturaCobrada(numserie As String, numfactu As String, fecfactu As String) As Boolean
Dim Sql As String
Dim ImpCob As Currency

    FacturaCobrada = False
    Sql = ""
    If Not vParamAplic.ContabilidadNueva Then
        Sql = DevuelveDesdeBDNew(cConta, "scobro", "impcobro", "numserie", numserie, "T", , "codfaccl", numfactu, "N", "fecfaccl", fecfactu, "F")
    Else
        Sql = DevuelveDesdeBDNew(cConta, "cobros", "impcobro", "numserie", numserie, "T", , "numfactu", numfactu, "N", "fecfactu", fecfactu, "F")
    End If
    If Sql = "" Then Exit Function
    ImpCob = DBLet(Sql, "N")
    
    If ImpCob <> 0 Then FacturaCobrada = True
    
End Function

' ### [Monica] 27/09/2006
Public Function ModificaClienteFacturaContabilidad(letraser As String, numfactu As String, fecfactu As String, CtaConta As String, Tipo As Byte) As Boolean
Dim Sql As String
Dim Anyo As Currency

    On Error GoTo eModificaClienteFacturaContabilidad

    ModificaClienteFacturaContabilidad = False

    Anyo = Year(CDate(fecfactu))
    
    If Not vParamAplic.ContabilidadNueva Then
        If Tipo = 0 Then
            Sql = "update cabfact set codmacta = " & DBSet(CtaConta, "T") & " where numserie = " & DBSet(letraser, "T") & " and " & _
                      "codfaccl = " & DBSet(numfactu, "N") & " and anofaccl = " & DBSet(Anyo, "N")
            ConnConta.Execute Sql
        End If
        
        Sql = "update scobro set codmacta = " & DBSet(CtaConta, "T") & " where numserie = " & DBSet(letraser, "T") & " and " & _
                     "codfaccl = " & DBSet(numfactu, "N") & " and fecfaccl = " & DBSet(fecfactu, "F")
                  
        ConnConta.Execute Sql
    Else
        If Tipo = 0 Then
            Sql = "update factcli set codmacta = " & DBSet(CtaConta, "T") & " where numserie = " & DBSet(letraser, "T") & " and " & _
                      "numfactu = " & DBSet(numfactu, "N") & " and anofactu = " & DBSet(Anyo, "N")
            ConnConta.Execute Sql
        End If
        
        Sql = "update cobros set codmacta = " & DBSet(CtaConta, "T") & " where numserie = " & DBSet(letraser, "T") & " and " & _
                     "numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(fecfactu, "F")
                  
        ConnConta.Execute Sql
    End If
              
    ModificaClienteFacturaContabilidad = True
    
eModificaClienteFacturaContabilidad:
    If Err.Number <> 0 Then
        MsgBox "Error en ModificaClienteFacturaContabilidad: " & Err.Description, vbExclamation
    End If

End Function

' ### [Monica] 27/09/2006
Public Sub ModificaFormaPagoTesoreria(letraser As String, numfactu As String, fecfactu As String, Forpa As String, forpaant As String)
Dim Sql As String
Dim Sql1 As String
Dim TipForpa As String
Dim TipForpaAnt As String
Dim cadWHERE As String

    
    If Not vParamAplic.ContabilidadNueva Then
        cadWHERE = " numserie = " & DBSet(letraser, "T") & " and " & _
                  "codfaccl = " & numfactu & " and fecfaccl = " & DBSet(fecfactu, "F")
        
        Sql = "update scobro set codforpa = " & Forpa & " where " & cadWHERE
    Else
        cadWHERE = " numserie = " & DBSet(letraser, "T") & " and " & _
                  "numfactu = " & numfactu & " and fecfactu = " & DBSet(fecfactu, "F")
        
        Sql = "update cobros set codforpa = " & Forpa & " where " & cadWHERE
    
    End If
    ConnConta.Execute Sql

End Sub

'' ### [Monica] 29/09/2006
'Public Function ModificaImportesFacturaContabilidad(letraser As String, numfactu As String, fecfactu As String, Importe As String, Forpa As String, vtabla As String) As Boolean
'Dim SQL As String
'Dim vWhere As String
'Dim b As Boolean
'Dim CadValues As String
'Dim vSocio As CSocio
'Dim RS As ADODB.Recordset
'Dim TipForpa As String
'
'    On Error GoTo eModificaImportesFacturaContabilidad
'
'    b = False
'
'    vWhere = "numserie = " & DBSet(letraser, "T") & " and codfaccl = " & _
'              numfactu & " and anofaccl = " & Format(Year(fecfactu), "0000")
'
'
'    SQL = "select codsocio from " & vtabla & " where letraser = " & DBSet(letraser, "T") & " and numfactu = " & _
'           numfactu & " and fecfactu = " & DBSet(fecfactu, "F")
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    If Not RS.EOF Then RS.MoveFirst
'
'    Set vSocio = New CSocio
'    If vSocio.LeerDatos(RS.Fields(0).Value) Then
'    '********************************+estoy aqui
'
'        If vtabla = "schfac" Then
'            SQL = "delete from linfact where " & vWhere
'            ConnConta.Execute SQL
'
'            SQL = "delete from cabfact where " & vWhere
'            ConnConta.Execute SQL
'
'            SQL = "schfac.letraser = " & DBSet(letraser, "T") & " and numfactu = " & numfactu
'            SQL = SQL & " and fecfactu = " & DBSet(fecfactu, "F")
'
'
'            b = CrearTMPErrFact("schfac")
'            If b Then b = PasarFactura2(SQL, vSocio)
'        Else
'            b = CrearTMPErrFact("schfacr")
'        End If
'
'        ' 09/02/2007
'        TipForpa = DevuelveDesdeBDNew(cAgro, "sforpa", "tipforpa", "codforpa", Forpa, "N")
'        If TipForpa <> "0" And b Then
'            b = ModificaCobroTesoreria(letraser, numfactu, fecfactu, vSocio, vtabla)
'        End If
'    End If
'
'    ModificaImportesFacturaContabilidad = b
'
'eModificaImportesFacturaContabilidad:
'    If Err.Number <> 0 Then
'        MsgBox "Error en ModificaImportesFacturaContabilidad: " & Err.Description, vbExclamation
'    End If
'End Function

'Public Function ModificaCobroTesoreria(letraser As String, numfactu As String, fecfactu As String, vSocio As CSocio, vtabla As String) As Boolean
'Dim SQL As String
'Dim RS As ADODB.Recordset
'Dim cadwhere As String
'Dim BanPr As String
'Dim Mens As String
'Dim b As Boolean
'
'    On Error GoTo eModificaCobroTesoreria
'
'    ModificaCobroTesoreria = False
'    b = True
'
'    ' antes de borrar he de obtener la fecha de vencimiento y el codmacta para sacar el banco propio que le pasaré
'    ' a la rutina de InsertarEnTesoreria
'
'    SQL = "select fecvenci, ctabanc1 from scobro where numserie = " & DBSet(letraser, "T") & " and codfaccl = " & DBSet(numfactu, "N")
'    SQL = SQL & " and fecfaccl = " & DBSet(fecfactu, "F") & " and numorden = 1"
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
'    If Not RS.EOF Then
'        RS.MoveFirst
'
'        cadwhere = vtabla & ".letraser =" & DBSet(letraser, "T") & " and numfactu=" & DBLet(numfactu, "N")
'        cadwhere = cadwhere & " and fecfactu=" & DBSet(fecfactu, "F")
'
'        BanPr = ""
'        BanPr = DevuelveDesdeBDNew(cAgro, "sbanco", "codbanpr", "codmacta", RS.Fields(1).Value, "T")
'
'        SQL = "delete from scobro where "
'        SQL = SQL & " numserie = " & DBSet(letraser, "T") & " and codfaccl = " & numfactu
'        SQL = SQL & " and fecfaccl = " & DBSet(fecfactu, "F")
'
'        ConnConta.Execute SQL
'
'        ' hemos de crear el cobro nuevamente
'        Mens = "Insertando en Tesoreria "
'        b = InsertarEnTesoreria(cadwhere, CStr(RS.Fields(0).Value), BanPr, Mens, vSocio, vtabla)
'    End If
'
'    ModificaCobroTesoreria = b
'
'eModificaCobroTesoreria:
'    If Err.Number <> 0 Then
'        MsgBox "Error en ModificaCobroTesoreria " & Err.Description, vbExclamation
'    End If
'End Function


Public Function CalcularIva(Importe As String, articulo As String) As Currency
'devuelve el iva del Importe
'Ej el 16% de 120 = 19.2
Dim vImp As Currency
Dim vIva As Currency
Dim vArt As Currency
Dim CodIva As String

Dim IvaArt As Integer
Dim iva As String
Dim impiva As Currency
On Error Resume Next

    Importe = ComprobarCero(Importe)
    articulo = ComprobarCero(articulo)
    
    CodIva = DevuelveDesdeBD("codigiva", "sartic", "codartic", articulo, "N")
    iva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CodIva, "N")
    
    vImp = CCur(Importe)
    vIva = CCur(iva)
    
    impiva = ((vImp * vIva) / 100)
    impiva = Round(impiva, 2)
    
    CalcularIva = CStr(impiva)
    If Err.Number <> 0 Then Err.Clear

End Function


Public Function CalcularBase(Importe As String, articulo As String) As Currency
'devuelve la base del Importe
'Ej el 16% de 120 = 120-19.2 = 100.8
Dim vImp As Currency
Dim vIva As Currency
Dim vArt As Currency
Dim CodIva As String

Dim IvaArt As Integer
Dim iva As String
Dim impiva As Currency
On Error Resume Next

    Importe = ComprobarCero(Importe)
    articulo = ComprobarCero(articulo)
    
    CodIva = DevuelveDesdeBD("codigiva", "sartic", "codartic", articulo, "N")
    iva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CodIva, "N")
    
    vImp = CCur(Importe)
    vIva = CCur(iva)
    
    impiva = Round2(Importe / (1 + (vIva / 100)), 2)
    
    CalcularBase = CStr(impiva)
    If Err.Number <> 0 Then Err.Clear

End Function


'MONICA: Cuentas del la Contabilidad
Public Function NombreCuentaCorrecta(ByRef cuenta As String) As String
    'Comprueba si es numerica
    Dim Sql As String
    Dim OtroCampo As String
    
' ### [Monica] 27/10/2006 añadida la linea siguiente condicion vParamAplic.NumeroConta = 0
' para que no saque nada si no hay contabilidad
    If cuenta = "" Or vParamAplic.NumeroConta = 0 Then
         NombreCuentaCorrecta = ""
         Exit Function
    End If
    
    NombreCuentaCorrecta = ""
    If cuenta = "" Then
        MsgBox "Cuenta vacia", vbExclamation
        Exit Function
    End If

    If Not IsNumeric(cuenta) Then
        MsgBox "La cuenta debe de ser numérica: " & cuenta, vbExclamation
        Exit Function
    End If

    'BD 2: conexion a BD Conta
    Sql = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", cuenta, "T")
    If Sql = "" Then
        MsgBox "No existe la cuenta : " & cuenta, vbExclamation
    Else
        NombreCuentaCorrecta = Sql
    End If

End Function

