Attribute VB_Name = "Norma34"
Option Explicit

Dim AuxD As String
Private NumeroTransferencia As Integer

'----------------------------------------------------------------------
'  Copia fichero generado bajo
Public Sub CopiarFicheroNorma43(Destino As String)

    
    'If Not CopiarEnDisquette(True, 3) Then
        AuxD = Destino
        CopiarEnDisquette False, 0  'A disco
    
        
End Sub

Private Function CopiarEnDisquette(A_disquetera As Boolean, Intentos As Byte) As Boolean
Dim I As Integer
Dim cad As String

On Error Resume Next

    CopiarEnDisquette = False
    
    If A_disquetera Then
        For I = 1 To Intentos
            cad = "Introduzca un disco vacio. (" & I & ")"
            MsgBox cad, vbInformation
            FileCopy App.Path & "\norma34.txt", "a:\norma34.txt"
            If Err.Number <> 0 Then
                MuestraError Err.Number, "Copiar En Disquette"
            Else
                CopiarEnDisquette = True
                Exit For
            End If
        Next I
    Else
        If AuxD = "" Then
            cad = Format(Now, "ddmmyyhhnn")
            cad = App.Path & "\" & cad & ".txt"
        Else
            cad = AuxD
        End If
        FileCopy App.Path & "\norma34.txt", cad
        If Err.Number <> 0 Then
            MsgBox "Error creando copia fichero. Consulte soporte técnico." & vbCrLf & Err.Description, vbCritical
        Else
            MsgBox "El fichero esta guardado como: " & cad, vbInformation
        End If
            
    End If
End Function



'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'Cuenta propia tendra empipados entidad|sucursal|cc|cuenta|
Public Function GeneraFicheroNorma34(CIF As String, Fecha As Date, CuentaPropia As String, ConceptoTransferencia As String, vNumeroTransferencia As Integer, ByRef ConceptoTr As String, Pagos As Boolean) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Im As Currency
Dim Rs As ADODB.Recordset
Dim Aux As String
Dim cad As String


    On Error GoTo EGen
    GeneraFicheroNorma34 = False
    
    NumeroTransferencia = vNumeroTransferencia
    
    
    'Cargamos la cuenta
    cad = "Select * from ctabancaria where codmacta='" & CuentaPropia & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = Right("    " & CIF, 10)
    If Rs.EOF Then
        cad = ""
    Else
        If IsNull(Rs!entidad) Then
            cad = ""
        Else
            cad = Format(Rs!entidad, "0000") & "|" & Format(DBLet(Rs!oficina, "T"), "0000") & "|" & DBLet(Rs!Control, "T") & "|" & Format(DBLet(Rs!CtaBanco, "T"), "0000000000") & "|"
            CuentaPropia = cad
        End If
        
        'Identificador norma bancaria
        If Not IsNull(Rs!idnorma34) Then Aux = Rs!idnorma34
    End If
    Rs.Close
    Set Rs = Nothing
    If cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\norma34.txt" For Output As #NFich
    
    
    
    'Codigo ordenante
    '---------------------------------------------------
    'Si el banco tiene puesto si ID de norma34 entonces
    'la pongo aquin. Lo he cargado antes sobre la variable AUX
    CodigoOrdenante = Left(Aux & "          ", 10)  'CIF EMPRESA
    
    
    'CABECERA
    Cabecera1 NFich, CodigoOrdenante, Fecha, CuentaPropia, cad
    Cabecera2 NFich, CodigoOrdenante, cad
    Cabecera3 NFich, CodigoOrdenante, cad
    Cabecera4 NFich, CodigoOrdenante, cad
    
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma34
    Set Rs = New ADODB.Recordset
    If Pagos Then
        Aux = "Select spagop.*,nommacta,dirdatos,codposta,dirdatos,despobla from spagop,cuentas"
        Aux = Aux & " where codmacta=ctaprove and transfer =" & NumeroTransferencia
    Else
        'ABONOS
         '
        Aux = "Select scobro.codbanco as entidad,scobro.codsucur as oficina,scobro.cuentaba,scobro.digcontr as CC"
        Aux = Aux & ",nommacta,dirdatos,codposta,dirdatos,despobla,impvenci,scobro.codmacta from scobro,cuentas"
        Aux = Aux & " where cuentas.codmacta=scobro.codmacta and transfer =" & NumeroTransferencia
    End If
    Rs.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If Rs.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not Rs.EOF
            If Pagos Then
                Im = DBLet(Rs!imppagad, "N")
                Im = Rs!impefect - Im
                Aux = RellenaAceros(Rs!CtaProve, False, 12)
            
            Else
                Im = Abs(Rs!ImpVenci)
                Aux = RellenaAceros(Rs!Codmacta, False, 12)
            End If
            
            'Cad = "06"
            'Cad = Cad & "56"
            'Cad = Cad & " "
            'Aux = "06" & "56" & " " & CodigoOrdenante & Aux  'Ordenante y socio juntos
        
            Aux = "06" & "56" & CodigoOrdenante & Aux   'Ordenante y socio juntos
        
'            Linea1 NFich, Aux, Rs, Im, cad, ConceptoTransferencia
            Linea2 NFich, Aux, Rs, cad
            Linea3 NFich, Aux, Rs, cad
            Linea4 NFich, Aux, Rs, cad
            Linea5 NFich, Aux, Rs, cad
            Linea6 NFich, Aux, Rs, cad, ConceptoTr, Pagos
            If Pagos Then Linea7 NFich, Aux, Rs, cad
        
        
        
        
            Importe = Importe + Im
            Regs = Regs + 1
            Rs.MoveNext
        Wend
        'Imprimimos totales
        Totales NFich, CodigoOrdenante, Importe, Regs, cad, Pagos
    End If
    Rs.Close
    Set Rs = Nothing
    Close (NFich)
    If Regs > 0 Then GeneraFicheroNorma34 = True
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description

End Function

'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'Cuenta propia tendra empipados entidad|sucursal|cc|cuenta|
Public Function GeneraFicheroNorma34New(CIF As String, Fecha As Date, CuentaPropia As String, ConceptoTransferencia As String, vNumeroTransferencia As Integer, ByRef ConceptoTr As String, CodigoOrden As String, ConcepTransf As Byte) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Im As Currency
Dim Rs As ADODB.Recordset
Dim Aux As String
Dim cad As String
Dim Pagos As Boolean
Dim Concepto As Byte

    On Error GoTo EGen
    GeneraFicheroNorma34New = False
    
    NumeroTransferencia = vNumeroTransferencia
    
    If CodigoOrden = "" Then
'        Aux = Right("    " & CIF, 10)
        Aux = RellenaABlancos(CIF, True, 10)
    Else
        Aux = CodigoOrden
    End If

    NFich = FreeFile
    Open App.Path & "\norma34.txt" For Output As #NFich
    
    
    'Codigo ordenante
    '---------------------------------------------------
    'Si el banco tiene puesto si ID de norma34 entonces
    'la pongo aquin. Lo he cargado antes sobre la variable AUX
    CodigoOrdenante = Aux 'Left(Aux & "          ", 10)  'CIF EMPRESA
    
    
    'CABECERA
    Cabecera1 NFich, CodigoOrdenante, Fecha, CuentaPropia, cad
    Cabecera2 NFich, CodigoOrdenante, cad
    Cabecera3 NFich, CodigoOrdenante, cad
    Cabecera4 NFich, CodigoOrdenante, cad
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma34
    Set Rs = New ADODB.Recordset
    
    Aux = "select tmpimpor.*, straba.codbanco as entidad, straba.codsucur as oficina, straba.digcontr as CC, straba.cuentaba as cuentaba, "
    Aux = Aux & " straba.nomtraba as nommacta, straba.domtraba as dirdatos, straba.codpobla as codposta, straba.pobtraba as despobla, straba.niftraba as niftraba "
    Aux = Aux & " from tmpimpor, straba where tmpimpor.codtraba = straba.codtraba "
    
    Rs.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If Rs.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not Rs.EOF
            Im = DBLet(Rs!Importe, "N")
'--monica:20/08/08 sustituida por la siguiente
'            Aux = RellenaAceros("0", False, 12) 'Rs!Codmacta, False, 12)
'++monica:20/08/08
            Aux = RellenaABlancos(DBLet(Rs!niftraba, "T"), True, 12)
'            Aux = Mid(Left(DBLet(Rs!niftraba, "T"), 12), 1, 12)
            'Cad = "06"
            'Cad = Cad & "56"
            'Cad = Cad & " "
            'Aux = "06" & "56" & " " & CodigoOrdenante & Aux  'Ordenante y socio juntos

            Aux = "06" & "56" & CodigoOrdenante & Aux   'Ordenante y socio juntos
        
            Select Case ConcepTransf
                Case 0
                    Concepto = 1
                    ConceptoTransferencia = "Nómina"
                Case 1
                    Concepto = 8
                    ConceptoTransferencia = "Pensión"
                Case 2
                    Concepto = 9
                    ConceptoTransferencia = "Otros Conceptos"
            End Select
        
        
            Linea1 NFich, Aux, Rs, Im, cad, Concepto, ConceptoTransferencia
            Linea2 NFich, Aux, Rs, cad
            Linea3 NFich, Aux, Rs, cad
            Linea4 NFich, Aux, Rs, cad
            Linea5 NFich, Aux, Rs, cad
            Linea6 NFich, Aux, Rs, cad, ConceptoTransferencia, Pagos
            If Pagos Then Linea7 NFich, Aux, Rs, cad
        
            Importe = Importe + Im
            Regs = Regs + 1
            Rs.MoveNext
        Wend
        'Imprimimos totales
        Totales NFich, CodigoOrdenante, Importe, Regs, cad, Pagos
    End If
    Rs.Close
    Set Rs = Nothing
    Close (NFich)
    If Regs > 0 Then GeneraFicheroNorma34New = True
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description

End Function




Private Function RellenaABlancos(cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim cad As String
    
    cad = Space(longitud)
    If PorLaDerecha Then
        cad = cadena & cad
        RellenaABlancos = Left(cad, longitud)
    Else
        cad = cad & cadena
        RellenaABlancos = Right(cad, longitud)
    End If
    
End Function



Private Function RellenaAceros(cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim cad As String
    
    cad = Mid("00000000000000000000", 1, longitud)
    If PorLaDerecha Then
        cad = cadena & cad
        RellenaAceros = Left(cad, longitud)
    Else
        cad = cad & cadena
        RellenaAceros = Right(cad, longitud)
    End If
    
End Function



'Private Sub Cabecera1(NF As Integer,ByRef CodOrde As String)
'Dim Cad As String
'
'End Sub

Private Sub Cabecera1(NF As Integer, ByRef CodOrde As String, Fecha As Date, cta As String, ByRef cad As String)

    cad = "03"
    cad = cad & "56"
    'cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "001"
    cad = cad & Format(Now, "ddmmyy")
    cad = cad & Format(Fecha, "ddmmyy")
    'Cuenta bancaria
    cad = cad & RecuperaValor(cta, 1)
    cad = cad & RecuperaValor(cta, 2)
    cad = cad & RecuperaValor(cta, 4)
    cad = cad & "0"  'Sin relacion
    cad = cad & "   " & RecuperaValor(cta, 3)  'Digito de control bancario
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub



Private Sub Cabecera2(NF As Integer, ByRef CodOrde As String, ByRef cad As String)
    cad = "03"
    cad = cad & "56"
    'cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "002"
    
    cad = cad & RellenaABlancos(vParam.NombreEmpresa, True, 30)   'Nombre empresa
  
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Cabecera3(NF As Integer, ByRef CodOrde As String, ByRef cad As String)
    cad = "03"
    cad = cad & "56"
    'cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "003"
    
    
'    AuxD = DevuelveDesdeBD("direccion", "empresa2", "codigo", 1, "N")
    cad = cad & RellenaABlancos(vParam.DomicilioEmpresa, True, 30) 'AuxD, True, 30)   'Nombre empresa
    cad = cad & RellenaABlancos("", True, 30)   'Nombre empresa
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub



Private Sub Cabecera4(NF As Integer, ByRef CodOrde As String, ByRef cad As String)

    cad = "03"
    cad = cad & "56"
    'cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "004"
    
'    AuxD = DevuelveDesdeBD("codpos", "empresa2", "codigo", 1, "N")
    cad = cad & RellenaABlancos(vParam.CPostal, False, 5) '   AuxD, False, 5)
    cad = cad & " "
'    AuxD = DevuelveDesdeBD("provincia", "empresa2", "codigo", 1, "N")
    cad = cad & RellenaABlancos(vParam.Provincia, True, 30) 'AuxD, True, 30)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub



'ConceptoTransferencia
'1.- Abono nomina
'9.- Transferencia ordinaria
Private Sub Linea1(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef importe1 As Currency, ByRef cad As String, vconcepto As Byte, vConceptoTransferencia As String)


   
    '
    cad = CodOrde   'llevara tb la ID del socio
    cad = cad & "010"
    cad = cad & RellenaAceros(CStr(Round(importe1, 2) * 100), False, 12)
    
    cad = cad & RellenaAceros(CStr(DBLet(RS1!entidad, "N")), False, 4)    'Entidad
    cad = cad & RellenaAceros(CStr(DBLet(RS1!oficina, "N")), False, 4)  'Sucur
    cad = cad & RellenaAceros(CStr(DBLet(RS1!CuentaBa, "T")), False, 10) 'Cta
    cad = cad & "1" & Format(vconcepto, "0") '& vConceptoTransferencia
    cad = cad & "  "
    cad = cad & RellenaAceros(CStr(DBLet(RS1!CC, "T")), False, 2) 'CC
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea2(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "011"
    cad = cad & RellenaABlancos(DBLet(RS1!Nommacta, "T"), False, 36)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea3(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "012"
    cad = cad & RellenaABlancos(DBLet(RS1!dirdatos, "T"), False, 36)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea4(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "013"
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea5(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "014"
    cad = cad & RellenaABlancos(DBLet(RS1!codposta, "T"), False, 5) & " "
    cad = cad & RellenaABlancos(DBLet(RS1!desPobla, "T"), False, 30)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea6(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String, ByRef ConceptoT As String, Pagos As Boolean)
Dim Aux As String

    Aux = ConceptoT
    If Pagos Then
        'Tiene dos campos para las descripcion. Si no tiene nada pondre la descripcion de la transferencia
        Aux = Trim(DBLet(RS1!Text1csb, "T"))
        If Aux = "" Then Aux = ConceptoT
    End If

    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "016"
    cad = cad & RellenaABlancos(Aux, False, 35)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea7(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)


    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "017"
    cad = cad & RellenaABlancos(DBLet(RS1!Text2csb, "T"), False, 35)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub




Private Sub Totales(NF As Integer, ByRef CodOrde As String, Total As Currency, Registros As Integer, ByRef cad As String, Pagos As Boolean)
    cad = "08" & "56"
    cad = cad & CodOrde    'llevara tb la ID del socio
    cad = cad & Space(15)
    cad = cad & RellenaAceros(CStr(Int(Round(Total * 100, 2))), False, 12)
    cad = cad & RellenaAceros(CStr(Registros), False, 8)
    If Pagos Then
        cad = cad & RellenaAceros(CStr((Registros * 7) + 4 + 1), False, 10)
    Else
        cad = cad & RellenaAceros(CStr((Registros * 6) + 4 + 1), False, 10)
    End If
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub



'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'
'
'
'
'            SSSSSS         EEEEEEEE             PPPPPPP                 A
'           SS              EE                   PP     P               A A
'            SS             EE                   PP     P              A   A
'              SSS          EEEEEEEE             PPPPPPP              AAAAAAA
'                SS         EE                   PP                  A       A
'               SS          EE                   PP                 A         A
'           SSSSS           EEEEEEEE             PP                A           A
'
'
'
'
'
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************


Public Function GeneraFicheroNorma34SEPA(CIF As String, Fecha As Date, CuentaPropia2 As String, cadSQL As String, DescripcionTrans As String, Tipo As Byte, CodigoOrden34 As String) As Boolean
Dim Regs As Integer
Dim Importe As Currency
Dim Im As Currency
Dim cad As String
Dim Aux As String
Dim NF As Integer
Dim Sql As String
Dim Bic As String




    On Error GoTo EGen2
    GeneraFicheroNorma34SEPA = False
    

    
    Set miRsAux = New ADODB.Recordset
    
    'Cargamos la cuenta
    
    
        
    cad = RecuperaValor(CuentaPropia2, 5) & RecuperaValor(CuentaPropia2, 1) & RecuperaValor(CuentaPropia2, 2) & RecuperaValor(CuentaPropia2, 3) & RecuperaValor(CuentaPropia2, 4)
        
            
    If Len(cad) <> 24 Then
        MsgBox "Error leyendo datos para: " & CuentaPropia2, vbExclamation
        Exit Function
    End If
    CuentaPropia2 = cad
    NF = FreeFile
    Open App.Path & "\norma34.txt" For Output As #NF
    
    
    
    'SEPA
    '1.- Cabecera ordenante
    '------------------------------------------------------------------------
    cad = "01" & "ORD" & "34145" & "001" & CIF
        
    'sufijo (Tenemos el OEM, que se utiliza para las otras normas antiguas
    '[Monica]29/01/2014: antes "000"
    cad = cad & Right("000" & CodigoOrden34, 3)
    cad = cad & Format(Now, "yyyymmdd")
    cad = cad & Format(Fecha, "yyyymmdd")
    cad = cad & "A" 'IBAN
     
    'EL IBAN propiamente
    cad = cad & FrmtStr(CuentaPropia2, 34)
    '[Monica]24/02/2014: hacemos un cargo único antes era un 1
    cad = cad & "0" 'Cargo por cada operacion
    'Nombre
    miRsAux.Open "Select * from empresas", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = cad & FrmtStr(miRsAux!nomempre, 70)

        'Direccion   nomempre domempre codpobla pobempre proempre
        cad = cad & FrmtStr(Trim(miRsAux!domempre), 50)
        cad = cad & FrmtStr(Trim(DBLet(miRsAux!CodPobla, "T") & " " & miRsAux!pobempre), 50)
        cad = cad & FrmtStr(DBLet(miRsAux!proempre, "T"), 40)
    
    miRsAux.Close
    
    'Pais y libre
    cad = cad & "ES" & FrmtStr("", 311)
    Print #NF, cad
  
  
  
    '2.- Registro cabecera TRANSFERENCIA
    '------------------------------------------------------------------------
    cad = "02" & "SCT" & "34145" & CIF
        
    'sufijo (Tenemos el OEM, que se utiliza para las otras normas
    '[Monica]29/01/2014: antes "000"
    cad = cad & Right("000" & CodigoOrden34, 3)
    cad = cad & FrmtStr("", 578)
    Print #NF, cad
    
    '[Monica]22/11/2013: añadido
    If cadSQL = "" Then
        Sql = "select tmpimpor.codtraba, tmpimpor.importe, straba.codbanco as entidad,straba.codsucur "
        Sql = Sql & " as oficina,straba.digcontr as CC,straba.cuentaba as cuentaba ,straba.nomtraba as nommacta, "
        Sql = Sql & " straba.domtraba as dirdatos, straba.codpobla as codposta,straba.pobtraba as despobla, straba.protraba as desprovi, "
        Sql = Sql & " straba.niftraba as refbenef,straba.iban"
        Sql = Sql & " from tmpimpor, straba"
        Sql = Sql & " where tmpimpor.codtraba = straba.codtraba "
        
        cadSQL = Sql
    End If

    cad = cadSQL
    

    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Regs = 0
    Importe = 0
    If miRsAux.EOF Then
        'No hayningun registro

    Else
        While Not miRsAux.EOF
            
            
            Im = miRsAux!Importe
            Aux = miRsAux!refbenef
            
            Aux = FrmtStr(Aux, 10)
            Importe = Importe + Im
            Regs = Regs + 1
            
            'Campo 1,2,3
            cad = "03" & "SCT" & "34145" & "002"
            
            'Campo 5 . Referencia del ordenante
            
            
            Aux = miRsAux!refbenef & " " & Format(Fecha, "dd/mm/yyyy")
            cad = cad & FrmtStr(Aux, 35)
            
            'Campo 6
            cad = cad & "A"
            
            'IBAN
            cad = cad & FrmtStr(IBAN_Destino(), 34)
            
            'Campo8 Importe
            cad = cad & Format(Im * 100, String(11, "0")) ' Importe
            
            'Campo9
            cad = cad & "3" 'gastos compartidos
            'Campo 10
            Bic = DevuelveDesdeBDNew(cConta, "sbic", "bic", "entidad", miRsAux!entidad, "N")
            cad = cad & FrmtStr(Bic, 11) 'FrmtStr("", 11)  'BIC

            'nommacta,dirdatos,codposta,dirdatos,despobla,impvenci,scobro.codmacta
            'Datos Basicos del beneficiario
            cad = cad & DatosBasicosDelDeudor
            
            
                '`text33csb` `text41csb`
            'Aux = DBLet(miRsAux!text33csb, "T") & DBLet(miRsAux!text41csb, "T")
            Aux = DescripcionTrans ' he quitado esto de David & "   Cod: " & Format(miRsAux!CodTraba, "0000")
            cad = cad & FrmtStr(Aux, 140)
            
            'Campo17
            cad = cad & FrmtStr("", 35)  'Reservado
            
            'Campo18 y campo19
'            cad = cad & "TRAD"
'            cad = cad & "TRAD"
            Select Case Tipo
                Case 0:
                    cad = cad & "SALA"
                    cad = cad & "SALA"
                Case 1:
                    cad = cad & "PENS"
                    cad = cad & "PENS"
                Case 2:
                    cad = cad & "TRAD"
                    cad = cad & "TRAD"
            End Select
            
            
            cad = cad & FrmtStr("", 99)  'libre
            
            Print #NF, cad
            
            miRsAux.MoveNext
        Wend
        
    
        'TOTALES
        '----------------------------------
        'Total trasnferencia SEPA
        'Campo 1,2
        cad = "04" & "SCT"
        
        'Campo3 Importe total
        cad = cad & Format(Importe * 100, String(17, "0")) ' Importe
        cad = cad & Format(Regs, String(8, "0")) ' Importe
        'Total registros son
        'Reg(numreo de adeudos + 1 reg01 + un reg02 + reg04
        cad = cad & Format(Regs + 2, String(10, "0")) ' Importe
        cad = cad & FrmtStr("", 560)  'libre
        Print #NF, cad
        
        'Total general
        cad = "99" & "ORD"
        
        'Campo3 Importe total
        cad = cad & Format(Importe * 100, String(17, "0")) ' Importe
        cad = cad & Format(Regs, String(8, "0")) ' Importe
        
        'Igual que arriba as uno
        'Reg(numreo de adeudos + 1 reg01 + un reg02 + reg04  +1
        cad = cad & Format(Regs + 4, String(10, "0")) ' Importe
        cad = cad & FrmtStr("", 560)  'libre
        Print #NF, cad
        
        
        
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    Close (NF)
    If Regs > 0 Then GeneraFicheroNorma34SEPA = True
    Exit Function
EGen2:
    MuestraError Err.Number, Err.Description

End Function


Public Function FrmtStr(campo As String, longitud As Integer) As String
    FrmtStr = Mid(Trim(campo) & Space(longitud), 1, longitud)
End Function


Private Function IBAN_Destino() As String
    
        IBAN_Destino = FrmtStr(miRsAux!Iban, 4)  ' ES00
        IBAN_Destino = IBAN_Destino & Format(miRsAux!entidad, "0000") ' Código de entidad receptora
        IBAN_Destino = IBAN_Destino & Format(miRsAux!oficina, "0000") ' Código de oficina receptora
        IBAN_Destino = IBAN_Destino & Format(miRsAux!CC, "00") ' Dígitos de control
        IBAN_Destino = IBAN_Destino & Format(miRsAux!CuentaBa, "0000000000") ' Código de cuenta
'    Else
'
'        'entidad oficina CC cuentaba
'        IBAN_Destino = FrmtStr(miRsAux!IBAN, 4)  ' ES00
'        IBAN_Destino = IBAN_Destino & Format(miRsAux!entidad, "0000") ' Código de entidad receptora
'        IBAN_Destino = IBAN_Destino & Format(miRsAux!oficina, "0000") ' Código de oficina receptora
'        IBAN_Destino = IBAN_Destino & Format(miRsAux!CC, "00") ' Dígitos de control
'        IBAN_Destino = IBAN_Destino & Format(miRsAux!cuentaba, "0000000000") ' Código de cuenta
'    End If
End Function

Private Function DatosBasicosDelDeudor() As String
        DatosBasicosDelDeudor = FrmtStr(miRsAux!Nommacta, 70)
        'dirdatos,codposta,despobla,pais desprovi
        DatosBasicosDelDeudor = DatosBasicosDelDeudor & FrmtStr(miRsAux!dirdatos, 50)
        DatosBasicosDelDeudor = DatosBasicosDelDeudor & FrmtStr(Trim(DBLet(miRsAux!codposta, "T") & " " & miRsAux!desPobla), 50)
        DatosBasicosDelDeudor = DatosBasicosDelDeudor & FrmtStr(miRsAux!desProvi, 40)
        DatosBasicosDelDeudor = DatosBasicosDelDeudor & "ES"
        
        'If IsNull(miRsAux!PAIS) Then
        '    DatosBasicosDelDeudor = DatosBasicosDelDeudor & "ES"
        'Else
        '    DatosBasicosDelDeudor = DatosBasicosDelDeudor & Mid(miRsAux!PAIS, 1, 2)
        'End If
End Function






'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'
'
'
'
'
'               Norma 34 SEPA XML
'
'
'
'
'
'
'
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************


Public Function GeneraFicheroNorma34SEPA_XML(CIF As String, Fecha As Date, CuentaPropia2 As String, cadSQL As String, ConceptoTr As String, Tipo As Byte, SufijoOEM As String) As Boolean
Dim Regs As Integer
Dim Importe As Currency
Dim Im As Currency
Dim cad As String
Dim Aux As String
Dim NFic As Integer
Dim EsPersonaJuridica2 As Boolean
'Dim miRsAux As ADODB.Recordset

    On Error GoTo EGen3
    GeneraFicheroNorma34SEPA_XML = False

    NFic = -1


    Set miRsAux = New ADODB.Recordset

    cad = RecuperaValor(CuentaPropia2, 5) & RecuperaValor(CuentaPropia2, 1) & RecuperaValor(CuentaPropia2, 2) & RecuperaValor(CuentaPropia2, 3) & RecuperaValor(CuentaPropia2, 4)
    CuentaPropia2 = cad

    If Len(cad) <> 24 Then
        MsgBox "Error IBAN banco : " & CuentaPropia2, vbExclamation
        Exit Function
    End If

    'Esta comprobacion deberia hacerla antes
'
'    Cad = "SELECT tmpNorma34.CodSoc, tmpNorma34.Nombre, tmpNorma34.Banco1, tmpNorma34.Banco2, tmpNorma34.Banco3"
'    Cad = Cad & ", tmpNorma34.Banco4, tmpNorma34.Domicilio, tmpNorma34.Codpos, tmpNorma34.Poblacion, tmpNorma34.Concepto,"
'    Cad = Cad & "tmpNorma34.Importe, tmpNorma34.tipo"
'
'    Cad = Cad & ",Trabajadores.*, sbic.bic"
'    Cad = Cad & " FROM (tmpNorma34 INNER JOIN Trabajadores ON tmpNorma34.CodSoc = Trabajadores.IdTrabajador)"
'    Cad = Cad & " LEFT JOIN sbic ON Trabajadores.entidad = sbic.entidad;"
'    miRsAux.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
'
'    If Not miRsAux.EOF Then
'        Cad = "#"
'        While Not miRsAux.EOF
'            If IsNull(miRsAux!Bic) Then
'                If InStr(1, Cad, "#" & miRsAux!banco1 & "#") = 0 Then Cad = Cad & miRsAux!banco1 & "#"
'            End If
'            miRsAux.MoveNext
'        Wend
'        miRsAux.MoveFirst
'
'
'        If Len(Cad) > 1 Then
'            Cad = Mid(Cad, 2)
'            Cad = Mid(Cad, 1, Len(Cad) - 1)
'            Cad = Replace(Cad, "#", "   /   ")
'            Cad = "Bancos sin BIC asignado:" & vbCrLf & Cad & vbCrLf & vbCrLf & "¿Continuar?"
'            If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then
'                miRsAux.Close
'                Set miRsAux = Nothing
'                Exit Function
'            End If
'        End If
'
'    End If
'    miRsAux.Close
'










    NFic = FreeFile
    Open App.Path & "\norma34.txt" For Output As NFic


    Print #NFic, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    Print #NFic, "<Document xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""urn:iso:std:iso:20022:tech:xsd:pain.001.001.03"">"
    Print #NFic, "<CstmrCdtTrfInitn>"
    Print #NFic, "   <GrpHdr>"

    '                   NumeroTransferencia
    cad = "TRANPAG" & Format(0, "000000") & "F" & Format(Now, "yyyymmddThhnnss")
    Print #NFic, "      <MsgId>" & cad & "</MsgId>"
    Print #NFic, "      <CreDtTm>" & Format(Now, "yyyy-mm-ddThh:nn:ss") & "</CreDtTm>"

    'Registrp cabecera con totales

    Aux = "importe"
    cad = "tmpimpor"

    cad = "Select count(*),sum(" & Aux & ") FROM " & cad & " WHERE 1 =1"
    Aux = "0|0|"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(1)) Then Aux = miRsAux.Fields(0) & "|" & Format(miRsAux.Fields(1), "#.00") & "|"
    End If
    miRsAux.Close

'    '[Monica]
'    'Nombre
'    miRsAux.Open "Select * from empresas", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    Cad = Cad & FrmtStr(miRsAux!nomempre, 70)
'
'        'Direccion   nomempre domempre codpobla pobempre proempre
'        Cad = Cad & FrmtStr(Trim(miRsAux!domempre), 50)
'        Cad = Cad & FrmtStr(Trim(DBLet(miRsAux!CodPobla, "T") & " " & miRsAux!pobempre), 50)
'        Cad = Cad & FrmtStr(DBLet(miRsAux!proempre, "T"), 40)
'
'    miRsAux.Close
'    '[Monica]


    Print #NFic, "      <NbOfTxs>" & RecuperaValor(Aux, 1) & "</NbOfTxs>"
    Print #NFic, "      <CtrlSum>" & TransformaComasPuntos(RecuperaValor(Aux, 2)) & "</CtrlSum>"
    Print #NFic, "      <InitgPty>"
    Print #NFic, "         <Nm>" & XML(vEmpresa.nomempre) & "</Nm>"  'MiEmpresa.NomEmpresa
    Print #NFic, "         <Id>"
    cad = Mid(CIF, 1, 1)

    EsPersonaJuridica2 = Not IsNumeric(cad)




    cad = "PrvtId"
    If EsPersonaJuridica2 Then cad = "OrgId"

    Print #NFic, "           <" & cad & ">"
    Print #NFic, "               <Othr>"
    Print #NFic, "                  <Id>" & CIF & SufijoOEM & "</Id>"
    Print #NFic, "               </Othr>"
    Print #NFic, "           </" & cad & ">"

    Print #NFic, "         </Id>"
    Print #NFic, "      </InitgPty>"
    Print #NFic, "   </GrpHdr>"

    Print #NFic, "   <PmtInf>"

    Print #NFic, "      <PmtInfId>" & Format(Now, "yyyymmddhhnnss") & CIF & "</PmtInfId>"
    Print #NFic, "      <PmtMtd>TRF</PmtMtd>"
    Print #NFic, "      <ReqdExctnDt>" & Format(Fecha, "yyyy-mm-dd") & "</ReqdExctnDt>"
    Print #NFic, "      <Dbtr>"

     'Nombre
    Print #NFic, "         <Nm>" & XML(vEmpresa.nomempre) & "</Nm>" 'MiEmpresa.NomEmpresa
    Print #NFic, "         <PstlAdr>"
    Print #NFic, "            <Ctry>ES</Ctry>"

    cad = vParam.DomicilioEmpresa & " "
    cad = cad & Trim(vParam.Poblacion) & " " & vParam.Provincia & " "

    Print #NFic, "            <AdrLine>" & XML(Trim(cad)) & "</AdrLine>"

    Print #NFic, "         </PstlAdr>"
    Print #NFic, "         <Id>"

    Aux = "PrvtId"
    If EsPersonaJuridica2 Then Aux = "OrgId"


    Print #NFic, "            <" & Aux & ">"

    Print #NFic, "               <Othr>"
    Print #NFic, "                  <Id>" & CIF & SufijoOEM & "</Id>"
    Print #NFic, "               </Othr>"
    Print #NFic, "            </" & Aux & ">"
    Print #NFic, "         </Id>"
    Print #NFic, "    </Dbtr>"


    Print #NFic, "    <DbtrAcct>"
    Print #NFic, "       <Id>"
    Print #NFic, "          <IBAN>" & Trim(CuentaPropia2) & "</IBAN>"
    Print #NFic, "       </Id>"
    Print #NFic, "       <Ccy>EUR</Ccy>"
    Print #NFic, "    </DbtrAcct>"
    Print #NFic, "    <DbtrAgt>"
    Print #NFic, "       <FinInstnId>"

    cad = Mid(CuentaPropia2, 5, 4)
    
    '[Monica]02/05/2017: sbic es bics en ariconta
    If vParamAplic.ContabilidadNueva Then
        cad = DevuelveDesdeBDNew(cConta, "bics", "bic", "entidad", cad, "T")
    Else
        cad = DevuelveDesdeBDNew(cConta, "sbic", "bic", "entidad", cad, "T")
    End If
    
'    Dim SqlBic As String
'    Dim RsBic As ADODB.Recordset
'    SqlBic = "select bic from sbic where entidad = " & DBSet(cad, "N")
'
'    cad = ""
'
'    Set RsBic = New ADODB.Recordset
'    RsBic.Open SqlBic, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not RsBic.EOF Then
'        cad = DBLet(RsBic.Fields(0).Value, "T")
'    End If
    
    
    Print #NFic, "          <BIC>" & Trim(cad) & "</BIC>"
    Print #NFic, "       </FinInstnId>"
    Print #NFic, "    </DbtrAgt>"



    If cadSQL = "" Then
'        Cad = "SELECT tmpNorma34.CodSoc, tmpNorma34.Nombre, tmpNorma34.Banco1, tmpNorma34.Banco2, tmpNorma34.Banco3"
'        Cad = Cad & ", tmpNorma34.Banco4, tmpNorma34.Domicilio, tmpNorma34.Codpos, tmpNorma34.Poblacion, tmpNorma34.Concepto,"
'        Cad = Cad & "tmpNorma34.Importe, tmpNorma34.tipo,Trabajadores.*, sbic.bic"
'        Cad = Cad & " FROM (tmpNorma34 INNER JOIN Trabajadores ON tmpNorma34.CodSoc = Trabajadores.IdTrabajador)"
'        Cad = Cad & " LEFT JOIN sbic ON Trabajadores.entidad = sbic.entidad;"
    
        Dim Sql As String
    
        Sql = "select tmpimpor.codtraba, tmpimpor.importe, straba.codbanco as entidad,straba.codsucur "
        Sql = Sql & " as oficina,straba.digcontr as CC,straba.cuentaba as cuentaba ,straba.nomtraba as nombre, "
        Sql = Sql & " straba.domtraba as dirdatos, straba.codpobla as codposta,straba.pobtraba as despobla, straba.protraba as desprovi, "
        Sql = Sql & " straba.niftraba as refbenef,straba.iban"
        Sql = Sql & " from tmpimpor, straba"
        Sql = Sql & " where tmpimpor.codtraba = straba.codtraba "
    
        cadSQL = Sql
    
    End If

    miRsAux.Open cadSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    Regs = 0
    While Not miRsAux.EOF
        Print #NFic, "   <CdtTrfTxInf>"
        Print #NFic, "      <PmtId>"


        'IDentificador
         If IsNull(miRsAux!refbenef) Then
            Aux = DBLet(miRsAux!Concepto, "T") & " Tra:" & Format(miRsAux!CodTraba, "0000") & " F:" & Format(Fecha, "dd/mm")
        Else
            Aux = miRsAux!refbenef
        End If


        Print #NFic, "         <EndToEndId>" & Aux & "</EndToEndId>"
        Print #NFic, "      </PmtId>"
        Print #NFic, "      <PmtTpInf>"

        'Importe
        Im = miRsAux!Importe



        'Persona fisica o juridica
        cad = DBLet(miRsAux!refbenef, "T")
        cad = Mid(cad, 1, 1)
        EsPersonaJuridica2 = Not IsNumeric(cad)
        'Como da problemas Cajamar, siempre ponemos Perosna juridica. Veremos
        EsPersonaJuridica2 = True


        Importe = Importe + Im
        Regs = Regs + 1

        Print #NFic, "          <SvcLvl><Cd>SEPA</Cd></SvcLvl>"
'        If Tipo = "1" Then
'            Aux = "SALA"
'        ElseIf Tipo = "0" Then
'            Aux = "PENS"
'        Else
'            Aux = "TRAD"
'        End If
        Select Case Tipo
            Case 0:
                Aux = "SALA"
            Case 1:
                Aux = "PENS"
            Case Else
                Aux = "TRAD"
        End Select


        Print #NFic, "          <CtgyPurp><Cd>" & Aux & "</Cd></CtgyPurp>"
        Print #NFic, "       </PmtTpInf>"
        Print #NFic, "       <Amt>"
        Print #NFic, "          <InstdAmt Ccy=""EUR"">" & TransformaComasPuntos(CStr(Im)) & "</InstdAmt>"
        Print #NFic, "       </Amt>"
        Print #NFic, "       <CdtrAgt>"
        Print #NFic, "          <FinInstnId>"
        
        Dim Bic As String
        If vParamAplic.ContabilidadNueva Then
            Bic = DevuelveDesdeBDNew(cConta, "bics", "bic", "entidad", miRsAux!entidad, "N")
        Else
            Bic = DevuelveDesdeBDNew(cConta, "sbic", "bic", "entidad", miRsAux!entidad, "N")
        End If
        
        cad = DBLet(Bic, "T")
        If cad = "" Then Err.Raise 513, , "No existe BIC " & vbCrLf & "Entidad: " & miRsAux!entidad
        Print #NFic, "             <BIC>" & DBLet(Bic, "T") & "</BIC>"
        Print #NFic, "          </FinInstnId>"
        Print #NFic, "       </CdtrAgt>"
        Print #NFic, "       <Cdtr>"
        Print #NFic, "          <Nm>" & XML(miRsAux!Nombre) & "</Nm>"


        'Como cajamar da problemas, lo quitamos para todos
        'Print #NFic, "          <PstlAdr>"
        '
        'Cad = "ES"
        'If Not IsNull(miRsAux!PAIS) Then Cad = Mid(miRsAux!PAIS, 1, 2)
        'Print #NFic, "              <Ctry>" & Cad & "</Ctry>"
        '
        'If Not IsNull(miRsAux!dirdatos) Then Print #NFic, "              <AdrLine>" & XML(miRsAux!dirdatos) & "</AdrLine>"
        'Cad = XML(Trim(DBLet(miRsAux!codposta, "T") & " " & DBLet(miRsAux!despobla, "T")))
        'If Cad <> "" Then Print #NFic, "              <AdrLine>" & Cad & "</AdrLine>"
        'If Not IsNull(miRsAux!desprovi) Then Print #NFic, "              <AdrLine>" & XML(miRsAux!desprovi) & "</AdrLine>"
        'Print #NFic, "           </PstlAdr>"



        Print #NFic, "           <Id>"
        Aux = "PrvtId"
        If EsPersonaJuridica2 Then Aux = "OrgId"

        Print #NFic, "               <" & Aux & ">"
        Print #NFic, "                  <Othr>"

        Print #NFic, "                     <Id>" & miRsAux!refbenef & "</Id>"
        'Da problemas.... con Cajamar
        'Print #NFic, "                     <Issr>NIF</Issr>"
        Print #NFic, "                  </Othr>"
        Print #NFic, "               </" & Aux & ">"
        Print #NFic, "           </Id>"
        Print #NFic, "        </Cdtr>"
        Print #NFic, "        <CdtrAcct>"
        Print #NFic, "           <Id>"
        Print #NFic, "              <IBAN>" & IBAN_Destino() & "</IBAN>"
        Print #NFic, "           </Id>"
        Print #NFic, "        </CdtrAcct>"
        Print #NFic, "      <Purp>"


'        If Tipo = "1" Then
'            Aux = "SALA"
'        ElseIf Tipo = "0" Then
'            Aux = "PENS"
'        Else
'            Aux = "TRAD"
'        End If
        Select Case Tipo
            Case 0:
                Aux = "SALA"
            Case 1:
                Aux = "PENS"
            Case Else
                Aux = "TRAD"
        End Select


        Print #NFic, "         <Cd>" & Aux & "</Cd>"
        Print #NFic, "      </Purp>"
        Print #NFic, "      <RmtInf>"

        Aux = DBLet(ConceptoTr, "T") & " " & DBLet(Fecha, "T") & " Importe " & Format(Im, FormatoImporte)
        Print #NFic, "         <Ustrd>" & XML(Trim(Aux)) & "</Ustrd>"
        Print #NFic, "      </RmtInf>"
        Print #NFic, "   </CdtTrfTxInf>"




        miRsAux.MoveNext
    Wend
    Print #NFic, "   </PmtInf>"
    Print #NFic, "</CstmrCdtTrfInitn></Document>"


    miRsAux.Close
    Set miRsAux = Nothing
    Close (NFic)
    NFic = -1
    If Regs > 0 Then GeneraFicheroNorma34SEPA_XML = True
    Exit Function
EGen3:
    MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    If NFic > 0 Then Close (NFic)
End Function



Private Function XML(cadena As String) As String
Dim I As Integer
Dim Aux As String
Dim Le As String
Dim c As Integer
    'Carácter no permitido en XML  Representación ASCII
    '& (ampersand)          &amp;
    '< (menor que)          &lt;
    ' > (mayor que)         &gt;
    '" (dobles comillas)    &quot;
    '' (apóstrofe)          &apos;

    'La ISO recomienda trabajar con los carcateres:
    'a b c d e f g h i j k l m n o p q r s t u v w x y z
    'A B C D E F G H I J K L M N O P Q R S T U V W X Y Z
    '0 1 2 3 4 5 6 7 8 9
    '/ - ? : ( ) . , ' +
    'Espacio
    Aux = ""
    For I = 1 To Len(cadena)
        Le = Mid(cadena, I, 1)
        c = Asc(Le)


        Select Case c
        Case 40 To 57
            'Caracteres permitidos y numeros

        Case 65 To 90
            'Letras mayusculas

        Case 97 To 122
            'Letras minusculas

        Case 32
            'espacio en balanco

        Case Else
            Le = " "
        End Select
        Aux = Aux & Le
    Next
    XML = Aux
End Function



