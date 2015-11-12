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
Dim i As Integer
Dim Cad As String

On Error Resume Next

    CopiarEnDisquette = False
    
    If A_disquetera Then
        For i = 1 To Intentos
            Cad = "Introduzca un disco vacio. (" & i & ")"
            MsgBox Cad, vbInformation
            FileCopy App.Path & "\norma34.txt", "a:\norma34.txt"
            If Err.Number <> 0 Then
                MuestraError Err.Number, "Copiar En Disquette"
            Else
                CopiarEnDisquette = True
                Exit For
            End If
        Next i
    Else
        If AuxD = "" Then
            Cad = Format(Now, "ddmmyyhhnn")
            Cad = App.Path & "\" & Cad & ".txt"
        Else
            Cad = AuxD
        End If
        FileCopy App.Path & "\norma34.txt", Cad
        If Err.Number <> 0 Then
            MsgBox "Error creando copia fichero. Consulte soporte técnico." & vbCrLf & Err.Description, vbCritical
        Else
            MsgBox "El fichero esta guardado como: " & Cad, vbInformation
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
Dim Cad As String


    On Error GoTo EGen
    GeneraFicheroNorma34 = False
    
    NumeroTransferencia = vNumeroTransferencia
    
    
    'Cargamos la cuenta
    Cad = "Select * from ctabancaria where codmacta='" & CuentaPropia & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = Right("    " & CIF, 10)
    If Rs.EOF Then
        Cad = ""
    Else
        If IsNull(Rs!entidad) Then
            Cad = ""
        Else
            Cad = Format(Rs!entidad, "0000") & "|" & Format(DBLet(Rs!oficina, "T"), "0000") & "|" & DBLet(Rs!Control, "T") & "|" & Format(DBLet(Rs!CtaBanco, "T"), "0000000000") & "|"
            CuentaPropia = Cad
        End If
        
        'Identificador norma bancaria
        If Not IsNull(Rs!idnorma34) Then Aux = Rs!idnorma34
    End If
    Rs.Close
    Set Rs = Nothing
    If Cad = "" Then
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
    Cabecera1 NFich, CodigoOrdenante, Fecha, CuentaPropia, Cad
    Cabecera2 NFich, CodigoOrdenante, Cad
    Cabecera3 NFich, CodigoOrdenante, Cad
    Cabecera4 NFich, CodigoOrdenante, Cad
    
    
    
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
            Linea2 NFich, Aux, Rs, Cad
            Linea3 NFich, Aux, Rs, Cad
            Linea4 NFich, Aux, Rs, Cad
            Linea5 NFich, Aux, Rs, Cad
            Linea6 NFich, Aux, Rs, Cad, ConceptoTr, Pagos
            If Pagos Then Linea7 NFich, Aux, Rs, Cad
        
        
        
        
            Importe = Importe + Im
            Regs = Regs + 1
            Rs.MoveNext
        Wend
        'Imprimimos totales
        Totales NFich, CodigoOrdenante, Importe, Regs, Cad, Pagos
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
Dim Cad As String
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
    Cabecera1 NFich, CodigoOrdenante, Fecha, CuentaPropia, Cad
    Cabecera2 NFich, CodigoOrdenante, Cad
    Cabecera3 NFich, CodigoOrdenante, Cad
    Cabecera4 NFich, CodigoOrdenante, Cad
    
    
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
        
        
            Linea1 NFich, Aux, Rs, Im, Cad, Concepto, ConceptoTransferencia
            Linea2 NFich, Aux, Rs, Cad
            Linea3 NFich, Aux, Rs, Cad
            Linea4 NFich, Aux, Rs, Cad
            Linea5 NFich, Aux, Rs, Cad
            Linea6 NFich, Aux, Rs, Cad, ConceptoTransferencia, Pagos
            If Pagos Then Linea7 NFich, Aux, Rs, Cad
        
            Importe = Importe + Im
            Regs = Regs + 1
            Rs.MoveNext
        Wend
        'Imprimimos totales
        Totales NFich, CodigoOrdenante, Importe, Regs, Cad, Pagos
    End If
    Rs.Close
    Set Rs = Nothing
    Close (NFich)
    If Regs > 0 Then GeneraFicheroNorma34New = True
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description

End Function




Private Function RellenaABlancos(Cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim Cad As String
    
    Cad = Space(longitud)
    If PorLaDerecha Then
        Cad = Cadena & Cad
        RellenaABlancos = Left(Cad, longitud)
    Else
        Cad = Cad & Cadena
        RellenaABlancos = Right(Cad, longitud)
    End If
    
End Function



Private Function RellenaAceros(Cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim Cad As String
    
    Cad = Mid("00000000000000000000", 1, longitud)
    If PorLaDerecha Then
        Cad = Cadena & Cad
        RellenaAceros = Left(Cad, longitud)
    Else
        Cad = Cad & Cadena
        RellenaAceros = Right(Cad, longitud)
    End If
    
End Function



'Private Sub Cabecera1(NF As Integer,ByRef CodOrde As String)
'Dim Cad As String
'
'End Sub

Private Sub Cabecera1(NF As Integer, ByRef CodOrde As String, Fecha As Date, cta As String, ByRef Cad As String)

    Cad = "03"
    Cad = Cad & "56"
    'cad = cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "001"
    Cad = Cad & Format(Now, "ddmmyy")
    Cad = Cad & Format(Fecha, "ddmmyy")
    'Cuenta bancaria
    Cad = Cad & RecuperaValor(cta, 1)
    Cad = Cad & RecuperaValor(cta, 2)
    Cad = Cad & RecuperaValor(cta, 4)
    Cad = Cad & "0"  'Sin relacion
    Cad = Cad & "   " & RecuperaValor(cta, 3)  'Digito de control bancario
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub



Private Sub Cabecera2(NF As Integer, ByRef CodOrde As String, ByRef Cad As String)
    Cad = "03"
    Cad = Cad & "56"
    'cad = cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "002"
    
    Cad = Cad & RellenaABlancos(vParam.NombreEmpresa, True, 30)   'Nombre empresa
  
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Cabecera3(NF As Integer, ByRef CodOrde As String, ByRef Cad As String)
    Cad = "03"
    Cad = Cad & "56"
    'cad = cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "003"
    
    
'    AuxD = DevuelveDesdeBD("direccion", "empresa2", "codigo", 1, "N")
    Cad = Cad & RellenaABlancos(vParam.DomicilioEmpresa, True, 30) 'AuxD, True, 30)   'Nombre empresa
    Cad = Cad & RellenaABlancos("", True, 30)   'Nombre empresa
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub



Private Sub Cabecera4(NF As Integer, ByRef CodOrde As String, ByRef Cad As String)

    Cad = "03"
    Cad = Cad & "56"
    'cad = cad & " "
    Cad = Cad & CodOrde
    Cad = Cad & Space(12) & "004"
    
'    AuxD = DevuelveDesdeBD("codpos", "empresa2", "codigo", 1, "N")
    Cad = Cad & RellenaABlancos(vParam.CPostal, False, 5) '   AuxD, False, 5)
    Cad = Cad & " "
'    AuxD = DevuelveDesdeBD("provincia", "empresa2", "codigo", 1, "N")
    Cad = Cad & RellenaABlancos(vParam.Provincia, True, 30) 'AuxD, True, 30)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub



'ConceptoTransferencia
'1.- Abono nomina
'9.- Transferencia ordinaria
Private Sub Linea1(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef importe1 As Currency, ByRef Cad As String, vconcepto As Byte, vConceptoTransferencia As String)


   
    '
    Cad = CodOrde   'llevara tb la ID del socio
    Cad = Cad & "010"
    Cad = Cad & RellenaAceros(CStr(Round(importe1, 2) * 100), False, 12)
    
    Cad = Cad & RellenaAceros(CStr(DBLet(RS1!entidad, "N")), False, 4)    'Entidad
    Cad = Cad & RellenaAceros(CStr(DBLet(RS1!oficina, "N")), False, 4)  'Sucur
    Cad = Cad & RellenaAceros(CStr(DBLet(RS1!CuentaBa, "T")), False, 10) 'Cta
    Cad = Cad & "1" & Format(vconcepto, "0") '& vConceptoTransferencia
    Cad = Cad & "  "
    Cad = Cad & RellenaAceros(CStr(DBLet(RS1!CC, "T")), False, 2) 'CC
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea2(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "011"
    Cad = Cad & RellenaABlancos(DBLet(RS1!nommacta, "T"), False, 36)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea3(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "012"
    Cad = Cad & RellenaABlancos(DBLet(RS1!dirdatos, "T"), False, 36)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea4(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "013"
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea5(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)
    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "014"
    Cad = Cad & RellenaABlancos(DBLet(RS1!codposta, "T"), False, 5) & " "
    Cad = Cad & RellenaABlancos(DBLet(RS1!desPobla, "T"), False, 30)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea6(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String, ByRef ConceptoT As String, Pagos As Boolean)
Dim Aux As String

    Aux = ConceptoT
    If Pagos Then
        'Tiene dos campos para las descripcion. Si no tiene nada pondre la descripcion de la transferencia
        Aux = Trim(DBLet(RS1!Text1csb, "T"))
        If Aux = "" Then Aux = ConceptoT
    End If

    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "016"
    Cad = Cad & RellenaABlancos(Aux, False, 35)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub


Private Sub Linea7(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Cad As String)


    Cad = CodOrde    'llevara tb la ID del socio
    Cad = Cad & "017"
    Cad = Cad & RellenaABlancos(DBLet(RS1!Text2csb, "T"), False, 35)
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
End Sub




Private Sub Totales(NF As Integer, ByRef CodOrde As String, Total As Currency, Registros As Integer, ByRef Cad As String, Pagos As Boolean)
    Cad = "08" & "56"
    Cad = Cad & CodOrde    'llevara tb la ID del socio
    Cad = Cad & Space(15)
    Cad = Cad & RellenaAceros(CStr(Int(Round(Total * 100, 2))), False, 12)
    Cad = Cad & RellenaAceros(CStr(Registros), False, 8)
    If Pagos Then
        Cad = Cad & RellenaAceros(CStr((Registros * 7) + 4 + 1), False, 10)
    Else
        Cad = Cad & RellenaAceros(CStr((Registros * 6) + 4 + 1), False, 10)
    End If
    Cad = RellenaABlancos(Cad, True, 72)
    Print #NF, Cad
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
Dim Cad As String
Dim Aux As String
Dim NF As Integer
Dim Sql As String
Dim Bic As String




    On Error GoTo EGen2
    GeneraFicheroNorma34SEPA = False
    

    
    Set miRsAux = New ADODB.Recordset
    
    'Cargamos la cuenta
    
    
        
    Cad = RecuperaValor(CuentaPropia2, 5) & RecuperaValor(CuentaPropia2, 1) & RecuperaValor(CuentaPropia2, 2) & RecuperaValor(CuentaPropia2, 3) & RecuperaValor(CuentaPropia2, 4)
        
            
    If Len(Cad) <> 24 Then
        MsgBox "Error leyendo datos para: " & CuentaPropia2, vbExclamation
        Exit Function
    End If
    CuentaPropia2 = Cad
    NF = FreeFile
    Open App.Path & "\norma34.txt" For Output As #NF
    
    
    
    'SEPA
    '1.- Cabecera ordenante
    '------------------------------------------------------------------------
    Cad = "01" & "ORD" & "34145" & "001" & CIF
        
    'sufijo (Tenemos el OEM, que se utiliza para las otras normas antiguas
    '[Monica]29/01/2014: antes "000"
    Cad = Cad & Right("000" & CodigoOrden34, 3)
    Cad = Cad & Format(Now, "yyyymmdd")
    Cad = Cad & Format(Fecha, "yyyymmdd")
    Cad = Cad & "A" 'IBAN
     
    'EL IBAN propiamente
    Cad = Cad & FrmtStr(CuentaPropia2, 34)
    '[Monica]24/02/2014: hacemos un cargo único antes era un 1
    Cad = Cad & "0" 'Cargo por cada operacion
    'Nombre
    miRsAux.Open "Select * from empresas", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = Cad & FrmtStr(miRsAux!nomempre, 70)

        'Direccion   nomempre domempre codpobla pobempre proempre
        Cad = Cad & FrmtStr(Trim(miRsAux!domempre), 50)
        Cad = Cad & FrmtStr(Trim(DBLet(miRsAux!CodPobla, "T") & " " & miRsAux!pobempre), 50)
        Cad = Cad & FrmtStr(DBLet(miRsAux!proempre, "T"), 40)
    
    miRsAux.Close
    
    'Pais y libre
    Cad = Cad & "ES" & FrmtStr("", 311)
    Print #NF, Cad
  
  
  
    '2.- Registro cabecera TRANSFERENCIA
    '------------------------------------------------------------------------
    Cad = "02" & "SCT" & "34145" & CIF
        
    'sufijo (Tenemos el OEM, que se utiliza para las otras normas
    '[Monica]29/01/2014: antes "000"
    Cad = Cad & Right("000" & CodigoOrden34, 3)
    Cad = Cad & FrmtStr("", 578)
    Print #NF, Cad
    
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

    Cad = cadSQL
    

    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
            Cad = "03" & "SCT" & "34145" & "002"
            
            'Campo 5 . Referencia del ordenante
            
            
            Aux = miRsAux!refbenef & " " & Format(Fecha, "dd/mm/yyyy")
            Cad = Cad & FrmtStr(Aux, 35)
            
            'Campo 6
            Cad = Cad & "A"
            
            'IBAN
            Cad = Cad & FrmtStr(IBAN_Destino(), 34)
            
            'Campo8 Importe
            Cad = Cad & Format(Im * 100, String(11, "0")) ' Importe
            
            'Campo9
            Cad = Cad & "3" 'gastos compartidos
            'Campo 10
            Bic = DevuelveDesdeBDNew(cConta, "sbic", "bic", "entidad", miRsAux!entidad, "N")
            Cad = Cad & FrmtStr(Bic, 11) 'FrmtStr("", 11)  'BIC

            'nommacta,dirdatos,codposta,dirdatos,despobla,impvenci,scobro.codmacta
            'Datos Basicos del beneficiario
            Cad = Cad & DatosBasicosDelDeudor
            
            
                '`text33csb` `text41csb`
            'Aux = DBLet(miRsAux!text33csb, "T") & DBLet(miRsAux!text41csb, "T")
            Aux = DescripcionTrans ' he quitado esto de David & "   Cod: " & Format(miRsAux!CodTraba, "0000")
            Cad = Cad & FrmtStr(Aux, 140)
            
            'Campo17
            Cad = Cad & FrmtStr("", 35)  'Reservado
            
            'Campo18 y campo19
'            cad = cad & "TRAD"
'            cad = cad & "TRAD"
            Select Case Tipo
                Case 0:
                    Cad = Cad & "SALA"
                    Cad = Cad & "SALA"
                Case 1:
                    Cad = Cad & "PENS"
                    Cad = Cad & "PENS"
                Case 2:
                    Cad = Cad & "TRAD"
                    Cad = Cad & "TRAD"
            End Select
            
            
            Cad = Cad & FrmtStr("", 99)  'libre
            
            Print #NF, Cad
            
            miRsAux.MoveNext
        Wend
        
    
        'TOTALES
        '----------------------------------
        'Total trasnferencia SEPA
        'Campo 1,2
        Cad = "04" & "SCT"
        
        'Campo3 Importe total
        Cad = Cad & Format(Importe * 100, String(17, "0")) ' Importe
        Cad = Cad & Format(Regs, String(8, "0")) ' Importe
        'Total registros son
        'Reg(numreo de adeudos + 1 reg01 + un reg02 + reg04
        Cad = Cad & Format(Regs + 2, String(10, "0")) ' Importe
        Cad = Cad & FrmtStr("", 560)  'libre
        Print #NF, Cad
        
        'Total general
        Cad = "99" & "ORD"
        
        'Campo3 Importe total
        Cad = Cad & Format(Importe * 100, String(17, "0")) ' Importe
        Cad = Cad & Format(Regs, String(8, "0")) ' Importe
        
        'Igual que arriba as uno
        'Reg(numreo de adeudos + 1 reg01 + un reg02 + reg04  +1
        Cad = Cad & Format(Regs + 4, String(10, "0")) ' Importe
        Cad = Cad & FrmtStr("", 560)  'libre
        Print #NF, Cad
        
        
        
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
        DatosBasicosDelDeudor = FrmtStr(miRsAux!nommacta, 70)
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




