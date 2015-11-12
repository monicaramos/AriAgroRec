Attribute VB_Name = "Norma19"
Option Explicit

Dim AuxD As String
Private NumeroTransferencia As Integer



Public Function GeneraFicheroNorma19(DatosBanco As String, FecPre As String, DatosExtra As String, TipoReferenciaCliente As Byte, FecCobro As Date) As Boolean
Dim ValorEnOpcionales As Boolean
    '-- Genera_Remesa: Esta función genera la remesa indicada, en el fichero correspondiente
    Dim mAux As String
    Dim SumaImportes As Currency
    Dim SumReg As Integer
    Dim SumTotal As Integer
    Dim DatosBanco As String  'oficina,sucursla,cta, sufijo
    Dim vSufijo As String
    Dim NifEmpresa As String
    Dim ImprimeOpc As Boolean
    Dim ValoresOpcionales As String
    Dim ImpEfe As Currency
    Dim J As Integer
    On Error GoTo Err_Remesa
    
    
    vSufijo = RecuperaValor(DatosExtra, 1)
    If Trim(vSufijo) = "" Then vSufijo = Mid(RSaux!sufijoem & "   ", 1, 3)
     'En datos extra dejo el CONCEPTO PPAL
    DatosExtra = RecuperaValor(DatosExtra, 2)
    
    
    If DatosBanco = "" Then Exit Function
    
    'If Not comprobarCuentasBancariasRecibos(Remesa) Then Exit Function
    
    
    'Ahora cargare el NIF y la empresa
    NifEmpresa = vParam.CifEmpresa
    
    '-- Abrir el fichero a enviar
    NF = FreeFile()
    Open App.Path & "\norma19.txt" For Output As #NF
    
    Sql = "select  tmpImporNeg.*,straba.nomtraba,straba.niftraba,straba.codbanco,straba.codsucur,straba.digcontr,straba.cuentaba from tmpImporNeg, straba where "
    Sql = Sql & " tmpImporNeg.codtraba = straba.codtraba "
    Sql = Sql & " order by tmpImporNeg.codtraba "
    
    'EL ORDEN QUE QUERAMOS
    Remesa = RecuperaValor(Remesa, 1)
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        'Rs.MoveFirst
    
'        '-- Registro 5180
        registro = "5180"
        registro = registro & FrmtStr(NifEmpresa, 9)   '-- Alinea NIF
        registro = registro & FrmtStr(vSufijo, 3) ' Sufijo
        registro = registro & Format(FecPre, "ddmmyy") ' Fecha de presentación
        registro = registro & FrmtStr(" ", 6) ' LIBRE
        registro = registro & FrmtStr(DatosExtra, 40)   ' Nombre del cliente presentador
        registro = registro & FrmtStr(" ", 20) ' LIBRE
        registro = registro & RecuperaValor(DatosBanco, 1)
        registro = registro & RecuperaValor(DatosBanco, 2)  ' Código de oficina receptora
        'IDENDIFICADOR DE REMESA
        '12 caracteres
        'Registro = Registro & "RE" & Format(vEmpresa.codempre, "00") & Format(miRsAux!codrem, "0000") & Format(miRsAux!Anyorem, "0000")
        registro = registro & FrmtStr(" ", 12) ' LIBRE
        registro = registro & FrmtStr(" ", 54) ' LIBRE
        SumTotal = SumTotal + 1
        Print #NF, registro
       
        
        '-- Registro 5380
        registro = "5380"
        registro = registro & FrmtStr(NifEmpresa, 9)  '-- Alinea NIF
        registro = registro & FrmtStr(vSufijo, 3) ' Sufijo
        registro = registro & Format(FecPre, "ddmmyy") ' Fecha de confección del soporte
        registro = registro & Format(FecCobro, "ddmmyy") ' Fecha de cargo de recibos'
        registro = registro & FrmtStr(DatosExtra, 40)  ' Nombre del cliente presentador
        registro = registro & RecuperaValor(DatosBanco, 1) ' Código de entidad receptora
        registro = registro & RecuperaValor(DatosBanco, 2) ' Código de oficina receptora
        registro = registro & RecuperaValor(DatosBanco, 3) 'Dígitos de control
        registro = registro & RecuperaValor(DatosBanco, 4) ' Código de cuenta
        registro = registro & FrmtStr(" ", 8) ' LIBRE
        registro = registro & "01" ' Fijo 01
        'Nuevo 24 Febrero 2006
        'Registro = Registro & "RE" & Format(vEmpresa.codempre, "00") & Format(miRsAux!codrem, "0000") & Format(miRsAux!Anyorem, "0000")
        registro = registro & FrmtStr(" ", 12) ' LIBRE
        
        registro = Mid(registro & Space(100), 1, 162)
        SumTotal = SumTotal + 1
        Print #NF, registro
        '-- Leemos secuencialmente las líneas de remesa
        While Not miRsAux.EOF
            'Tenemos k ver si imprimimos los opcionales
            ImprimeOpc = HayKImprimirOpcionales
        
            ValoresOpcionales = FrmtStr(NifEmpresa, 9)  '-- Alinea NIF
            ValoresOpcionales = ValoresOpcionales & FrmtStr(vSufijo, 3) ' Sufijo
            'Segun szea lo que quiera el cliente que le ponga como referencia
            Select Case TipoReferenciaCliente
            Case 1
                'ALZIRA. La referencia final de 12 es el ctan bancaria del cli + su CC
                    registro = Format(miRsAux!digcontr, "00") ' Dígitos de control
                    registro = registro & Format(miRsAux!cuentaba, "0000000000") ' Código de cuenta
            Case 2
                'NIF
                registro = DBLet(miRsAux!Niftraba, "T")
                If registro = "" Then registro = miRsAux!Codmacta
                registro = Mid(registro & Space(12), 1, 12)
                
            Case Else
                'Antes
                'Registro = miRsAux!NUmSerie & Format(miRsAux!codfaccl, "0000000000") & Format(miRsAux!numorden, "0")
                registro = miRsAux!Codmacta
                registro = Right("0000000000" & registro, 12)
            End Select
            registro = Mid(registro, 1, 12)
            ValoresOpcionales = ValoresOpcionales & registro
            
            'Registro = Registro & FrmtStr(NifEmpresa, 9)  '-- Alinea NIF
            'Registro = Registro & FrmtStr(vSufijo, 3) ' Sufijo
            'Registro = Registro & FrmtStr(miRsAux!NUmSerie, 3) & FrmtStr(miRsAux!codfaccl, 7) & "-" & FrmtStr(miRsAux!numorden, 1)
            
            
            '-- Registro 5680
            registro = "5680"
            registro = registro & ValoresOpcionales
            
            
            
            registro = registro & FrmtStr(DevNombreSQL(miRsAux!Nomtraba), 40)
            registro = registro & Format(miRsAux!codbanco, "0000") ' Código de entidad receptora
            registro = registro & Format(miRsAux!codsucur, "0000") ' Código de oficina receptora
            registro = registro & Format(miRsAux!digcontr, "00") ' Dígitos de control
            registro = registro & Format(miRsAux!cuentaba, "0000000000") ' Código de cuenta
            
            ImpEfe = miRsAux!Importe
            registro = registro & Format(ImpEfe * 100, String(10, "0")) ' Importe
            
            
            'ANTES
            'Registro = Registro & FrmtStr(Left(miRsAux!codmacta, 6), 6)  ' Identificador de domiciliación
            'Registro = Registro & miRsAux!NUmSerie & FrmtStr(Format(miRsAux!codfaccl, "000000000"), 9)
            'AHORA
            registro = registro & Format(miRsAux!fecfaccl, "ddmmyy")  ' Identificador de domiciliación
            registro = registro & miRsAux!numserie & FrmtStr(Format(miRsAux!codfaccl, "00000000"), 8) & Format(miRsAux!numorden, "0")
            
            registro = registro & Space(16)
            
            'Registro = Registro & FrmtStr(mAux, 10) ' Identificador de devolución
            mAux = DBLet(miRsAux!Concepto, "T")
            If mAux = "" Then mAux = "Nomina"
            
            registro = registro & FrmtStr(mAux, 40) ' Primer Concepto
            registro = registro & Format(miRsAux!FecVenci, "ddmmyy") & "  "
            Print #NF, registro
            SumReg = SumReg + 1
            SumTotal = SumTotal + 1
            SumaImportes = SumaImportes + ImpEfe
            
            If ImprimeOpc Then
                For J = 1 To 5
                    registro = ImprimeOpcionales(True, ValoresOpcionales, J, ValorEnOpcionales)
                    If ValorEnOpcionales Then
                        Print #NF, registro
                        SumTotal = SumTotal + 1
                    End If
                Next J
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        
        '-- Registro 5880
        registro = "5880"
        registro = registro & FrmtStr(NifEmpresa, 9)  '-- Alinea NIF
        registro = registro & FrmtStr(vSufijo, 3) ' Sufijo
        registro = registro & FrmtStr(" ", 72) ' LIBRE
        registro = registro & Format(SumaImportes * 100, String(10, "0")) ' Suma de importes
        registro = registro & FrmtStr(" ", 6) ' LIBRE
        registro = registro & Format(SumReg, String(10, "0")) ' Suma de registros 0680
        registro = registro & Format(SumTotal, String(10, "0")) ' Suma total de registros
        registro = registro & FrmtStr(" ", 38) ' LIBRE
        SumTotal = SumTotal + 1
        Print #NF, registro
        '-- Registro 5980
        registro = "5980"
        registro = registro & FrmtStr(NifEmpresa, 9)  '-- Alinea NIF
        registro = registro & FrmtStr(vSufijo, 3) ' Sufijo
        registro = registro & FrmtStr(" ", 52) ' LIBRE
        registro = registro & "0001" ' Suma de ordenantes (siempre es uno)
        registro = registro & FrmtStr(" ", 16) ' LIBRE
        registro = registro & Format(SumaImportes * 100, String(10, "0")) ' Suma de importes
        registro = registro & FrmtStr(" ", 6) ' LIBRE
        registro = registro & Format(SumReg, String(10, "0"))  ' Suma de registros 0680
        SumTotal = SumTotal + 1
        registro = registro & Format(SumTotal, String(10, "0")) ' Suma total de registros
        registro = registro & FrmtStr(" ", 38) ' LIBRE
        Print #NF, registro
    End If
    Close #NF
    If SumTotal > 0 Then GeneraFicheroNorma19 = True
    Exit Function
Err_Remesa:
    MsgBox "Err: " & Err.Number & vbCrLf & _
        Err.Description, vbCritical, "Grabación del fichero Pagos"
        
End Function












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
            FileCopy App.Path & "\Norma19.txt", "a:\Norma19.txt"
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
        FileCopy App.Path & "\Norma19.txt", cad
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
Public Function GeneraFicheroNorma19New(CIF As String, Fecha As Date, CuentaPropia As String, ConceptoTransferencia As String, vNumeroTransferencia As Integer, ByRef ConceptoTr As String, CodigoOrden As String, ConcepTransf As Byte) As Boolean
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
    GeneraFicheroNorma19New = False
    
    NumeroTransferencia = vNumeroTransferencia
    
    If CodigoOrden = "" Then
'        Aux = Right("    " & CIF, 10)
        Aux = RellenaABlancos(CIF, True, 10)
    Else
        Aux = CodigoOrden
    End If

    NFich = FreeFile
    Open App.Path & "\Norma19.txt" For Output As #NFich
    
    
    'Codigo ordenante
    '---------------------------------------------------
    'Si el banco tiene puesto si ID de Norma19 entonces
    'la pongo aquin. Lo he cargado antes sobre la variable AUX
    CodigoOrdenante = Aux 'Left(Aux & "          ", 10)  'CIF EMPRESA
    
    
    'CABECERA
    Cabecera1 NFich, CodigoOrdenante, Fecha, CuentaPropia, cad
    Cabecera2 NFich, CodigoOrdenante, cad
    Cabecera3 NFich, CodigoOrdenante, cad
    Cabecera4 NFich, CodigoOrdenante, cad
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma19
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
            Aux = RellenaABlancos(DBLet(Rs!Niftraba, "T"), True, 12)
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
    If Regs > 0 Then GeneraFicheroNorma19New = True
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description

End Function




Private Function RellenaABlancos(Cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim cad As String
    
    cad = Space(longitud)
    If PorLaDerecha Then
        cad = Cadena & cad
        RellenaABlancos = Left(cad, longitud)
    Else
        cad = cad & Cadena
        RellenaABlancos = Right(cad, longitud)
    End If
    
End Function



Private Function RellenaAceros(Cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim cad As String
    
    cad = Mid("00000000000000000000", 1, longitud)
    If PorLaDerecha Then
        cad = Cadena & cad
        RellenaAceros = Left(cad, longitud)
    Else
        cad = cad & Cadena
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
    cad = cad & RellenaAceros(CStr(DBLet(RS1!cuentaba, "T")), False, 10) 'Cta
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
        Aux = Trim(DBLet(RS1!text1csb, "T"))
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
    cad = cad & RellenaABlancos(DBLet(RS1!text2csb, "T"), False, 35)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub




Private Sub Totales(NF As Integer, ByRef CodOrde As String, total As Currency, Registros As Integer, ByRef cad As String, Pagos As Boolean)
    cad = "08" & "56"
    cad = cad & CodOrde    'llevara tb la ID del socio
    cad = cad & Space(15)
    cad = cad & RellenaAceros(CStr(Int(Round(total * 100, 2))), False, 12)
    cad = cad & RellenaAceros(CStr(Registros), False, 8)
    If Pagos Then
        cad = cad & RellenaAceros(CStr((Registros * 7) + 4 + 1), False, 10)
    Else
        cad = cad & RellenaAceros(CStr((Registros * 6) + 4 + 1), False, 10)
    End If
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub
