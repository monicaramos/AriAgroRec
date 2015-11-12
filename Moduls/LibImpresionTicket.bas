Attribute VB_Name = "LibImpresionTicket"
Option Explicit


Public Sub ImprimirElTicketDirecto2(NumTicket As String, FechaTicket As Date, Precio4Decimales As Boolean, Optional Entregado As Currency, Optional Cambio As Currency)   ' (RAFA/ALZIRA 05092006)
'    Dim NomImpre As String
  '  Dim FechaT As Date
    Dim RS1 As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim Rs4 As ADODB.Recordset
    Dim Sql As String
    Dim Lin As String ' línea de impresión
    Dim I As Integer
    Dim N As Integer
    Dim ImporteIva As Currency
    Dim EnEfectivo As Boolean
    
    Dim NomArtic As String
    Dim cajas As Long
    Dim Tara As Long
    
    Dim Nombre As String
    
    
On Error GoTo EImpTickD
   
    Printer.Font = "Courier New"
    
    
    '-- Obtenemos cabeceras y pies en un recordset (rs1)
    Sql = "select rentradas.*, rsocios.nomsocio, variedades.nomvarie, rpartida.nomparti, rsituacioncampo.nomsitua, clases.nomclase "
    Sql = Sql & " from rentradas, rsocios, variedades, rcampos, rpartida, rsituacioncampo, clases "
    Sql = Sql & " where rentradas.numnotac = " & DBSet(NumTicket, "N")
    Sql = Sql & " and rentradas.codsocio = rsocios.codsocio "
    Sql = Sql & " and rentradas.codvarie = variedades.codvarie "
    Sql = Sql & " and rentradas.codcampo = rcampos.codcampo "
    Sql = Sql & " and rcampos.codparti = rpartida.codparti "
    Sql = Sql & " and rcampos.codsitua = rsituacioncampo.codsitua "
    Sql = Sql & " and variedades.codclase = clases.codclase "
    
    Set RS1 = New ADODB.Recordset
    RS1.Open Sql, conn, adOpenForwardOnly
    If Not RS1.EOF Then
            '-- Impresión de la cabecera
'                Lin = "         1         2         3         4"
'                Printer.Print Lin
'                Lin = "1234567890123456789012345678901234567890"
'                Printer.Print Lin
            
            
            ' nombre empresa
            Lin = LineaCentrada(vParam.NombreEmpresa)
            If Lin <> "" Then Printer.Print Lin
            
            ' nombre empresa
            Lin = LineaCentrada("CIF: " & vParam.CifEmpresa)
            If Lin <> "" Then Printer.Print Lin
            '
            
            Printer.Print " "
            
            
            Lin = CuadraParteI(24, "Fecha: " & Format(RS1!FechaEnt, "dd/mm/yyyy") & "  " & Format(RS1!horaentr, "hh:mm")) & _
                  CuadraParteD(16, "Nota: " & Format(RS1!numnotac, "0000000"))
            Printer.Print Lin
            
            Printer.Print " "
            
            Lin = CuadraParteI(40, "Socio: " & Format(RS1!Codsocio, "000000") & " " & RS1!nomsocio)
            Printer.Print Lin
            
            Printer.Print " "
            
            Lin = CuadraParteI(40, "Campo: " & Format(RS1!codcampo, "00000000") & _
                  " " & RS1!nomparti)
            Printer.Print Lin
            
            Lin = CuadraParteI(40, "Variedad: " & Format(RS1!codvarie, "0000") & _
                  " " & RS1!nomvarie)
            Printer.Print Lin
            
            Printer.Print " "
            
            
            '[Monica]16/06/2014: añadido
            
            Lin = CuadraParteI(20, "PESO BRUTO " & Format(RS1!KilosBru, "###,##0"))
            Printer.Print Lin
            
            Printer.Print " "
            
            If DBLet(RS1!numcajo1, "N") <> 0 Then
                Nombre = DevuelveValor("select nomtipen from confenva where codtipen = " & DBSet(RS1!tipocajo1, "N"))
                
                Lin = CuadraParteI(20, "   " & Nombre) & " " & CuadraParteI(20, Format(RS1!numcajo1, "###,##0"))
                Printer.Print Lin
            End If
            If DBLet(RS1!numcajo2, "N") <> 0 Then
                Nombre = DevuelveValor("select nomtipen from confenva where codtipen = " & DBSet(RS1!tipocajo2, "N"))
                
                Lin = CuadraParteI(20, "   " & Nombre) & " " & CuadraParteI(20, Format(RS1!numcajo2, "###,##0"))
                Printer.Print Lin
            End If
            If DBLet(RS1!numcajo3, "N") <> 0 Then
                Nombre = DevuelveValor("select nomtipen from confenva where codtipen = " & DBSet(RS1!tipocajo3, "N"))
                
                Lin = CuadraParteI(20, "   " & Nombre) & " " & CuadraParteI(20, Format(RS1!numcajo3, "###,##0"))
                Printer.Print Lin
            End If
            If DBLet(RS1!numcajo4, "N") <> 0 Then
                Nombre = DevuelveValor("select nomtipen from confenva where codtipen = " & DBSet(RS1!tipocajo4, "N"))
                
                Lin = CuadraParteI(20, "   " & Nombre) & " " & CuadraParteI(20, Format(RS1!numcajo4, "###,##0"))
                Printer.Print Lin
            End If
            If DBLet(RS1!numcajo5, "N") <> 0 Then
                Nombre = DevuelveValor("select nomtipen from confenva where codtipen = " & DBSet(RS1!tipocajo5, "N"))
                
                Lin = CuadraParteI(20, "   " & Nombre) & " " & CuadraParteI(20, Format(RS1!numcajo5, "###,##0"))
                Printer.Print Lin
            End If
            
            Printer.Print " "
            
            cajas = 0
            
            '[Monica]13/06/2014: miramos si es caja en la tabla de envases de confeccion
            If EsCaja(CStr(DBLet(RS1!tipocajo1, "N"))) Then cajas = cajas + DBLet(RS1!numcajo1, "N")
            If EsCaja(CStr(DBLet(RS1!tipocajo2, "N"))) Then cajas = cajas + DBLet(RS1!numcajo2, "N")
            If EsCaja(CStr(DBLet(RS1!tipocajo3, "N"))) Then cajas = cajas + DBLet(RS1!numcajo3, "N")
            If EsCaja(CStr(DBLet(RS1!tipocajo4, "N"))) Then cajas = cajas + DBLet(RS1!numcajo4, "N")
            If EsCaja(CStr(DBLet(RS1!tipocajo5, "N"))) Then cajas = cajas + DBLet(RS1!numcajo5, "N")
            
            Tara = 0
            
            Tara = Tara + DBLet(RS1!taracaja1, "N")
            Tara = Tara + DBLet(RS1!taracaja2, "N")
            Tara = Tara + DBLet(RS1!taracaja3, "N")
            Tara = Tara + DBLet(RS1!taracaja4, "N")
            Tara = Tara + DBLet(RS1!taracaja5, "N")
            Tara = Tara + DBLet(RS1!TaraVehi, "N")
            
'            Lin = CuadraParteI(20, "Cajas: " & cajas) & _
'                  CuadraParteD(20, "Total Tara : " & Format(Tara, "###,##0"))
            Lin = CuadraParteI(20, "Total Tara : " & Format(Tara, "###,##0"))
            Printer.Print Lin
            
'            Lin = CuadraParteI(20, "PESO BRUTO " & Format(RS1!KilosBru, "###,##0")) & _
'                  CuadraParteD(20, "PESO NETO  " & Format(RS1!KilosNet, "###,##0"))

            Printer.Print " "

            Lin = CuadraParteI(20, "PESO NETO  " & Format(RS1!KilosNet, "###,##0"))
            Printer.Print Lin
            
            
            Printer.Print " "
            Printer.Print " "
            Printer.Print " "
            Printer.Print " "
            Printer.Print " "
            Printer.Print " "
            Printer.Print " "
            
            '-- Fin de impresión
            Printer.NewPage
            Printer.EndDoc


'            Dim Puerto As String
'            Dim nFicSalCajon As Integer
'            Puerto = "LPT1"
'            nFicSalCajon = FreeFile
'            Open Puerto For Output As #nFicSalCajon
'                Print #nFicSalCajon, Chr$(27); "i"
'            Close nFicSalCajon
'**********
    Else
        MsgBox "No se ha encontrado la entrada " & CStr(NumTicket) & " de " & Format(FechaTicket, "dd/mm/yyyy"), vbCritical
    End If
    
    RS1.Close
    
    Exit Sub
EImpTickD:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Imprimir ticket."
End Sub


Private Sub ImprimePorLaCom(Cadena As String)
    On Error GoTo EI
    
    Dim nFicSalCajon As Integer
    Dim Puerto As String
    
    'Marzo 2011
    'Puerto = "COM1"
'    Puerto = "COM" & vParamTPV.ComImpresora
    nFicSalCajon = FreeFile
    
    Open Puerto For Output As #nFicSalCajon
    'If Check1.Value = 1 Then
        Print #nFicSalCajon, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
    'Else
    '    Print #nFicSalCajon, Cadena
    'End If
    
    '- corta papel
    '        Print #IMPRESORA, Chr$(29) + Chr$(86) + "0"
    
    Close nFicSalCajon
    
    Exit Sub
EI:
    Cadena = "Error en COM: " & vbCrLf & vbCrLf & Err.Description
    MsgBox Cadena, vbCritical
End Sub

Private Sub CortaPapel()
    Printer.Print Chr(29) & Chr(56) & Chr(49)
'    Printer.EndDoc
End Sub

Private Function LineaCentrada(Lin As String) As String
    Dim queda As Integer
    Dim Parte As Integer
    queda = 40 - Len(Lin)
    Parte = queda / 2
    If Parte Then
        LineaCentrada = String(Parte, " ") & Lin & String(queda - Parte, " ")
    Else
        LineaCentrada = Lin
    End If
End Function

Private Function CuadraParteD(longitud As Integer, Cadena As String) As String
    CuadraParteD = Right(String(longitud, " ") & Cadena, longitud)
End Function

Private Function CuadraParteI(longitud As Integer, Cadena As String) As String
    CuadraParteI = Left(Cadena & String(longitud, " "), longitud)
End Function

