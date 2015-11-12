Attribute VB_Name = "libImpresionDirecta"
Option Explicit



Private Const LineasPorHoja = 48
Private Const MargenIzdo = 0   'Si las pruebas las estoy haciendo o no. Pruebas=6  Real=0
                
                
Private Const ModoImpresion = 2
    ' 0 .- Solo en modo DEBUG. No envia a la impresora
    ' 1 .- Objeto PRINTER
    ' 2 .- Direcatamente sobre LPT
        
    '  Diferencia IMPORTANTE.
    ' SI imprimimos directamente seleccionando la fuente en la impresora hay 36 LINEAS
    ' ni una ni mas ni una menos
    ' Sin embargo con el TPRINTER podemos llegar a las 37 lineas
    ' .....  como suena. ASIN ES!!!!!
        
Dim Cabecera As Collection
Dim Lineas As Collection
Dim Importes As Collection
                    
Dim RS1 As ADODB.Recordset
Dim LasObservaciones As String
Dim NF As Integer
                
    
                
                
                
Private Sub AccionesIniciales()
    
    If ModoImpresion = 1 Then
            Printer.Font = "Courier New"
            Printer.FontSize = "10"
    ElseIf ModoImpresion = 2 Then
        NF = FreeFile
        'Open "d:\t1.txt" For Output As #NF
        Open "LPT1" For Output As #NF
        
        
    End If
    LasObservaciones = ""
End Sub
                
                
                
                
                
                
                
                
'************************************************************
'************************************************************
'
'       Impresion directa. Para facturas, albaranes
'
'
'
'       De momento para 4tonda
'
'           COn lo cual:  El papel es el mismo para todo

Public Sub ImprimirDirectoAlb(cadSelect As String)
    Dim NomImpre As String
  '  Dim FechaT As Date
    Dim rsIVA As ADODB.Recordset
'    Dim vFactu As CFactura
    
    Dim Sql As String
    Dim Lin As String ' línea de impresión
    Dim i As Integer
    
    Dim Producto As String
    Dim Variedad As String
    Dim Partida As String
    Dim Termino As String
    Dim Hdas As Currency
    Dim Has As Currency
    Dim TipoEntrada As String
    Dim SegundaImpresion As Boolean
    Dim Mermas As Long
    Dim Taras As Long
    
    Dim Taras2 As Long
    
    Dim vSocio As CSocio
    
On Error GoTo EImpD
    
        AccionesIniciales
        
        Set RS1 = New ADODB.Recordset
        
        'Cabecera de la entrada
        Sql = "select * from rentradas WHERE " & cadSelect
        RS1.Open Sql, conn, adOpenForwardOnly
        
        Producto = DevuelveValor("select nomprodu from variedades inner join productos on variedades.codprodu = productos.codprodu where codvarie = " & DBSet(RS1!codvarie, "N"))
        Variedad = DevuelveValor("select nomvarie from variedades where codvarie = " & DBSet(RS1!codvarie, "N"))
        Termino = DevuelveValor("select despobla from (rcampos inner join rpartida on rcampos.codparti = rpartida.codparti) inner join rpueblos on rpartida.codpobla = rpueblos.codpobla where codcampo = " & DBSet(RS1!CodCampo, "N"))
        Partida = DevuelveValor("select nomparti from (rcampos inner join rpartida on rcampos.codparti = rpartida.codparti) where codcampo = " & DBSet(RS1!CodCampo, "N"))
        Has = DevuelveValor("select supcoope from rcampos where codcampo = " & DBSet(RS1!CodCampo, "N"))
        Hdas = Round2(Has / vParamAplic.Faneca, 4)
        
        
        Taras = DBLet(RS1!taracaja1, "N") + DBLet(RS1!taracaja2, "N") + DBLet(RS1!taracaja3, "N") + DBLet(RS1!taracaja4, "N") + DBLet(RS1!taracaja5, "N") + DBLet(RS1!TARAVEHISA, "N") + DBLet(RS1!otrastarasa, "N")
        
        '[Monica]16/12/2011: nueva variable de taras
        Taras2 = Taras - (DBLet(RS1!taracajasa1, "N") + DBLet(RS1!taracajasa2, "N") + DBLet(RS1!taracajasa3, "N") + DBLet(RS1!taracajasa4, "N") + DBLet(RS1!taracajasa5, "N"))
        
        Mermas = RS1!KilosBru - Taras + DBLet(RS1!taracajasa1, "N") + DBLet(RS1!taracajasa2, "N") + DBLet(RS1!taracajasa3, "N") + DBLet(RS1!taracajasa4, "N") + DBLet(RS1!taracajasa5, "N") - DBLet(RS1!KilosNet, "N")
        
        If Not EsVariedadGrupo5(CStr(DBLet(RS1!codvarie, "N"))) Then
            Select Case DBLet(RS1!TipoEntr, "N")
                Case 0
                    TipoEntrada = "Normal"
                Case 1
                    TipoEntrada = "V.Campo"
                Case 2
                    TipoEntrada = "P.Integrado"
                Case 3
                    TipoEntrada = "Ind.Directo"
                Case 4
                    TipoEntrada = "Retirada"
                Case 5
                    TipoEntrada = "Venta Directo"
            End Select
        Else
            Select Case DBLet(RS1!TipoEntr, "N")
                Case 0
                    TipoEntrada = "Dalt"
                Case 1
                    TipoEntrada = "V.Campo"
                Case 2
                    TipoEntrada = "P.Integrado"
                Case 3
                    TipoEntrada = "Ind.Directo"
                Case 4
                    TipoEntrada = "Terra"
                Case 5
                    TipoEntrada = "Venta Directo"
            End Select
        
        
        End If
        
        Set Cabecera = New Collection
        
        For i = 1 To 8 '[Monica]16/12/2011: antes 10
            Cabecera.Add " "
        Next i
        
        SegundaImpresion = EsSegundaImpresion(RS1!numnotac)
        
        
        Lin = Space(MargenIzdo) & Left("ALBARAN : " & Format(RS1!numnotac, "0000000") & Space(40), 40)
        'EN la impresora se alineara la linea roja del cabezal con la linea superiror del papel impreso (en verde)
        'Añadairemos una linea en blanco
        
        If SegundaImpresion Then
            Set vSocio = New CSocio
            If vSocio.LeerDatos(RS1!Codsocio) Then
                Lin = Lin & Left("No.Socio     : " & Format(RS1!Codsocio, "000000"), 40)
            End If
        End If
        
        Cabecera.Add Lin
        
        If SegundaImpresion Then
            Lin = Space(MargenIzdo) & Space(40) & Left(vSocio.Nombre & Space(40), 40)
            Cabecera.Add Lin
        Else
            Cabecera.Add " "
        End If
        
        
        Lin = Space(MargenIzdo) & Left("Fecha   : " & Format(RS1!FechaEnt, "dd/mm/yyyy") & Space(40), 40)
        
        If SegundaImpresion Then
            Lin = Lin & Left(vSocio.Direccion & Space(40), 40)
        End If
        Cabecera.Add Lin          '1234567890
        
        Lin = Space(MargenIzdo) & Left("Hora    : " & Format(RS1!horaentr, "hh:mm:ss") & Space(40), 40)
        If SegundaImpresion Then
            Lin = Lin & Left(vSocio.CPostal & "  " & vSocio.Poblacion & Space(40), 40)
        End If
        Cabecera.Add Lin          '1234567890
        
        '[Monica]25/03/2014: quito la linea en blanco pq la necesito para imprimir si tiene o no ausencia de plagas
'        Cabecera.Add " "

        Lin = Space(MargenIzdo) & "Huerto  : " & Format(RS1!CodCampo, "0000000")
        Cabecera.Add Lin          '1234567890
        
        Lin = Space(MargenIzdo) & "Termino : " & Termino
        Cabecera.Add Lin          '1234567890
        
        If SegundaImpresion Then
            Lin = Space(MargenIzdo) & Left("Partida : " & Partida & Space(40), 40)
            Lin = Lin & Left("Hdas.: " & Format(Hdas, "###,##0.00") & Space(40), 40)
            Cabecera.Add Lin          '1234567890
        Else
            Cabecera.Add " "
        End If
        
        '[Monica]25/03/2014: Imprimimos en la entrada si hay o no ausencia de plagas
        Lin = Space(MargenIzdo) & "Ausencia de Plagas: "
        If RS1!ausenciaplagas = 0 Then
            Lin = Lin & "NO"
        Else
            Lin = Lin & "SI"
        End If
        Cabecera.Add Lin
        
        Lin = Space(MargenIzdo) & Left("Producto: " & Producto & Space(40), 40)
        Lin = Lin & Left("Variedad     : " & Variedad, 40)
        Cabecera.Add Lin
        
        Lin = Space(MargenIzdo) & Left("Tipo Ent: " & TipoEntrada & Space(40), 40)
        Cabecera.Add Lin
        
        'Cabecera.Add " "
        
        '[Monica]25/09/2013: cambiamos la siguiente linea por las 2 de abajo
        'Lin = Space(MargenIzdo) & Space(40) & Left("Kilos Brutos : " & Format(RS1!KilosBru, "###,##0") & Space(40), 40)
        Lin = Space(MargenIzdo) & Left("Capataz : " & Format(RS1!codcapat, "0000") & Space(40), 40)
        Lin = Lin & Left("Kilos Brutos : " & Format(RS1!KilosBru, "###,##0") & Space(40), 40)
        Cabecera.Add Lin
        
        Lin = Space(MargenIzdo) & Left("ENVASES ENTRADA      NRO.    TARA" & Space(40), 40)
                                       '123456789012345678901234567890123
        If SegundaImpresion Then
            Lin = Lin & Left("Total Tara   : " & Format(Taras2, "###,##0") & Space(40), 40) '[Monica]16/12/2011: antes taras
        End If
        Cabecera.Add Lin
        
        Lin = Space(MargenIzdo) & Left("---------------------------------" & Space(40), 40)
                                       '123456789012345678901234567890123
        If SegundaImpresion Then
            '[Monica]26/09/2011: las mermas unicamente si no es de grupo de almazara
            If Not EsVariedadGrupo5(CStr(DBLet(RS1!codvarie, "N"))) Then
                Lin = Lin & Left("Total Mermas : " & Format(Mermas, "###,##0") & Space(40), 40)
            End If
        End If
        
        Cabecera.Add Lin
        
        If SegundaImpresion Then
            Lin = Space(MargenIzdo) & Space(40) & Left("KILOS NETOS  : " & Format(RS1!KilosNet, "###,##0") & Space(40), 40)
            Cabecera.Add Lin
        End If
        
        If DBLet(RS1!numcajo1, "N") <> 0 Then
            Lin = Space(MargenIzdo) & Left(vParamAplic.TipoCaja1 & Space(17), 17) & " " & Right(Space(7) & Format(RS1!numcajo1, "###,##0"), 7) & " "
            Lin = Lin & Right(Space(7) & Format(RS1!taracaja1, "###,##0"), 7)
            Cabecera.Add Lin
        End If
        
        If DBLet(RS1!numcajo2, "N") <> 0 Then
            Lin = Space(MargenIzdo) & Left(vParamAplic.TipoCaja2 & Space(17), 17) & " " & Right(Space(7) & Format(RS1!numcajo2, "###,##0"), 7) & " "
            Lin = Lin & Right(Space(7) & Format(RS1!taracaja2, "###,##0"), 7)
            Cabecera.Add Lin
        End If
        
        If DBLet(RS1!numcajo3, "N") <> 0 Then
            Lin = Space(MargenIzdo) & Left(vParamAplic.TipoCaja3 & Space(17), 17) & " " & Right(Space(7) & Format(RS1!numcajo3, "###,##0"), 7) & " "
            Lin = Lin & Right(Space(7) & Format(RS1!taracaja3, "###,##0"), 7)
            Cabecera.Add Lin
        End If
        
        If DBLet(RS1!numcajo4, "N") <> 0 Then
            Lin = Space(MargenIzdo) & Left(vParamAplic.TipoCaja4 & Space(17), 17) & " " & Right(Space(7) & Format(RS1!numcajo4, "###,##0"), 7) & " "
            Lin = Lin & Right(Space(7) & Format(RS1!taracaja4, "###,##0"), 7)
            Cabecera.Add Lin
        End If
        
        If DBLet(RS1!numcajo5, "N") <> 0 Then
            Lin = Space(MargenIzdo) & Left(vParamAplic.TipoCaja5 & Space(17), 17) & " " & Right(Space(7) & Format(RS1!numcajo5, "###,##0"), 7) & " "
            Lin = Lin & Right(Space(7) & Format(RS1!taracaja5, "###,##0"), 7)
            Cabecera.Add Lin
        End If
        
        '[Monica]15/06/2012: Producto certificado globalgap
        If DBLet(RS1!Codtarif) = 1 Then
            Lin = Space(MargenIzdo) & Space(40) & Left("Producto Certificado Globalgap", 40)
            Cabecera.Add Lin
        Else
            Cabecera.Add " "
        End If
        
        If SegundaImpresion Then
            Lin = Space(MargenIzdo) & Left("ENVASES SALIDA       NRO.    TARA" & Space(40), 40)
                                           '123456789012345678901234567890123
            Cabecera.Add Lin
            
            Lin = Space(MargenIzdo) & Left("---------------------------------" & Space(40), 40)
            Cabecera.Add Lin
            
            '[Monica]16/12/2011: las taras de caja de salida se tienen que imprimir
            If DBLet(RS1!numcajosa1, "N") <> 0 Then
                Lin = Space(MargenIzdo) & Left(vParamAplic.TipoCaja1 & Space(17), 17) & " " & Right(Space(7) & Format(RS1!numcajosa1, "###,##0"), 7) & " "
                Lin = Lin & Right(Space(7) & Format(RS1!taracajasa1, "###,##0"), 7)
                Cabecera.Add Lin
            End If
            If DBLet(RS1!numcajosa2, "N") <> 0 Then
                Lin = Space(MargenIzdo) & Left(vParamAplic.TipoCaja2 & Space(17), 17) & " " & Right(Space(7) & Format(RS1!numcajosa2, "###,##0"), 7) & " "
                Lin = Lin & Right(Space(7) & Format(RS1!taracajasa2, "###,##0"), 7)
                Cabecera.Add Lin
            End If
            If DBLet(RS1!numcajosa3, "N") <> 0 Then
                Lin = Space(MargenIzdo) & Left(vParamAplic.TipoCaja3 & Space(17), 17) & " " & Right(Space(7) & Format(RS1!numcajosa3, "###,##0"), 7) & " "
                Lin = Lin & Right(Space(7) & Format(RS1!taracajasa3, "###,##0"), 7)
                Cabecera.Add Lin
            End If
            If DBLet(RS1!numcajosa4, "N") <> 0 Then
                Lin = Space(MargenIzdo) & Left(vParamAplic.TipoCaja4 & Space(17), 17) & " " & Right(Space(7) & Format(RS1!numcajosa4, "###,##0"), 7) & " "
                Lin = Lin & Right(Space(7) & Format(RS1!taracajasa4, "###,##0"), 7)
                Cabecera.Add Lin
            End If
            If DBLet(RS1!numcajosa5, "N") <> 0 Then
                Lin = Space(MargenIzdo) & Left(vParamAplic.TipoCaja5 & Space(17), 17) & " " & Right(Space(7) & Format(RS1!numcajosa5, "###,##0"), 7) & " "
                Lin = Lin & Right(Space(7) & Format(RS1!taracajasa5, "###,##0"), 7)
                Cabecera.Add Lin
            End If
            '16/12/2011: hasta aqui
                        
            Cabecera.Add " "
        
            If DBLet(RS1!TARAVEHISA, "N") <> 0 Then
                Lin = Space(MargenIzdo) & Left("Tara Vehiculo" & Space(17), 17) & " " & Right(Space(7), 7) & " "
                Lin = Lin & Right(Space(7) & Format(RS1!TARAVEHISA, "###,##0"), 7)
                Cabecera.Add Lin
            End If
            If DBLet(RS1!otrastarasa, "N") <> 0 Then
                Lin = Space(MargenIzdo) & Left("Otras Taras" & Space(17), 17) & " " & Right(Space(7), 7) & " "
                Lin = Lin & Right(Space(7) & Format(RS1!otrastarasa, "###,##0"), 7)
                Cabecera.Add Lin
            End If
        
            If DBLet(RS1!TARAVEHISA, "N") <> 0 Or DBLet(RS1!otrastarasa, "N") Or _
               DBLet(RS1!numcajosa1, "N") <> 0 Or DBLet(RS1!numcajosa2, "N") <> 0 Or DBLet(RS1!numcajosa3, "N") <> 0 Or DBLet(RS1!numcajosa4, "N") <> 0 Or DBLet(RS1!numcajosa5, "N") <> 0 Then
               '[Monica]16/12/2011: añadida la segunda linea de numcajosa1..5 <> 0
                Lin = Space(MargenIzdo) & Left(Space(17), 17) & " " & Right(Space(7), 7) & " "
                Lin = Lin & Right(Space(7) & "-------", 7)
                Cabecera.Add Lin
            
                Lin = Space(MargenIzdo) & Left("TOTAL TARAS" & Space(17), 17) & " " & Right(Space(7), 7) & " "
'[Monica]16/12/2011: Lin = Lin & Right(Space(7) & Format(Taras, "###,##0"), 7) antes era esto
                Lin = Lin & Right(Space(7) & Format(Taras2, "###,##0"), 7)
                Cabecera.Add Lin
            End If
        End If
        
        
        
        'Cerramos el rs
        RS1.Close
        Set RS1 = Nothing
        
        
        
        'Ya tenemos todos los datos
        'Ahora manadmos a la impresora
        ImprimeEnPapel
        
        
        
EImpD:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Imprimir directo."
        Err.Clear
    End If
    
    
    Set Cabecera = Nothing
    Set Lineas = Nothing
    Set Importes = Nothing
    Set rsIVA = New ADODB.Recordset
    Set RS1 = New ADODB.Recordset
    Exit Sub
    
End Sub
        
Private Sub AjusteLineasImportes()
    'Linea en blaco deonde van los cuadrados de BImpo, porceta....
    Set Importes = New Collection
    Importes.Add " "

    If ModoImpresion = 2 Then
        'SOlo tiro uno p'abajo
    Else
        Importes.Add " "
    End If
End Sub


Private Sub ImprimeEnPapel()
    Dim i As Integer
    Dim J As Integer
    Dim PagActual As Integer
    Dim Lin As String
    Dim impor As Currency
    Dim NumeroPaginas As Integer
        'AHORA IMPRIMIMOS.
        'TEnemos cargada las lineas
        NumeroPaginas = ((Cabecera.Count - 1) \ LineasPorHoja) + 1
        i = 0
        PagActual = 1
        For J = 1 To Cabecera.Count
            
'            If i = 0 Then
'                '***********************************************************
'                'Imprimir cabecera
'                For i = 1 To Cabecera.Count
'                    ImprimeLaLinea Cabecera(i)
'                Next i
'                i = 0
'                'Si hay mas de una hoja pongo tambien el numero de hoja
'                If NumeroPaginas > 1 Then
'                    Lin = Space(MargenIzdo + 45) & "Pag: " & PagActual & " / " & NumeroPaginas
'                    ImprimeLaLinea Lin
'                Else
'                    ImprimeLaLinea " "
'                End If
'                ImprimeLaLinea " "
'                ImprimeLaLinea " "
'
'                PagActual = PagActual + 1
'            End If
            
            ImprimeLaLinea Cabecera(J)
            i = i + 1
            
'            'Si es la ultima linea NO hacemos nada
'            If J < Cabecera.Count Then
'                If i = LineasPorHoja Then
'                    ImprimeLaLinea " ": ImprimeLaLinea " ":
'                    If ModoImpresion = 1 Then ImprimeLaLinea " "
'                    ImprimeLaLinea Space(50) & "** ** **" 'los importes
'                    'Linea en blaco deonde van los cauadrados de BImpo, porceta....
'                    ' y las lineas finales
'                    'Ha rellenado todas. Si hay mas lineas que imprimir entonces
'
'
'                    For i = 1 To 5
'                        ImprimeLaLinea " "
'                    Next i
'
'                    i = 0
'                End If
'            End If
        
        Next
        
        
        'Para situar el cabezal en la impresion
        If i < LineasPorHoja And i <> 0 Then
            'Ha impreso i lineas
            'Hasta las 10 que caben...
            i = LineasPorHoja - i
            While i > 0
                ImprimeLaLinea ""
                i = i - 1
            Wend
            
        End If
        
'        'Los importes
'        For J = 1 To Importes.Count
'            ImprimeLaLinea Importes.item(J)
'        Next
        
        'Final hoja
        '--------------------
        If ModoImpresion = 1 Then
            Printer.EndDoc
        Else
            If ModoImpresion = 2 Then
                'Re situo el papel donde le toca

                For J = 1 To 3
                    ImprimeLaLinea " "
                Next
            
            
                Close (NF)
            End If
        End If
        
    'Volver la impresora a la predeterminada
    'EstablecerImpresora NomImpre
    
End Sub


Private Function LineaImportes(BaseIva As Currency, PorceIVA As Currency, ImpIVA As Currency, IvaRE As Currency, ImpIVARE As Currency, TotalFac As String) As String
Dim Lin As String
    
        Lin = Space(17) & Format(BaseIva, FormatoImporte)
        Lin = Right(Lin, 17) '17 es la longiyud de bas imponible
        Lin = Space(MargenIzdo + 16) & Lin
        Lin = Lin & "  " & Right(Space(5) & Format(PorceIVA, FormatoPorcen), 5)
         Lin = Lin & " "
        Lin = Lin & Right(Space(11) & Format(ImpIVA, FormatoImporte), 11)
        If IvaRE = 0 Then
            'No lleva % retencion
            Lin = Lin & Space(17)
        Else
            'SI LLEVA
            
        End If
        
        LineaImportes = Lin & Right(Space(16) & TotalFac, 16)
        
        
End Function


''Como los campos del albaran y de la factura son los mismos...
'' Paso Opcion por si acaso tengo que hacer algo a las facturas o a los albaranes...
'Private Sub CargaEncabezado2(Opcion As Byte, ByRef Rs As ADODB.Recordset)
'Dim L As String
'        L = Space(35) & Format(Rs!CodClien, "000") & Space(15)
'        L = Mid(L, 1, (MargenIzdo + 45)) & Rs!NomClien
'        'linea 4
'        Cabecera.Add L
'        Cabecera.Add Space(MargenIzdo + 45) & DBLet(Rs!domclien, "T")
'        Cabecera.Add Space(MargenIzdo + 45) & Rs!pobclien
'        Cabecera.Add Space(MargenIzdo + 45) & Format(Rs!codpobla, "00000") & " " & Rs!proclien
'        Cabecera.Add Space(MargenIzdo + 45) & "C.I.F.: " & Rs!nifClien
'        L = Space(MargenIzdo + 2) & vEmpresa.nomempre & Space(40)
'        L = Mid(L, 1, MargenIzdo + 45) & "Forma de pago: " & DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "Codforpa", Rs!codforpa)
'        Cabecera.Add L
'        Cabecera.Add Space(MargenIzdo + 2) & vParam.DomicilioEmpresa
'        L = Space(MargenIzdo + 2) & vParam.CPostal & " " & vParam.Poblacion & " " & vParam.Provincia
'        Cabecera.Add L
'        L = Space(MargenIzdo + 2) & "Tfno: " & vParam.Telefono & " " & vParam.CifEmpresa
'        Cabecera.Add L
'
'End Sub

Private Sub ImprimeLaLinea(Linea As String)
    Debug.Print Linea
    If ModoImpresion = 0 Then Exit Sub  'Solo debug
    If ModoImpresion = 1 Then
        Printer.Print Linea
    Else
        Print #NF, Linea
    End If
    
End Sub


Private Function EsSegundaImpresion(Nota As Long) As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String

    Sql = "select * from rentradas where numnotac = " & DBSet(Nota, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    EsSegundaImpresion = False
    
    If Not RS.EOF Then
        EsSegundaImpresion = Not (IsNull(RS!taracajasa1) And IsNull(RS!taracajasa2) And IsNull(RS!taracajasa2) And IsNull(RS!taracajasa4) And IsNull(RS!taracajasa5) And IsNull(RS!TARAVEHISA) And IsNull(RS!otrastarasa))
    End If

    Set RS = Nothing

End Function


'
'
'
''------------------------------------------------------
'' FACTURAS TPV
'
'Public Sub ImprimirDirectoFact(cadSelect As String)
'    Dim NomImpre As String
'  '  Dim FechaT As Date
'
'    Dim rsIVA As ADODB.Recordset
'    Dim vFactu As CFactura
'
'    Dim Sql As String
'    Dim Lin As String ' línea de impresión
'    Dim TieneObsAlbaran As Integer
'
'
'
'On Error GoTo EImpD
'
'    'Establecemos la impresora de ticket
''    If vParamTPV.NomImpresora <> "" Then
''        If Printer.DeviceName <> vParamTPV.NomImpresora Then
''            'guardamos la impresora que habia
''            NomImpre = Printer.DeviceName
''            'establecemos la de ticket
''            EstablecerImpresora vParamTPV.NomImpresora
''        End If
''    End If
'
'
'        AccionesIniciales
'
'        Set rs1 = New ADODB.Recordset
'
'
'
'
'
'
'        'Lineas de albaranes
'        'SQL:
'
'        'Guardo las obseraviaciones
'        Sql = " FROM scafac INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom AND "
'        Sql = Sql & " scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu "
'        Sql = Sql & " WHERE " & cadSelect
'
'        rs1.Open "Select observa1,scafac1.numalbar " & Sql, conn, adOpenForwardOnly
'        TieneObsAlbaran = 0
'        While Not rs1.EOF
'            Sql = DBLet(rs1!observa1, "T")
'            Lin = "[" & Format(rs1!NumAlbar, "000000") & "]"
'            rs1.MoveNext
'            If Not rs1.EOF Then TieneObsAlbaran = 1 'Para que pinte el numero de albaran
'            If Sql <> "" Then
'                If TieneObsAlbaran = 1 Then Sql = Lin & "   " & Sql
'                LasObservaciones = LasObservaciones & "- " & Sql & "|"
'            End If
'        Wend
'        rs1.Close
'
'
'
'
'        'El select para las lineas de albaran
'        Sql = " FROM ((scafac INNER JOIN scafac1 ON ((scafac.codtipom=scafac1.codtipom) AND "
'        Sql = Sql & " (scafac.numfactu=scafac1.numfactu)) AND (scafac.fecfactu=scafac1.fecfactu)) "
'        Sql = Sql & " INNER JOIN slifac ON ((((scafac1.codtipom=slifac.codtipom) AND "
'        Sql = Sql & " (scafac1.numfactu=slifac.numfactu)) AND (scafac1.fecfactu=slifac.fecfactu)) AND "
'        Sql = Sql & " (scafac1.codtipoa=slifac.codtipoa)) AND (scafac1.numalbar=slifac.numalbar)) "
'        Sql = Sql & " INNER JOIN sartic ON slifac.codartic=sartic.codartic"
'        'Y el albaran
'        Sql = Sql & " AND " & cadSelect
'
'
'
'
'        'Tipos de IVA
'        Set rsIVA = New ADODB.Recordset
'        rsIVA.Open "Select * from tiposiva", ConnConta, adOpenKeyset, adLockPessimistic, adCmdText
'
'
'
'
'        'Cabecera del albaran
'        Lin = "select * from scafac WHERE " & cadSelect
'        rs1.Open Lin, conn, adOpenForwardOnly
'
'
'        Lin = Space(MargenIzdo + 45) & "FAC.   " & rs1!CodTipom & Format(rs1!numfactu, "000000") & Space(12) & Format(rs1!fecfactu, "dd/mm/yyyy")
'        Set Cabecera = New Collection
'        'EN la impresora se alineara la linea roja del cabezal con la linea superiror del papel impreso (en verde)
'        'Añadairemos una linea en blanco
'        Cabecera.Add " "
'        Cabecera.Add Lin
'        Cabecera.Add Space(MargenIzdo + 45)
'
'        'Lineas 2 a 7 , datos cliente  nomclien  domclien  codpobla  pobclien  proclien  nifclien
'        CargaEncabezado2 1, rs1
'
'
'        'Leo estos valores para el final del albaran dtoppago dtognral
'        Set vFactu = New CFactura
'        vFactu.DtoPPago = rs1!DtoPPago
'        vFactu.DtoGnral = rs1!DtoGnral
'        vFactu.Cliente = rs1!CodClien
'        vFactu.numfactu = rs1!numfactu
'        vFactu.fecfactu = rs1!fecfactu
'        vFactu.CodTipom = rs1!CodTipom
'
'        'Cerramos el rs
'        rs1.Close
'
'
'
'        Lin = "select slifac.*,codigiva,numserie " & Sql
'        Lin = Lin & " ORDER BY numalbar,numlinea"
'        rs1.Open Lin, conn, adOpenForwardOnly
'
'
'        Set Lineas = New Collection
'        While Not rs1.EOF
'
'            'Las lineas correspondientes
'            Lin = Right(Space(16) & rs1!codArtic, 16)  '16 es la longiyud de codartic
'            Lin = Space(MargenIzdo) & Lin
'            Lin = Lin & " " & Left(rs1!NomArtic & Space(30), 30)
'
'            Lin = Lin & Right(Space(9) & Format(rs1!cantidad, FormatoCantidad), 9) & Space(2)
'            Lin = Lin & Right(Space(10) & Format(rs1!precioar, FormatoPrecio), 10)
'            'El IVA.
'            rsIVA.Find "codigiva = " & rs1!Codigiva, , adSearchForward, 1
'            If rsIVA.EOF Then
'                Lin = Lin & " * "
'            Else
'                Lin = Lin & " " & Format(rsIVA!PorceIVA, "00")
'            End If
'            Lin = Lin & Right(Space(15) & Format(rs1!ImporteL, FormatoPrecio), 15)
'            Lineas.Add Lin
'            'El numero de serie
'            Lin = DBLet(rs1!numserie, "T")
'            If Lin <> "" Then
'                Lin = Space(14) & " N. Reg: " & Space(12) & Lin
'                Lineas.Add Lin
'            End If
'            rs1.MoveNext
'
'
'        Wend
'        rs1.Close
'        rsIVA.Close
'
'        'Las observaciones de la factura
'        'Las tenemos cargadas, empipadas, en LasObservaciones
'        If LasObservaciones <> "" Then
'            Lineas.Add " "  'Un espacio en blanco
'            Lineas.Add Space(MargenIzdo) & "Observaciones"
'            While LasObservaciones <> ""
'                TieneObsAlbaran = InStr(1, LasObservaciones, "|")
'                If TieneObsAlbaran = 0 Then
'                    LasObservaciones = ""
'                Else
'                    Sql = Mid(LasObservaciones, 1, TieneObsAlbaran - 1)
'                    LasObservaciones = Mid(LasObservaciones, TieneObsAlbaran + 1)
'                    Lineas.Add Space(MargenIzdo) & Sql
'                End If
'            Wend
'        End If
'
'        'Los importes. Los cargo desde la factura
'        If Not CargarImportesDesdeFactura(vFactu, Lin) Then
'            If Not vFactu.CalcularDatosFactura(cadSelect, "scafac", "slifac", False) Then
'                MsgBox "Importes factura NO encontrados NI calculados", vbExclamation
'            Else
'                MsgBox "Importes factura NO encontrados. Se han calculado para la impresion", vbExclamation
'            End If
'        End If
'
'        'TRozo final de los importes
'        AjusteLineasImportes
'        'Linea uno. SEGURO QUE LA IMPRIME
'        '--------------------------------
'
'        'Voy a cargar todos los datos de  importes de la factura
'
'        Lin = Format(vFactu.TotalFac, FormatoImporte)
'
'
'        Lin = LineaImportes(vFactu.BaseIVA1, vFactu.PorceIVA1, vFactu.ImpIVA1, vFactu.PorceIVA1RE, vFactu.ImpIVA1RE, Lin)
'        Importes.Add Lin
'
'        If vFactu.BaseIVA2 <> 0 Then
'            Lin = LineaImportes(vFactu.BaseIVA2, vFactu.PorceIVA2, vFactu.ImpIVA2, vFactu.PorceIVA2RE, vFactu.ImpIVA2RE, "")
'        Else
'            Lin = ""
'        End If
'        Importes.Add Lin
'
'        If vFactu.BaseIVA3 <> 0 Then
'            Lin = LineaImportes(vFactu.BaseIVA3, vFactu.PorceIVA3, vFactu.ImpIVA3, vFactu.PorceIVA3RE, vFactu.ImpIVA3RE, "")
'        Else
'            Lin = ""
'        End If
'        Importes.Add Lin
'
'
'
'        'Ya tenemos todos los datos
'        'Ahora manadmos a la impresora
'        ImprimeEnPapel
'
'
'
'EImpD:
'    If Err.Number <> 0 Then
'        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Imprimir directo."
'        Err.Clear
'    End If
'
'
'    Set Cabecera = Nothing
'    Set Lineas = Nothing
'    Set Importes = Nothing
'    Set rsIVA = New ADODB.Recordset
'    Set rs1 = New ADODB.Recordset
'    Exit Sub
'
'End Sub
'
'
'
'
'
''------------------------------------------------------
'' REimpresion de facturas. Pone lo del albaran y eso
'
'
'
'Public Sub ReImprimirDirectoFact(cadSelect As String)
'
'  '  Dim FechaT As Date
'
'    Dim vFactu As CFactura
'    Dim Grupo As String
'    Dim Sql As String
'    Dim Lin As String ' línea de impresión
'    Dim I As Integer
'    Dim NumeroPaginas  As Integer
'    Dim Importe As Currency
'    Dim Albaran As String
'On Error GoTo EImpD
'
'
'
'
'
'
'        Set rs1 = New ADODB.Recordset
'
'        AccionesIniciales
'
'
'
'        'Cogeremos. los albaranes de las facturas y los articulos que tengan nºregistro
'        'SQL:
'        Sql = "Select scafac.*,slifac.*,CodTraba,FechaAlb,numSerie"
'        Sql = Sql & " FROM ((scafac INNER JOIN scafac1 ON ((scafac.codtipom=scafac1.codtipom) AND "
'        Sql = Sql & " (scafac.numfactu=scafac1.numfactu)) AND (scafac.fecfactu=scafac1.fecfactu)) "
'        Sql = Sql & " INNER JOIN slifac ON ((((scafac1.codtipom=slifac.codtipom) AND "
'        Sql = Sql & " (scafac1.numfactu=slifac.numfactu)) AND (scafac1.fecfactu=slifac.fecfactu)) AND "
'        Sql = Sql & " (scafac1.codtipoa=slifac.codtipoa)) AND (scafac1.numalbar=slifac.numalbar)) "
'        Sql = Sql & " INNER JOIN sartic ON slifac.codartic=sartic.codartic"
'
'        'Y el albaran
'        Sql = Sql & " AND " & cadSelect
'
'        rs1.Open Sql, conn, adOpenForwardOnly
'
'
'        Lin = Space(MargenIzdo + 45) & "FAC.   " & rs1!CodTipom & Format(rs1!numfactu, "000000") & Space(12) & Format(rs1!fecfactu, "dd/mm/yyyy")
'        Set Cabecera = New Collection
'        'EN la impresora se alineara la linea roja del cabezal con la linea superiror del papel impreso (en verde)
'        'Añadairemos una linea en blanco
'        Cabecera.Add " "
'        Cabecera.Add Lin
'        Cabecera.Add Space(MargenIzdo + 45)
'
'        'Lineas 2 a 7 , datos cliente  nomclien  domclien  codpobla  pobclien  proclien  nifclien
'        CargaEncabezado2 1, rs1
'
'
'        'Leo estos valores para el final del albaran dtoppago dtognral
'        Set vFactu = New CFactura
'        vFactu.DtoPPago = rs1!DtoPPago
'        vFactu.DtoGnral = rs1!DtoGnral
'        vFactu.Cliente = rs1!CodClien
'        vFactu.numfactu = rs1!numfactu
'        vFactu.fecfactu = rs1!fecfactu
'        vFactu.CodTipom = rs1!CodTipom
'
'
'        'En sql tendremos los numeros de lote
'        Sql = ""
'        Grupo = ""
'        'vamos imprimiendo los albaranes
'        Set Lineas = New Collection
'        I = 0
'        While Not rs1.EOF
'            Lin = rs1!codTipoa & Format(rs1!NumAlbar, "0000000")
'            If Lin <> Grupo Then
'                If Grupo <> "" Then LineaAlbaranFactura Albaran, Importe, Sql, I
'
'
'                Grupo = Lin
'                Lin = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", rs1!CodTraba)
'                If Lin <> "" Then Lin = " Venta realizada por " & Lin
'                Albaran = "Albarán: " & Grupo & " de fecha " & Format(rs1!FechaAlb, "dd/mm/yyyy") & " " & Lin
'                'Faltara añadir el importe
'                Importe = 0
'
'                Sql = "|" 'Llevaremos los nº de lote en este albaran
'
'            End If
'            'El numero de serie
'            Lin = DBLet(rs1!numserie, "T")
'            If Lin <> "" Then
'                If InStr(1, Sql, "|" & Lin & "|") = 0 Then Sql = Sql & Lin & "|"
'
'            End If
'            Importe = Importe + rs1!ImporteL
'            rs1.MoveNext
'        Wend
'        rs1.Close
'        LineaAlbaranFactura Albaran, Importe, Sql, I
'
'
'
'        'Los importes. Los cargo desde la factura
'        If Not CargarImportesDesdeFactura(vFactu, Lin) Then
'            If Not vFactu.CalcularDatosFactura(cadSelect, "scafac", "slifac", False) Then
'                MsgBox "Importes factura NO encontrados NI calculados", vbExclamation
'            Else
'                MsgBox "Importes factura NO encontrados. Se han calculado para la impresion", vbExclamation
'            End If
'        End If
'
'
'        'TRozo final de los importes
'        AjusteLineasImportes
'
'        'Linea uno. SEGURO QUE LA IMPRIME
'        '--------------------------------
'        'Campo BAse imponible. Empieza hasta el 41, si alineamos a la derecha
'        Lin = Format(vFactu.TotalFac, FormatoImporte)
'        Lin = LineaImportes(vFactu.BaseIVA1, vFactu.PorceIVA1, vFactu.ImpIVA1, vFactu.PorceIVA1RE, vFactu.ImpIVA1RE, Lin)
'        Importes.Add Lin
'
'        If vFactu.BaseIVA2 <> 0 Then
'            Lin = LineaImportes(vFactu.BaseIVA2, vFactu.PorceIVA2, vFactu.ImpIVA2, vFactu.PorceIVA2RE, vFactu.ImpIVA2RE, "")
'        Else
'            Lin = ""
'        End If
'        Importes.Add Lin
'
'        If vFactu.BaseIVA3 <> 0 Then
'            Lin = LineaImportes(vFactu.BaseIVA3, vFactu.PorceIVA3, vFactu.ImpIVA3, vFactu.PorceIVA3RE, vFactu.ImpIVA3RE, "")
'        Else
'            Lin = ""
'        End If
'        Importes.Add Lin
'
'
'
'        'Ya tenemos todos los datos
'        'Ahora manadmos a la impresora
'        'NumeroPaginas = ((i - 1) \ LineasPorHoja) + 1
'        'If I > 13 Then Stop
'        ImprimeEnPapel
'
'
'
'EImpD:
'    If Err.Number <> 0 Then
'        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Imprimir directo."
'        Err.Clear
'    End If
'
'
'    Set Cabecera = Nothing
'    Set Lineas = Nothing
'    Set Importes = Nothing
'    Set rs1 = New ADODB.Recordset
'    Exit Sub
'
'End Sub
'
'
'Private Sub LineaAlbaranFactura(L As String, Importe As Currency, ArticulosConNumeroSerie As String, ByRef ContadorDeLineas As Integer)
'Dim I As Integer
'        L = Space(MargenIzdo) & L & Space(30)
'        L = Mid(L, 1, 78)
'        L = L & Right(Space(15) & Format(Importe, FormatoImporte), 15)
'        Lineas.Add L
'        ContadorDeLineas = ContadorDeLineas + 1
'
'        If ArticulosConNumeroSerie <> "|" Then
'            ArticulosConNumeroSerie = Mid(ArticulosConNumeroSerie, 2)
'            I = 1
'            Lineas.Add ""
'            ContadorDeLineas = ContadorDeLineas + 1
'
'            While I <> 0
'                I = InStr(1, ArticulosConNumeroSerie, "|")
'                If I > 0 Then
'                    L = Mid(ArticulosConNumeroSerie, 1, I - 1)
'                    ArticulosConNumeroSerie = Mid(ArticulosConNumeroSerie, I + 1)
'                    L = Space(14) & " N. Reg: " & Space(12) & L
'                    Lineas.Add L
'                    ContadorDeLineas = ContadorDeLineas + 1
'                End If
'            Wend
'        End If
'End Sub
'
'
'Private Function CargarImportesDesdeFactura(ByRef F As CFactura, ByRef auxiliar As String) As Boolean
'    CargarImportesDesdeFactura = False
'    auxiliar = "Select * from scafac where codtipom=" & DBSet(F.CodTipom, "T")
'    auxiliar = auxiliar & " AND numfactu=" & DBSet(F.numfactu, "N")
'    auxiliar = auxiliar & " AND fecfactu=" & DBSet(F.fecfactu, "F")
'    rs1.Open auxiliar, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If rs1.EOF Then
'
'
'
'    Else
'        CargarImportesDesdeFactura = True
'
'        F.BaseIVA1 = DBLet(rs1!baseimp1, "N")
'        F.PorceIVA1 = DBLet(rs1!porciva1, "N")
'        F.ImpIVA1 = DBLet(rs1!imporiv1, "N")
'        F.PorceIVA1RE = DBLet(rs1!porciva1re, "N")
'        F.ImpIVA1RE = DBLet(rs1!imporiv1re, "N")
'
'
'
'        F.BaseIVA2 = DBLet(rs1!baseimp2, "N")
'        F.PorceIVA2 = DBLet(rs1!porciva2, "N")
'        F.ImpIVA2 = DBLet(rs1!imporiv2, "N")
'        F.PorceIVA2RE = DBLet(rs1!porciva2re, "N")
'        F.ImpIVA2RE = DBLet(rs1!imporiv2re, "N")
'
'        F.BaseIVA3 = DBLet(rs1!baseimp3, "N")
'        F.PorceIVA3 = DBLet(rs1!porciva3, "N")
'        F.ImpIVA3 = DBLet(rs1!imporiv3, "N")
'        F.PorceIVA3RE = DBLet(rs1!porciva3re, "N")
'        F.ImpIVA3RE = DBLet(rs1!imporiv3re, "N")
'
'        F.TotalFac = rs1!TotalFac
'
'
'    End If
'    rs1.Close
'End Function
'


'************************************************************
'************************************************************
'
'       Impresion directa. Para albaranes de bodega
'
'
'
'       De momento para 4tonda
'
'           COn lo cual:  El papel es el mismo para todo

Public Sub ImprimirDirectoAlbBodega(cadSelect As String)
    Dim NomImpre As String
  '  Dim FechaT As Date
    Dim rsIVA As ADODB.Recordset
'    Dim vFactu As CFactura
    
    Dim Sql As String
    Dim Lin As String ' línea de impresión
    Dim i As Integer
    
    Dim Producto As String
    Dim Variedad As String
    Dim Partida As String
    Dim Termino As String
    Dim Hdas As Currency
    Dim Has As Currency
    Dim TipoEntrada As String
    Dim SegundaImpresion As Boolean
    Dim Mermas As Long
    Dim Taras As Long
    
    Dim vSocio As CSocio
    
On Error GoTo EImpD
    
        AccionesIniciales
        
        Set RS1 = New ADODB.Recordset
        
        'Cabecera de la entrada
        Sql = "select rhisfruta.*, rhisfruta_entradas.fechaent, rhisfruta_entradas.horaentr, rhisfruta_entradas.observac from rhisfruta inner join rhisfruta_entradas on rhisfruta.numalbar = rhisfruta_entradas.numalbar WHERE " & cadSelect
        RS1.Open Sql, conn, adOpenForwardOnly
        
        Producto = DevuelveValor("select nomprodu from variedades inner join productos on variedades.codprodu = productos.codprodu where codvarie = " & DBSet(RS1!codvarie, "N"))
        Variedad = DevuelveValor("select nomvarie from variedades where codvarie = " & DBSet(RS1!codvarie, "N"))
        Termino = DevuelveValor("select despobla from (rcampos inner join rpartida on rcampos.codparti = rpartida.codparti) inner join rpueblos on rpartida.codpobla = rpueblos.codpobla where codcampo = " & DBSet(RS1!CodCampo, "N"))
        Partida = DevuelveValor("select nomparti from (rcampos inner join rpartida on rcampos.codparti = rpartida.codparti) where codcampo = " & DBSet(RS1!CodCampo, "N"))
        Has = DevuelveValor("select supcoope from rcampos where codcampo = " & DBSet(RS1!CodCampo, "N"))
        Hdas = Round2(Has / vParamAplic.Faneca, 4)
        
        
        Taras = DBLet(RS1!tarabodega, "N")
        
        Mermas = 0
        
        Select Case DBLet(RS1!TipoEntr, "N")
            Case 0
                TipoEntrada = "Normal"
            Case 1
                TipoEntrada = "V.Campo"
            Case 2
                TipoEntrada = "P.Integrado"
            Case 3
                TipoEntrada = "Ind.Directo"
            Case 4
                TipoEntrada = "Retirada"
            Case 5
                TipoEntrada = "Venta Directo"
        End Select
        
        
        Set Cabecera = New Collection
        
        For i = 1 To 10
            Cabecera.Add " "
        Next i
        
        
        
        Lin = Space(MargenIzdo) & Left("ALBARAN :  " & Format(RS1!numalbar, "0000000") & Space(40), 40)
        'EN la impresora se alineara la linea roja del cabezal con la linea superiror del papel impreso (en verde)
        'Añadairemos una linea en blanco
        
        Set vSocio = New CSocio
        If vSocio.LeerDatos(RS1!Codsocio) Then
            Lin = Lin & Left("No.Socio     : " & Format(RS1!Codsocio, "000000"), 40)
        End If
        
        Cabecera.Add Lin
        
        Lin = Space(MargenIzdo) & Space(40) & Left(vSocio.Nombre & Space(40), 40)
        Cabecera.Add Lin
        
        
        Lin = Space(MargenIzdo) & Left("Fecha   :  " & Format(RS1!FechaEnt, "dd/mm/yyyy") & Space(40), 40)
        
        Lin = Lin & Left(vSocio.Direccion & Space(40), 40)
        Cabecera.Add Lin          '1234567890
        
        Lin = Space(MargenIzdo) & Left("Hora    :  " & Format(RS1!horaentr, "hh:mm:ss") & Space(40), 40)
        Lin = Lin & Left(vSocio.CPostal & "  " & vSocio.Poblacion & Space(40), 40)
        Cabecera.Add Lin          '1234567890
        
        Cabecera.Add " "
        Lin = Space(MargenIzdo) & "Huerto  : " & Format(RS1!CodCampo, "0000000")
        Cabecera.Add Lin          '1234567890
        
        Lin = Space(MargenIzdo) & "Termino : " & Termino
        Cabecera.Add Lin          '1234567890
        
        Lin = Space(MargenIzdo) & Left("Partida : " & Partida & Space(40), 40)
        Lin = Lin & Left("Hdas.: " & Format(Hdas, "###,##0.00") & Space(40), 40)
        Cabecera.Add Lin          '1234567890
        
        Lin = Space(MargenIzdo) & Left("Producto: " & Producto & Space(40), 40)
        Lin = Lin & Left("Variedad     : " & Variedad, 40)
        Cabecera.Add Lin
        
'        Lin = Space(MargenIzdo) & Left("Tipo Ent: " & TipoEntrada & Space(40), 40)

        'grados
        Lin = Space(MargenIzdo) & Space(40) & Left("Grados       : " & Format(DBLet(RS1!PrEstimado, "N"), "###0.00") & Space(40), 40)
        Cabecera.Add Lin
        
        'tolva
        Lin = Space(MargenIzdo) & Space(40) & Left("Num.Tolva    : " & Format(DBLet(RS1!tolva, "N") + 1, "######0") & Space(40), 40)
        Cabecera.Add Lin
        
        'Kilos brutos
        'Cabecera.Add " "
        Lin = Space(MargenIzdo) & Space(40) & Left("Kilos Brutos : " & Format(RS1!KilosBru, "###,##0") & Space(40), 40)
        Cabecera.Add Lin
        
        Lin = Space(MargenIzdo) & Left("ENVASES ENTRADA      NRO.    TARA" & Space(40), 40)
                                       '123456789012345678901234567890123
        Lin = Lin & Left("Total Tara   : " & Format(Taras, "###,##0") & Space(40), 40)
        Cabecera.Add Lin
        
        Lin = Space(MargenIzdo) & Left("---------------------------------" & Space(40), 40)
                                       '123456789012345678901234567890123
'        Lin = Lin & Left("Total Mermas : " & Format(Mermas, "###,##0") & Space(40), 40)
        
        Cabecera.Add Lin
        
        Lin = Space(MargenIzdo) & Space(40) & Left("KILOS NETOS  : " & Format(RS1!KilosNet, "###,##0") & Space(40), 40)
        Cabecera.Add Lin
        
        
        Cabecera.Add " "
        
        Lin = Space(MargenIzdo) & Left("ENVASES SALIDA       NRO.    TARA" & Space(40), 40)
                                       '123456789012345678901234567890123
        Cabecera.Add Lin
        
        Lin = Space(MargenIzdo) & Left("---------------------------------" & Space(40), 40)
        Cabecera.Add Lin
    
        Cabecera.Add " "
    
        
        Lin = Space(MargenIzdo) & Left("Tara Vehiculo" & Space(17), 17) & " " & Right(Space(7), 7) & " "
        
        If DBLet(Taras, "N") <> 0 Then
            Lin = Lin & Right(Space(7) & Format(Taras, "###,##0"), 7)
        End If
        
        Cabecera.Add Lin
        
        Lin = Space(MargenIzdo) & Left(Space(17), 17) & " " & Right(Space(7), 7) & " "
        Lin = Lin & Right(Space(7) & "-------", 7)
        Cabecera.Add Lin

        Lin = Space(MargenIzdo) & Left("TOTAL TARAS" & Space(17), 17) & " " & Right(Space(7), 7) & " "
        
        If DBLet(Taras, "N") <> 0 Then
            Lin = Lin & Right(Space(7) & Format(Taras, "###,##0"), 7)
        End If
        
        Cabecera.Add Lin
        
        'Cerramos el rs
        RS1.Close
        Set RS1 = Nothing
        
        
        
        'Ya tenemos todos los datos
        'Ahora manadmos a la impresora
        ImprimeEnPapel
        
        
        
EImpD:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Imprimir directo."
        Err.Clear
    End If
    
    
    Set Cabecera = Nothing
    Set Lineas = Nothing
    Set Importes = Nothing
    Set rsIVA = New ADODB.Recordset
    Set RS1 = New ADODB.Recordset
    Exit Sub
    
End Sub



'***************************************
'******* ImprimirDirectoTickets ********
'***************************************
Public Sub ImprimirDirectoTickets(cadSelect As String)
    Dim NomImpre As String
  '  Dim FechaT As Date
    Dim rsIVA As ADODB.Recordset
'    Dim vFactu As CFactura
    
    Dim Sql As String
    Dim Lin As String ' línea de impresión
    Dim i As Integer
    
    Dim Producto As String
    Dim Variedad As String
    Dim Partida As String
    Dim Termino As String
    Dim Hdas As Currency
    Dim Has As Currency
    Dim TipoEntrada As String
    Dim SegundaImpresion As Boolean
    Dim Mermas As Long
    Dim Taras As Long
    Dim Taras2 As Long
    Dim Socio As String
    
On Error GoTo EImpD
    
        AccionesIniciales
        
        Set RS1 = New ADODB.Recordset
        
        'Cabecera de la entrada
        Sql = "select * from trzpalets WHERE " & cadSelect
        RS1.Open Sql, conn, adOpenForwardOnly
        
        Socio = DevuelveValor("select nomsocio from rsocios where codsocio = " & DBSet(RS1!Codsocio, "N"))
        Variedad = DevuelveValor("select nomvarie from variedades where codvarie = " & DBSet(RS1!codvarie, "N"))
        
        Set Cabecera = New Collection
        
        For i = 1 To 1
            Cabecera.Add " "
        Next i
        
        Lin = Space(MargenIzdo) & Left(Socio & Space(40), 40)
        'EN la impresora se alineara la linea roja del cabezal con la linea superiror del papel impreso (en verde)
        'Añadairemos una linea en blanco
        Cabecera.Add Lin
        
        Cabecera.Add " "
        
        
        Lin = Space(MargenIzdo) & Left(Variedad & Space(40), 40)
        Cabecera.Add Lin          '1234567890
        
        Cabecera.Add " "          '1234567890
        
        
        Lin = Space(MargenIzdo) & "Barras : "
        Cabecera.Add Lin          '1234567890
        
        
        'Cerramos el rs
        RS1.Close
        Set RS1 = Nothing
        
        
        'Ya tenemos todos los datos
        'Ahora manadmos a la impresora
        ImprimeEnPapel
        
EImpD:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Imprimir directo."
        Err.Clear
    End If
    
    
    Set Cabecera = Nothing
    Set Lineas = Nothing
    Set Importes = Nothing
    Set rsIVA = New ADODB.Recordset
    Set RS1 = New ADODB.Recordset
    Exit Sub
    
End Sub


