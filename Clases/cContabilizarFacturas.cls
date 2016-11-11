VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cContabilizarFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------
'
'   David Gandul
'   ------------
'
'   Fecha: Enero 2009
'
'   Tendra DOS metodos publicos que serviran para contabilizar
'   las facturas
'
'   -Clientes:
'   -Proveedores:
'
'   Antes de llamar a cualquiera de los metodos, tendremos que establecer
'   los valores inciales de cntabilizacion. Seran variables itnernas
'   que llevaran, por ejemplo, digitos niveles, conceptos de integracion
'   con lo cual, podria darse el caso de que no se cargaran correctamente.
'   Para ello ademas que la funcion devolvera un FALSE, habra una propiedad
'   de la clase que sabremos si esta bien incializada o no
'
'   Se asume que (IMPORTANTISIMO):
'
'       -- Habra una conexion(connconta) que apuntara a la contabilidad,
'       y se pasa por referencia en el metodo
'
'       -- Hay un clase
'       -- la llamada al metodo se realizara DENTRO de una transaccion.
'       estara en manos del que use esta libreria  las acciones a realizar
'       al finalizar. bien sea el rollbacktrans o cualquier otro tipo de accion
'
'       -- Algunas funciones y variables deben estar an otras librerias:
'           :FormatoFecha, DBSET,recuperavalor
'   Ante cualquier duda o sugerencia, aqui en mi mesa(abajo de ella) tengo
'   el buzon donde depositarla ;)   jejeje




'
'       27 SEPTIEMBRE 2010
'       Vamos a borrar una factura que esta contbailizada. Con lo cual vamos a
'       a partir de un numero de factura eliminar de cbfact y si tuviera apunte, borrar
'       el apunte
'
'
'
'
'
'
'
'
'



Private mvarRealizarContabilizacion As Boolean
Private mvarNumeroFactura As Long
Private mvarAnofac As Integer
Private mvarSerie As String

Private mvarModoErrores As Byte
    '0. Inicial
    '1.  Mostrare un msg por cada una.
    
'Variables que se cargan al establecer ValoresInciales
Private miConexion As Connection
Private Observaciones As String
Private AmpliacionParametros As Byte
Private AmpliacionFacurasCli As String
Private AmpliacionFacurasPro As String
Private CCenFacturas As Boolean
Private Abononeg As Boolean
Private Autocoste As Boolean  'Si lleva anailitaca o no
Private DigitosNiveles(9) As Integer  'Digitos nivel
Private NumNiveles As Integer
Private FechaFinEjercicio As Date
Private Diarios As String
Private NumgeisEnAmpliacionPRO As Boolean
'Variables del asiento
Private FechaAsi As Date
Private Numasiento As Long
Private DiarioFacturas As Integer

'Variables
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim cuenta As String
Dim CCost As String
Dim Rs As ADODB.Recordset
Dim Sql As String


Dim ColErrores As Collection

'SOLO GET
Public Property Get RealizarContabilizacion() As Boolean
     RealizarContabilizacion = mvarRealizarContabilizacion
End Property

Public Property Get NumeroFactura() As Long
     NumeroFactura = mvarNumeroFactura
End Property


Public Property Get Anofac() As Integer
     Anofac = mvarAnofac
End Property


Public Property Get Serie() As String
     Serie = mvarSerie
End Property


'-----------------------------------------------------------
'-----------------------------------------------------------
'Con esta funcion, fijaremos las variables inciales para la contabilizacion
'   Esto puede hacerse , antes del bucle de contabilziacion, y de ese modo
'   solo lo hacemos una vez
'
Public Function EstablecerValoresInciales(ByRef mConn As Connection) As Boolean
Dim ParametroContabilizacion As Boolean
On Error GoTo EEstablecerValoresInciales
    
        EstablecerValoresInciales = False
        mvarRealizarContabilizacion = False
        ParametroContabilizacion = False
        Set Rs = New ADODB.Recordset
        Sql = "Select * from parametros"
        Set miConexion = mConn
        
        
        Rs.Open Sql, miConexion, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = ""
        If Rs.EOF Then Sql = "Falta configurar parametros"
        
        If Sql = "" Then
        
            'SI EL PARAMETRO de contabilziaccion autmatica de facturas  no esta habilitado entonces
            ' NO la contabilizo
            If vParamAplic.ContabilidadNueva Then
                ParametroContabilizacion = True
            Else
                If DBLet(Rs!ContabilizaFact, "N") = 1 Then ParametroContabilizacion = True
            End If
        
            If Not ASignarConceptosFacturas Then
                Sql = "Error al asignar los conceptos de facturas"
            Else
                Sql = ""
            End If
        End If
               
        If Sql = "" Then
            'hasta ahora va todo bien
            Sql = "Ampliacion, CC,abononeg,analitica,fechaFinEjercicio,diario facturas(cli,pro)"
            AmpliacionParametros = Rs!nctafact
            CCenFacturas = DBLet(Rs!CCenFacturas, "N") = 1
            Abononeg = DBLet(Rs!Abononeg, "N") = 1
            Autocoste = DBLet(Rs!Autocoste, "N")
            FechaFinEjercicio = Rs!FechaFin
            NumgeisEnAmpliacionPRO = Rs!CodiNume = 1
            Diarios = Rs!numdiacl & "|" & Rs!numdiapr & "|"
            
            Sql = ""
        End If
        Rs.Close
        
        If Sql <> "" Then
            MsgBox Sql, vbExclamation
            Exit Function
        End If
        
        
        'Cargo los niveles desde la tabla empresa
        Sql = "numdigi1,numdigi2,numdigi3,numdigi4,numdigi5,numdigi6,numdigi7,numdigi8,numdigi9,numdigi10,numnivel"
        Sql = "Select " & Sql & " from empresa"
        Rs.Open Sql, miConexion, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = ""
        If Rs.EOF Then Sql = "Falta configurar empresa"
        
        If Sql = "" Then
            'Asigno niveles
            Sql = "Datos basicos empresa en contabilidad."
            NumNiveles = Rs!numnivel
            
            'DigitosNiveles(0) = RS!numdigi1
            For Numasiento = 0 To 9
                DigitosNiveles(Numasiento) = DBLet(Rs.Fields(Numasiento), "N")
            Next Numasiento
        End If
        Rs.Close
        Set Rs = Nothing
        'Esto es perosnalilable por aplicacion
        Observaciones = "Contabilizado desde ARIAGROREC. Fecha: " & Format(Now, "dd/mm/yyyy")
        


        'Todo bien leido. Podra contabilizar si asi lo indica el parametro
        mvarRealizarContabilizacion = ParametroContabilizacion
        EstablecerValoresInciales = True
        'Para los errores de contabilizacion
        If ParametroContabilizacion Then Set ColErrores = New Collection
        
        Exit Function
EEstablecerValoresInciales:
    MsgBox "Error estableciendo valores iniciales: " & vbCrLf & Err.Description & vbCrLf & Sql, vbExclamation
End Function

Public Sub AnyadeElError(CadenaError As String)
    CadenaError = Trim(CadenaError)
    If CadenaError <> "" Then
        Sql = mvarSerie & "|" & mvarNumeroFactura & "|" & mvarAnofac & "|" & CadenaError & "|"
        ColErrores.Add Sql
    End If
End Sub

'Con este sub fijaremos los valores de la factura, para luego contabilizarlos y no tener que pasarles variables pa'arriba y p`'abajo
Public Sub FijarNumeroFactura(numFac As Long, Anofac As Integer, numserie As String)
    mvarNumeroFactura = numFac
    mvarAnofac = Anofac
    mvarSerie = numserie
End Sub





'-----------------------------------------------------------
'-----------------------------------------------------------
'-----------------------------------------------------------
'SI TODO OK devuelve ""
Public Function IntegraLaFacturaCliente(numFac As Long, Anofac As Integer, numserie As String) As String
Dim cad As String
Dim cad2 As String
Dim Cad3 As String
Dim Amplia2 As String
Dim DocConcAmp As String
Dim RF As Recordset
Dim ImporteNegativo As Boolean
Dim Importe0 As Boolean
Dim PrimeraContrapartida As String
Dim A_Donde As String
Dim NumLinea As Integer

Dim Importe As Currency
Dim DatosRetencion As String
Dim Cliente As String

    On Error GoTo EIntegraLaFactura
    'Sabemos que
    'numfac     --> CODIGO FACTURA
    'anofac      --> AÑO FACTURA
    'NUmSerie       --> SERIE DE LA FACTURA

    
    
    'Obtenemos los datos de la factura
    A_Donde = "Fra: " & numserie & Format(numFac, "000000") & " / " & Anofac
    Set RF = New ADODB.Recordset
    
    If vParamAplic.ContabilidadNueva Then
        Sql = "SELECT numserie, numfactu codfaccl, fecfactu fecfaccl,factcli.codmacta"
        Sql = Sql & ", totfaccl ,factcli.observa confaccl  ,cuereten  , trefaccl"
        Sql = Sql & " ,numasien,cuentas.nommacta FROM factcli,cuentas"
        Sql = Sql & " WHERE numserie='" & numserie
        Sql = Sql & "' AND numfactu= " & numFac
        Sql = Sql & " AND anofactu=" & Anofac
        Sql = Sql & " AND factcli.codmacta = cuentas.codmacta"
    
    Else
        Sql = "SELECT *,nommacta FROM cabfact,cuentas"
        Sql = Sql & " WHERE numserie='" & numserie
        Sql = Sql & "' AND codfaccl= " & numFac
        Sql = Sql & " AND anofaccl=" & Anofac
        Sql = Sql & " AND cabfact.codmacta = cuentas.codmacta"
    
    End If
        
    
    RF.Open Sql, miConexion, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If RF.EOF Then
        IntegraLaFacturaCliente = "No se encuentra la factura: " & vbCrLf & A_Donde
        RF.Close
        Exit Function
    End If
    
    
    If Not IsNull(RF!numasien) Then
        IntegraLaFacturaCliente = "Factura contabilizada: " & vbCrLf & A_Donde & "  Num: " & RF!numasien
        RF.Close
        Exit Function
    End If
    
    'COnseguir contador y eso
    FechaAsi = RF!fecfaccl
    DiarioFacturas = Val(RecuperaValor(Diarios, 1)) 'Clientes 1,proveedores 2
    'Consigo contador
    '****************************************+
   
    Do
        Numasiento = ConseguirContador(FechaAsi <= FechaFinEjercicio, A_Donde)
        If Numasiento = 0 Then
            'Ha habido algun tipo de error.
            
           
                IntegraLaFacturaCliente = A_Donde
            
                RF.Close
                Exit Function
       End If
    Loop Until Numasiento > 0
    
    'Cabecera del hco de apuntes
    A_Donde = "Inserta cabecera hco apuntes"
    Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari"
    If vParamAplic.ContabilidadNueva Then Sql = Sql & ",feccreacion,usucreacion,desdeaplicacion"
    Sql = Sql & ") VALUES ("
    Sql = Sql & DiarioFacturas & ",'" & Format(FechaAsi, FormatoFecha) & "'," & Numasiento
    Sql = Sql & "," & DBSet(Observaciones, "T", "S")
    If vParamAplic.ContabilidadNueva Then Sql = Sql & "," & DBSet(Now, "FH") & "," & DBSet(vSesion.Login, "T") & ",'Arigasol'"
    miConexion.Execute Sql & ")"
    
    'Lineas fijas, es decir la linea de cliente, importes y tal y tal
    'Para el sql
    
    cad = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, "
    cad = cad & "codconce,ampconce, timporteD, timporteH,codccost, ctacontr, idcontab, punteada)"
    
    
    
    
    cad = cad & " VALUES (" & DiarioFacturas & ",'" & Format(FechaAsi, FormatoFecha) & "'," & Numasiento & ","
    NumLinea = 1 'Contador de lineas
    
    
    A_Donde = "Linea cliente"
    
    'Guardo unos datos
    DatosRetencion = ""
    If Not IsNull(RF!cuereten) Then DatosRetencion = RF!cuereten & "|" & RF!trefaccl & "|"
    
    '-------------------------------------------------------------------
    'LINEA Cliente
    Sql = NumLinea & ",'" & RF!Codmacta & "',"
    

    ' en AmpliacionFacurasCli_ van:
    '  concepto normal|nomconce|conceto abono|nomconce|
    If RF!totfaccl >= 0 Then
        DocConcAmp = RecuperaValor(AmpliacionFacurasCli, 1)
    Else
        DocConcAmp = RecuperaValor(AmpliacionFacurasCli, 3)
    End If
    DocConcAmp = "'" & Format(numFac, "00000") & "'," & DocConcAmp & ",'"
    
    
    'Ampliacion segun parametros
    Select Case AmpliacionParametros
    Case 1
        If RF!totfaccl < 0 Then
            cad2 = RecuperaValor(AmpliacionFacurasCli, 4)
        Else
            cad2 = RecuperaValor(AmpliacionFacurasCli, 2)
        End If
        '28/02/2007.
        'Añado numerie
        cad2 = cad2 & " " & numserie & Format(numFac, "00000")
    Case 2
        cad2 = DevNombreSQL(DBLet(RF!Nommacta))
    Case Else
        cad2 = DBLet(RF!confaccl)
    End Select
    
    '   Modificacion para k aparezca en la ampliacio el CC en la ampliacion de codmacta
    '
    Amplia2 = cad2
    If CCenFacturas Then
        A_Donde = "CC en Facturas."
        Cad3 = DevuelveCentroCosteFactura(True, PrimeraContrapartida, numFac, Anofac, numserie)
        If Cad3 <> "" Then
            If Len(Amplia2) > 21 Then Amplia2 = Mid(Amplia2, 1, 21)
            'Opcion1
            'Amplia2 = Amplia2 & " .CC:" & Cad3
            'Opcion2
            Amplia2 = Amplia2 & " [" & Cad3 & "]"
        End If
    End If
    A_Donde = "Linea cliente. Ampliacion2"
    
    
    Sql = Sql & DocConcAmp & Amplia2 & "'"
    DocConcAmp = DocConcAmp & cad2 & "'"   'DocConcAmp Sirve para el IVA
    
    'Esta variable sirve para las demas
    ImporteNegativo = (RF!totfaccl < 0)
    
    'Importes, atencion importes negativos
    '  antes --> Cad2 = CadenaImporte(ImporteNegativo, True, RF!totfaccl)
    cad2 = CadenaImporte(True, RF!totfaccl, Importe0)
    Sql = Sql & "," & cad2 & ",NULL,"
    
    'Contrpartida. 28 Marzo 2006
    If PrimeraContrapartida <> "" Then
        Sql = Sql & "'" & PrimeraContrapartida & "'"
    Else
        Sql = Sql & "NULL"
    End If
    Sql = Sql & ",'FRACLI',0)"
    
    
    miConexion.Execute cad & Sql
    NumLinea = NumLinea + 1 'Es el contador de lineaapunteshco
    Cliente = RF!Codmacta
    
    
    If vParamAplic.ContabilidadNueva Then
        
        RF.Close
        
        'Abrimos el RF con las lineas del IVA
                
        Sql = "SELECT numserie, numfactu codfaccl, fecfactu fecfaccl"
        Sql = Sql & ", baseimpo,codigiva,porciva,porcrec,impoiva,imporec  FROM factcli_totales "
        Sql = Sql & " WHERE numserie='" & numserie
        Sql = Sql & "' AND numfactu= " & numFac
        Sql = Sql & " AND anofactu=" & Anofac
        RF.Open Sql, miConexion, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RF.EOF
            
            
            A_Donde = "IVA " & NumLinea - 1
            Cad3 = "cuentarr"
            cad2 = DevuelveDesdeBDModulo("cuentare", "tiposiva", "codigiva", RF!Codigiva, "N", Cad3)
            If cad2 <> "" Then
                Sql = NumLinea & ",'" & cad2 & "'," & DocConcAmp
                cad2 = CadenaImporte(False, RF!ImpoIva, Importe0)
                Sql = Sql & "," & cad2 & ","
                Sql = Sql & "NULL,'" & Cliente & "','FRACLI',0)"
                If Not Importe0 Then
                    miConexion.Execute cad & Sql
                    NumLinea = NumLinea + 1
                End If
                
                'La de recargo  1-----------------
                If Not IsNull(RF!ImpoREC) Then
                         Sql = NumLinea & "," & Cad3 & "," & DocConcAmp
                        'Importes, atencion importes negativos
                        cad2 = CadenaImporte(False, RF!ImpoREC, Importe0)
                        Sql = Sql & "," & cad2 & ","
                        Sql = Sql & "NULL,'" & Cliente & "','FRACLI',0)"
                        If Not Importe0 Then
                            miConexion.Execute cad & Sql
                            NumLinea = NumLinea + 1
                        End If
                End If
            
            End If
        
        
        
            NumLinea = NumLinea + 1 'Es el contador de lineaapunteshco
            RF.MoveNext
        Wend
        
    Else
        '------------------------------------------------------------
        'Cuentas de IVA. Cuenta 1
        A_Donde = "IVA 1"
        Cad3 = "cuentarr"
        cad2 = DevuelveDesdeBDModulo("cuentare", "tiposiva", "codigiva", RF!tp1faccl, "N", Cad3)
        If cad2 <> "" Then
            Sql = NumLinea & ",'" & cad2 & "'," & DocConcAmp
            cad2 = CadenaImporte(False, RF!ti1faccl, Importe0)
            Sql = Sql & "," & cad2 & ","
            Sql = Sql & "NULL,'" & RF!Codmacta & "','FRACLI',0)"
            If Not Importe0 Then
                miConexion.Execute cad & Sql
                NumLinea = NumLinea + 1
            End If
            
            'La de recargo  1-----------------
            If Not IsNull(RF!tr1faccl) Then
                     Sql = NumLinea & "," & Cad3 & "," & DocConcAmp
                    'Importes, atencion importes negativos
                    cad2 = CadenaImporte(False, RF!tr1faccl, Importe0)
                    Sql = Sql & "," & cad2 & ","
                    Sql = Sql & "NULL,'" & RF!Codmacta & "','FRACLI',0)"
                    If Not Importe0 Then
                        miConexion.Execute cad & Sql
                        NumLinea = NumLinea + 1
                    End If
            End If
        Else
            
            IntegraLaFacturaCliente = "Error leyendo TIPO de IVA: " & RF!tp1faccl
            RF.Close
            Exit Function
        End If
        
        
        
        
        
        
        '------------------------------------------------------------
        'Cuentas de IVA. Cuenta 2
        A_Donde = "IVA 2"
        If Not IsNull(RF!tp2faccl) Then
            Cad3 = "cuentarr"
            cad2 = DevuelveDesdeBDModulo("cuentare", "tiposiva", "codigiva", RF!tp2faccl, "N", Cad3)
            If cad2 <> "" Then
                Sql = NumLinea & ",'" & cad2 & "'," & DocConcAmp
                'Importes, atencion importes negativos
                cad2 = CadenaImporte(False, RF!ti2faccl, Importe0)
                Sql = Sql & "," & cad2 & ","
                Sql = Sql & "NULL,'" & RF!Codmacta & "','FRACLI',0)"
                If Not Importe0 Then
                    miConexion.Execute cad & Sql
                    NumLinea = NumLinea + 1
                End If
                
                'La de recargo  1-----------------
                If Not IsNull(RF!tr2faccl) Then
                        Sql = NumLinea & ",'" & Cad3 & "'," & DocConcAmp
                        'Importes, atencion importes negativos
                        cad2 = CadenaImporte(False, RF!tr2faccl, Importe0)
                        Sql = Sql & "," & cad2 & ","
                        Sql = Sql & "NULL,'" & RF!Codmacta & "','FRACLI',0)"
                        If Not Importe0 Then
                            miConexion.Execute cad & Sql
                            NumLinea = NumLinea + 1
                        End If
                End If
            Else
                IntegraLaFacturaCliente = "Error leyendo TIPO de IVA2: " & RF!tp2faccl
                RF.Close
                Exit Function
            End If
            
        End If 'De IVA 2
        
        
        
         '------------------------------------------------------------
        'Cuentas de IVA. Cuenta 2
        A_Donde = "IVA 3"
        If Not IsNull(RF!tp3faccl) Then
            Cad3 = "cuentarr"
            cad2 = DevuelveDesdeBDModulo("cuentare", "tiposiva", "codigiva", RF!tp3faccl, "N", Cad3)
            If cad2 <> "" Then
                Sql = NumLinea & ",'" & cad2 & "'," & DocConcAmp
                'Importes, atencion importes negativos
                cad2 = CadenaImporte(False, RF!ti3faccl, Importe0)
                Sql = Sql & "," & cad2 & ","
                Sql = Sql & "NULL,'" & RF!Codmacta & "','FRACLI',0)"
                If Not Importe0 Then
                    miConexion.Execute cad & Sql
                    NumLinea = NumLinea + 1
                End If
                
                'La de recargo  1-----------------
                If Not IsNull(RF!tr3faccl) Then
                         Sql = NumLinea & ",'" & Cad3 & "'," & DocConcAmp
                        'Importes, atencion importes negativos
                        cad2 = CadenaImporte(False, RF!tr3faccl, Importe0)
                        Sql = Sql & "," & cad2 & ","
                        Sql = Sql & "NULL,'" & RF!Codmacta & "','FRACLI',0)"
                        If Not Importe0 Then
                            miConexion.Execute cad & Sql
                            NumLinea = NumLinea + 1
                        End If
                End If
            Else
                IntegraLaFacturaCliente = "Error leyendo TIPO de IVA 3: " & Rs!tp3faccl
                RF.Close
                Exit Function
            End If
            
        End If 'De IVA 3
    End If  'DE LOS IVAS
    
    RF.Close
    
    '-------------------------------------
    ' RETENCION
    A_Donde = "Retencion"
    
    'hay una cadena con los datos, pq el rf deberaimos haberlo cerrado ya
    
    If DatosRetencion <> "" Then
        
        Sql = NumLinea & ",'" & RF!cuereten & "'," & DocConcAmp
        'Importes, atencion importes negativos
        cad2 = CadenaImporte(True, RF!trefaccl, Importe0)
        Sql = Sql & "," & cad2 & ","
        Sql = Sql & "NULL,NULL,'FRACLI',0)"
       
        miConexion.Execute cad & Sql
        NumLinea = NumLinea + 1 'Es el contador de lineaapunteshco
    End If
    
    
    
    '------------------------------------------------------------
    'Las lineas de la factura.
    
    A_Donde = "Leyendo lineas factura"
    
    
    If vParamAplic.ContabilidadNueva Then
        Sql = "SELECT codmacta codtbase, baseimpo impbascl,codccost  FROM factcli_lineas "
        Sql = Sql & " WHERE numserie='" & numserie
        Sql = Sql & "' AND numfactu= " & numFac
        Sql = Sql & " AND anofactu=" & Anofac
        
    Else
        Sql = "Select linfact.* , codmacta FROM linfact,Cuentas "
        Sql = Sql & " WHERE numserie='" & numserie
        Sql = Sql & "' AND codfaccl= " & numFac
        Sql = Sql & " AND anofaccl=" & Anofac
        Sql = Sql & " AND linfact.codtbase = Cuentas.codmacta"
    End If
    RF.Open Sql, miConexion, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Para cada linea insertamos
    cad2 = ""
    A_Donde = "Procesando lineas"
    While Not RF.EOF
        'Importes, atencion importes negativos
        If cad2 = "" Then PrimeraContrapartida = RF!codtbase
        Sql = NumLinea & ",'" & RF!codtbase & "'," & DocConcAmp
        cad2 = CadenaImporte(False, RF!impbascl, Importe0)
        Sql = Sql & "," & cad2 & ","
        If IsNull(RF!CodCCost) Then
            cad2 = "NULL"
        Else
            cad2 = "'" & RF!CodCCost & "'"
        End If
        
        Sql = Sql & cad2 & ",'" & Cliente & "','FRACLI',0)"
    
        miConexion.Execute cad & Sql
        NumLinea = NumLinea + 1 'Es el contador de lineaapunteshco
        

        RF.MoveNext
        If Not RF.EOF Then PrimeraContrapartida = ""
    Wend
    RF.Close
    
    
    
    
    'AHora viene lo bueno.  MARZO 2006
    'Si el valor fuera true YA lo habria insertado en la cabcera
    If CCenFacturas Then
        If PrimeraContrapartida <> "" Then
            Sql = "UPDATE hlinapu SET ctacontr ='" & PrimeraContrapartida & "'"
            Sql = Sql & " WHERE numdiari = " & DiarioFacturas & " AND fechaent ='" & Format(FechaAsi, FormatoFecha) & "' and numasien = " & Numasiento
            Sql = Sql & " AND linliapu =1 " 'LA PRIMERA LINEA SIEMPRE ES LA DE LA CUENTA
            miConexion.Execute Sql  '
        End If
    End If
        
    
    
    
    'Actualimos en factura, el nº de asiento
    If vParamAplic.ContabilidadNueva Then
        Sql = "UPDATE factcli SET numdiari = " & DiarioFacturas & ", fechaent = '" & Format(FechaAsi, FormatoFecha) & "', numasien =" & Numasiento
        Sql = Sql & " WHERE numserie='" & numserie
        Sql = Sql & "' AND numfactu= " & numFac
        Sql = Sql & " AND anofactu=" & Anofac
    Else
        Sql = "UPDATE cabfact SET numdiari = " & DiarioFacturas & ", fechaent = '" & Format(FechaAsi, FormatoFecha) & "', numasien =" & Numasiento
        Sql = Sql & " WHERE numserie='" & numserie
        Sql = Sql & "' AND codfaccl= " & numFac
        Sql = Sql & " AND anofaccl=" & Anofac
    End If
    miConexion.Execute Sql
    
    'Para los saldos ponemos el numero de asiento donde toca
    If Not vParamAplic.ContabilidadNueva Then
        A_Donde = CalcularLineasYSaldosFacturas(True)
        IntegraLaFacturaCliente = A_Donde
    End If
    Set RF = Nothing
    
    Exit Function
EIntegraLaFactura:
    IntegraLaFacturaCliente = A_Donde & vbCrLf & Err.Description & vbCrLf & Sql
    Set RF = Nothing
End Function





'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
' Facturas INTERNAS.
' No se han pasado al registro. La contabilizacion es la creacion del apunte
' Saca los datos desde ariges scafac
'
'   le paso el Recordset con los datos cargados.
'---------------------------------------------------------------------------
'SI TODO OK devuelve ""
Public Function IntegraLaFacturaClienteINTERNA(ByRef RCab As ADODB.Recordset, ByRef RLin As ADODB.Recordset) As String
Dim cad As String
Dim cad2 As String
Dim Cad3 As String
Dim Amplia2 As String
Dim DocConcAmp As String
Dim ImporteNegativo As Boolean
Dim Importe0 As Boolean
Dim PrimeraContrapartida As String
Dim A_Donde As String
Dim NumLinea As Integer

Dim Importe As Currency


    On Error GoTo eIntegraLaFacturaClienteINTERNA
    
    
    'Obtenemos los datos de la factura
    A_Donde = "Fra interna: " & Format(RCab!numfactu, "000000") & " / " & RCab!fecfactu
    
    
    
    'COnseguir contador y eso
    FechaAsi = RCab!fecfactu
    DiarioFacturas = Val(RecuperaValor(Diarios, 1)) 'Clientes 1,proveedores 2
    'Consigo contador
    '****************************************+
   
    Do
        Numasiento = ConseguirContador(FechaAsi <= FechaFinEjercicio, A_Donde)
        If Numasiento = 0 Then
            'Ha habido algun tipo de error.
                IntegraLaFacturaClienteINTERNA = A_Donde
                Exit Function
       End If
    Loop Until Numasiento > 0
    
    'Cabecera del hco de apuntes
    A_Donde = "Inserta cabecera hco apuntes"
    Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari"
    If vParamAplic.ContabilidadNueva Then Sql = Sql & ",feccreacion,usucreacion,desdeaplicacion"
    Sql = Sql & ") VALUES ("
    'Ant 17/OCT/2005
    'SQL = SQL & vParam.numdiacl & ",'" & Fecha & "'," & NumAsiento
    Sql = Sql & DiarioFacturas & ",'" & Format(FechaAsi, FormatoFecha) & "'," & Numasiento
    cad = Observaciones & vbCrLf & "Factura interna: " & Format(RCab!numfactu, "00000") & " - " & RCab!fecfactu
    Sql = Sql & "," & DBSet(cad, "T", "S")
    If vParamAplic.ContabilidadNueva Then Sql = Sql & "," & DBSet(Now, "FH") & "," & DBSet(vSesion.Login, "T") & ",'ARIAGROREC'"
    Sql = Sql & ")"
    cad = ""
    miConexion.Execute Sql
    
    'Lineas fijas, es decir la linea de cliente, importes y tal y tal
    'Para el sql
    cad = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, "
    cad = cad & "codconce,ampconce, timporteD, timporteH,codccost, ctacontr, idcontab, punteada)"
    'Ant 17/oct/05
    'Cad = Cad & " VALUES (" & vParam.numdiacl & ",'" & Fecha & "'," & NumAsiento & ","
    cad = cad & " VALUES (" & DiarioFacturas & ",'" & Format(FechaAsi, FormatoFecha) & "'," & Numasiento & ","
    NumLinea = 1 'Contador de lineas
    
    
    A_Donde = "Linea cliente"
    '-------------------------------------------------------------------
    'LINEA Cliente
    Sql = NumLinea & ",'" & RCab!Codmacta & "',"
    

    ' en AmpliacionFacurasCli_ van:
    '  concepto normal|nomconce|conceto abono|nomconce|
    If RCab!TotalFac >= 0 Then
        DocConcAmp = RecuperaValor(AmpliacionFacurasCli, 1)
    Else
        DocConcAmp = RecuperaValor(AmpliacionFacurasCli, 3)
    End If
    'DocConcAmp = "'" & Format(numFac, "00000") & "'," & DocConcAmp & ",'"
    DocConcAmp = "'" & Format(RCab!numfactu, "0000000000") & "'," & DocConcAmp & ",'"
    
    
    'Ampliacion segun parametros
    Select Case AmpliacionParametros
    Case 1
        If RCab!TotalFac < 0 Then
            cad2 = RecuperaValor(AmpliacionFacurasCli, 4)
        Else
            cad2 = RecuperaValor(AmpliacionFacurasCli, 2)
        End If
        cad2 = cad2 & " interna " & Format(RCab!numfactu, "00000")
    Case 2
        cad2 = DevuelveDesdeBD(cConta, "nommacta", "cuentas", "codmacta", RCab!Codmacta, "T")
        cad2 = DevNombreSQL(DBLet(cad2))
    Case Else
        'Este no deberia ser
        cad2 = "" 'DBLet(RCab!confaccl)
    End Select
    
    '   Modificacion para k aparezca en la ampliacio el CC en la ampliacion de codmacta
    '
    Amplia2 = cad2
    If CCenFacturas Then
        'FALTA ###
    
'        A_Donde = "CC en Facturas."
'        Cad3 = DevuelveCentroCosteFactura(True, PrimeraContrapartida, numFac, Anofac, numSerie)
'        If Cad3 <> "" Then
'            If Len(Amplia2) > 21 Then Amplia2 = Mid(Amplia2, 1, 21)
'            'Opcion1
'            'Amplia2 = Amplia2 & " .CC:" & Cad3
'            'Opcion2
'            Amplia2 = Amplia2 & " [" & Cad3 & "]"
'        End If
    End If
    A_Donde = "Linea cliente. Ampliacion2"
    
    
    Sql = Sql & DocConcAmp & Amplia2 & "'"
    DocConcAmp = DocConcAmp & cad2 & "'"   'DocConcAmp Sirve para el IVA
    
    'Esta variable sirve para las demas
    ImporteNegativo = (RCab!TotalFac < 0)
    
    'Importes, atencion importes negativos
    '  antes --> Cad2 = CadenaImporte(ImporteNegativo, True, RF!totfaccl)
    cad2 = CadenaImporte(True, RCab!TotalFac, Importe0)
    Sql = Sql & "," & cad2 & ",NULL,"
    
    'Contrpartida. 28 Marzo 2006
    If PrimeraContrapartida <> "" Then
        Sql = Sql & "'" & PrimeraContrapartida & "'"
    Else
        Sql = Sql & "NULL"
    End If
    Sql = Sql & ",'CONTAB',0)"
    
    
    miConexion.Execute cad & Sql
    NumLinea = NumLinea + 1 'Es el contador de lineaapunteshco
    
    
    
    '------------------------------------------------------------
    '
    '   No llevan IVA
    '
    '------------------------------------------------------------
    
    '------------------------------------------------------------
    'Las lineas de la factura. Para ello guardaremos algunos datos
    cad2 = RCab!Codmacta
    Importe = RCab!TotalFac
    
    
    
    'Cerramos el RF
    cuenta = RCab!Codmacta
    
    
    
    
    A_Donde = "Leyendo lineas factura"
    'Para cada linea insertamos
    cad2 = ""
    A_Donde = "Procesando lineas"
    While Not RLin.EOF
        'Importes, atencion importes negativos
        If cad2 = "" Then PrimeraContrapartida = RLin!cuenta
        Sql = NumLinea & ",'" & RLin!cuenta & "'," & DocConcAmp
        cad2 = CadenaImporte(False, RLin!Importe, Importe0)
        Sql = Sql & "," & cad2 & ","
        cad2 = "NULL"
        'Llevara CC
        If RLin.Fields.Count > 6 Then
            If Not IsNull(RLin!CodCCost) Then cad2 = "'" & RLin!CodCCost & "'"
        End If
        
        Sql = Sql & cad2 & ",'" & cuenta & "','contab',0)"
    
        miConexion.Execute cad & Sql
        NumLinea = NumLinea + 1 'Es el contador de lineaapunteshco
        

        RLin.MoveNext
        If Not RLin.EOF Then PrimeraContrapartida = ""
    Wend
    RLin.Close
    
    
    
    
    'AHora viene lo bueno.  MARZO 2006
    'Si el valor fuera true YA lo habria insertado en la cabcera
    If NumLinea = 3 Then
        If PrimeraContrapartida <> "" Then
            Sql = "UPDATE hlinapu SET ctacontr ='" & PrimeraContrapartida & "'"
            Sql = Sql & " WHERE numdiari = " & DiarioFacturas & " AND fechaent ='" & Format(FechaAsi, FormatoFecha) & "' and numasien = " & Numasiento
            Sql = Sql & " AND linliapu =1 " 'LA PRIMERA LINEA SIEMPRE ES LA DE LA CUENTA
            miConexion.Execute Sql  '
        End If
    End If
        
    
    
    
    'Para los saldos ponemos el numero de asiento donde toca
    If Not vParamAplic.ContabilidadNueva Then
        A_Donde = CalcularLineasYSaldosFacturas(True)
        IntegraLaFacturaClienteINTERNA = A_Donde
    End If
    
        
    
    Exit Function
eIntegraLaFacturaClienteINTERNA:
    IntegraLaFacturaClienteINTERNA = A_Donde & vbCrLf & Err.Description & vbCrLf & Sql
End Function














'-----------------------------------------------------------
'-----------------------------------------------------------
'-----------------------------------------------------------
'-----------------------------------------------------------
Public Function IntegraLaFacturaProv(NumRegis As Long, Anofac As Integer) As String
Dim cad As String
Dim cad2 As String
Dim Cad3 As String
Dim DocConcAmp As String
Dim Amplia2 As String
Dim RF As Recordset
Dim ImporteNegativo As Boolean
Dim Importe0 As Boolean 'Para saber si el importe es 0
Dim PrimeraContrapartida As String  'Si hay solo una linea entonces la pondremos como contrapartida de la primera base
Dim A_Donde As String
Dim contLinea As Integer
Dim ColumnaIVA As String
Dim DatosRetencion  As String
    
    
'Marzo 2015
' Factura inversion de sujeto pasivo
' Por cada IVA entrará un linea mas , ya que ademas del soportado, pondremos el repercutido
'Dim InversionSujetoPasivo As Boolean    ver bajo feb2016
'Febrero 2016
'El tratamiento contable de las facturas intracomunitarias es el mismo
'Por lo tanto La variable que era InversionSujetoPasivo pasa a ser NoSumaIVA
Dim NoSumaIVA As Boolean

    
    
    
    
    
    On Error GoTo EIntegraLaFactura2
    
    
    'Obtenemos los datos de la factura
    A_Donde = NumRegis & " / " & Anofac
    Set RF = New ADODB.Recordset
    
    
    If vParamAplic.ContabilidadNueva Then
        Sql = "SELECT numserie, numfactu numfacpr, fecfactu fecfacpr,factpro.codmacta,fecharec fecrecpr,"
        Sql = Sql & " totfacpr ,factpro.observa confacpr,cuereten  , trefacpr ,numasien,cuentas.nommacta,codopera"
        Sql = Sql & " FROM factpro,cuentas "
        Sql = Sql & " WHERE numserie='" & SerieFraPro
        Sql = Sql & "' AND numregis= " & NumRegis
        Sql = Sql & " AND anofactu=" & Anofac
        Sql = Sql & " AND factpro.codmacta = cuentas.codmacta"
    
    Else
    
        Sql = "SELECT *,nommacta FROM cabfactprov,cuentas"
        Sql = Sql & " WHERE numregis = " & NumRegis
        Sql = Sql & " AND anofacpr=" & Anofac
        Sql = Sql & " AND cabfactprov.codmacta = cuentas.codmacta"
    End If
    RF.Open Sql, miConexion, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RF.EOF Then
        IntegraLaFacturaProv = "No se encuentra la factura: " & A_Donde
        RF.Close
        Exit Function
    End If
    
    If Not IsNull(RF!numasien) Then
        IntegraLaFacturaProv = "Ya esta contabilizada: " & A_Donde
        RF.Close
        Exit Function
    End If
    
    
    
    'COnseguir contador y eso
    FechaAsi = RF!fecrecpr
    DiarioFacturas = Val(RecuperaValor(Diarios, 2)) 'Clientes 1,proveedores 2
    'Consigo contador
    Do
        Numasiento = ConseguirContador(FechaAsi <= FechaFinEjercicio, A_Donde)
        If Numasiento = 0 Then
            IntegraLaFacturaProv = A_Donde
            RF.Close
            Exit Function
        End If
    Loop Until Numasiento > 0
    
    
    
    
    'Cabecera del hco de apuntes
    A_Donde = "Inserta cabecera hco apuntes"
    Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari"
    If vParamAplic.ContabilidadNueva Then Sql = Sql & ",feccreacion,usucreacion,desdeaplicacion"
    Sql = Sql & ") VALUES ("
    Sql = Sql & DiarioFacturas & ",'" & Format(FechaAsi, FormatoFecha) & "'," & Numasiento
    Sql = Sql & "," & DBSet(Observaciones, "T", "S")
    If vParamAplic.ContabilidadNueva Then Sql = Sql & "," & DBSet(Now, "FH") & "," & DBSet(vSesion.Login, "T") & ",'ARIAGRO'"
    Sql = Sql & ")"
    miConexion.Execute Sql
    
    
    If vParamAplic.ContabilidadNueva Then
       
        NoSumaIVA = RF!codopera = 4 Or RF!codopera = 1  'Febrerp 2016
        
        
    Else
        NoSumaIVA = RF!extranje = 3 Or RF!extranje = 1  'Febrerp 2016
    End If
    
    
    
    'Lineas fijas, es decir la linea de cliente, importes y tal y tal
    'Para el sql
    cad = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, "
    cad = cad & "codconce,ampconce, timporteD, timporteH,codccost, ctacontr, idcontab, punteada)"
    cad = cad & " VALUES (" & DiarioFacturas & ",'" & Format(FechaAsi, FormatoFecha) & "'," & Numasiento & ","
    
    contLinea = 1 'Contador de lineas
    PrimeraContrapartida = ""
    
    'Esta variable sirve para las demas
    ImporteNegativo = (RF!totfacpr < 0)
    A_Donde = "Linea proveedor"
    '-------------------------------------------------------------------
    'LINEA Proveedor
    Sql = contLinea & ",'" & RF!Codmacta & "',"
    
    
    'Documento "numdocum"
    If NumgeisEnAmpliacionPRO Then
        cad2 = Format(NumRegis, "0000000000")
    Else
        cad2 = DBLet(RF!numfacpr)
    End If
    
    'Conepto segin importe + o -
    If ImporteNegativo Then
        DocConcAmp = RecuperaValor(AmpliacionFacurasPro, 3)
    Else
        DocConcAmp = RecuperaValor(AmpliacionFacurasPro, 1)
    End If
    DocConcAmp = "'" & cad2 & "'," & DocConcAmp & ",'"
    
    
    'Ampliacion segun parametros
    Select Case AmpliacionParametros
    Case 1
        If ImporteNegativo Then
            cad2 = RecuperaValor(AmpliacionFacurasPro, 4)
        Else
            cad2 = RecuperaValor(AmpliacionFacurasPro, 2)
        End If
        cad2 = cad2 & " " & DevNombreSQL(RF!numfacpr)
        
        cad2 = cad2 & " (" & Format(RF!fecfacpr, "ddmmyy") & ")"
    Case 2
        cad2 = DevNombreSQL(DBLet(RF!Nommacta))
    Case Else
        cad2 = DBLet(RF!confacpr)
    End Select
    
        
    
    'Modificacion para k aparezca en la ampliacio el CC en la ampliacion de codmacta
    '
    Amplia2 = cad2
    If CCenFacturas Then
        A_Donde = "CC en Facturas."
        Cad3 = DevuelveCentroCosteFactura(False, PrimeraContrapartida, NumRegis, Anofac, "")
        If Cad3 <> "" Then
            If Len(Amplia2) > 26 Then Amplia2 = Mid(Amplia2, 1, 26)
            'Opcion1
            'Amplia2 = Amplia2 & " .CC:" & Cad3
            'Opcion2
            Amplia2 = Amplia2 & "[" & Cad3 & "]"
        End If
    End If
    A_Donde = "Linea cliente"
    
    
    Sql = Sql & DocConcAmp & Amplia2 & "'"
    DocConcAmp = DocConcAmp & cad2 & "'"   'DocConcAmp Sirve para el IVA
    
    
    'Importes, atencion importes negativos
    cad2 = CadenaImporte(False, RF!totfacpr, Importe0)
    Sql = Sql & "," & cad2 & ",NULL,"
    
    'Contrpartida. 28 Marzo 2006
    If PrimeraContrapartida <> "" Then
        Sql = Sql & "'" & PrimeraContrapartida & "'"
    Else
        Sql = Sql & "NULL"
    End If
    Sql = Sql & ",'FRAPRO',0)"
    
    miConexion.Execute cad & Sql
    contLinea = contLinea + 1 'Es el contador de lineaapunteshco
    
    
    DatosRetencion = ""
    If Not IsNull(RF!cuereten) Then DatosRetencion = RF!cuereten & "|" & RF!trefacpr & "|"
    
    '-------------------------------------------------------------------
    'LINEA Cliente
    Sql = contLinea & ",'" & RF!Codmacta & "',"
    
    
    
    
    
    
    
    If vParamAplic.ContabilidadNueva Then
        ColumnaIVA = "cuentaso"   'La normal
    Else
        If RF!Nodeducible = 1 Then
            'Es iva NO deducible
            ColumnaIVA = "cuentasn"
        Else
            ColumnaIVA = "cuentaso"   'La normal
        End If
    End If
    
    'Gauradmos estose valors para las lineas de factura
    ImporteD = RF!totfacpr
    cuenta = RF!Codmacta
    
    If vParamAplic.ContabilidadNueva Then
        
        RF.Close
        
        'Abrimos el RF con las lineas del IVA
                
        Sql = "SELECT numserie, numregis , fecharec fecfactu , baseimpo,codigiva,porciva, "
        Sql = Sql & "porcrec,impoiva,imporec  FROM factPRO_totales "
        Sql = Sql & " WHERE numserie='" & SerieFraPro
        Sql = Sql & "' AND NUMREGIS= " & NumRegis
        Sql = Sql & " AND anofactu=" & Anofac
        RF.Open Sql, miConexion, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RF.EOF
            
            
            A_Donde = "IVA " & contLinea - 1
            Cad3 = "cuentasr"
            cad2 = DevuelveDesdeBDModulo(ColumnaIVA, "tiposiva", "codigiva", RF!Codigiva, "N", Cad3)
            If cad2 <> "" Then
                Sql = contLinea & ",'" & cad2 & "'," & DocConcAmp
                cad2 = CadenaImporte(True, RF!ImpoIva, Importe0)
                Sql = Sql & "," & cad2 & ","
                Sql = Sql & "NULL,'" & cuenta & "','FRAPRO',0)"
                If Not Importe0 Then
                    miConexion.Execute cad & Sql
                    contLinea = contLinea + 1
                End If
                
                'La de recargo  1-----------------
                If Not IsNull(RF!ImpoREC) Then
                         Sql = contLinea & "," & Cad3 & "," & DocConcAmp
                        'Importes, atencion importes negativos
                        cad2 = CadenaImporte(True, RF!ImpoREC, Importe0)
                        Sql = Sql & "," & cad2 & ","
                        Sql = Sql & "NULL,'" & cuenta & "','FRAPRO',0)"
                        If Not Importe0 Then
                            miConexion.Execute cad & Sql
                            contLinea = contLinea + 1
                        End If
                End If
            
            End If
        
        
        
            RF.MoveNext
        Wend
        
    Else
    
          'Lo dejamos como estaba
    
        '------------------------------------------------------------
        'Cuentas de IVA. Cuenta 1
        A_Donde = "IVA 1"
        Cad3 = "cuentasr"
        cad2 = DevuelveDesdeBDModulo(ColumnaIVA, "tiposiva", "codigiva", RF!tp1facpr, "N", Cad3)
        If cad2 <> "" Then
            Sql = contLinea & ",'" & cad2 & "'," & DocConcAmp
            cad2 = CadenaImporte(True, RF!ti1facpr, Importe0)
            Sql = Sql & "," & cad2 & ","
            Sql = Sql & "NULL,'" & RF!Codmacta & "','FRAPRO',0)"
            If Not Importe0 Then
                miConexion.Execute cad & Sql
                contLinea = contLinea + 1
            End If
            
            'La de recargo  1-----------------
            If Not IsNull(RF!tr1facpr) Then
                     Sql = contLinea & "," & Cad3 & "," & DocConcAmp
                    'Importes, atencion importes negativos
                    cad2 = CadenaImporte(True, RF!tr1facpr, Importe0)
                    Sql = Sql & "," & cad2 & ","
                    Sql = Sql & "NULL,'" & RF!Codmacta & "','FRAPRO',0)"
                    If Not Importe0 Then
                        miConexion.Execute cad & Sql
                        contLinea = contLinea + 1
                    End If
            End If
              
            
                    'MARZO 2015
            '=================================================================================
            'If DBLet(RF!extranje, "N") = 3 Then
            If NoSumaIVA Then
                Cad3 = "cuentarr"
                cad2 = DevuelveDesdeBDModulo("cuentare", "tiposiva", "codigiva", RF!tp1facpr, "N", Cad3)
                
                Cad3 = cad2 & "|" & Cad3 & "|"
                
                
                Sql = contLinea & ",'" & RecuperaValor(Cad3, 1) & "'," & DocConcAmp
                cad2 = CadenaImporte(False, RF!ti1facpr, Importe0)
                Sql = Sql & "," & cad2 & ","
                Sql = Sql & "NULL,'" & RF!Codmacta & "','FRAPRO',0)"
                If Not Importe0 Then
                    miConexion.Execute cad & Sql
                    contLinea = contLinea + 1
                End If
               
                If Not IsNull(RF!tr1facpr) Then
                     Sql = contLinea & "," & RecuperaValor(Cad3, 2) & "," & DocConcAmp
                    'Importes, atencion importes negativos
                    cad2 = CadenaImporte(False, RF!tr1facpr, Importe0)
                    Sql = Sql & "," & cad2 & ","
                    Sql = Sql & "NULL,'" & RF!Codmacta & "','FRAPRO',0)"
                    If Not Importe0 Then
                        miConexion.Execute cad & Sql
                        contLinea = contLinea + 1
                    End If
                End If
            End If
            
    
            
            
        Else
            IntegraLaFacturaProv = "Error leyendo TIPO de IVA: " & Rs!tp1facpr
            RF.Close
            Exit Function
        End If
        
        
        
        
        
        
        '------------------------------------------------------------
        'Cuentas de IVA. Cuenta 2
        A_Donde = "IVA 2"
        If Not IsNull(RF!tp2facpr) Then
            Cad3 = "cuentasr"
            cad2 = DevuelveDesdeBDModulo(ColumnaIVA, "tiposiva", "codigiva", RF!tp2facpr, "N", Cad3)
            If cad2 <> "" Then
                Sql = contLinea & ",'" & cad2 & "'," & DocConcAmp
                'Importes, atencion importes negativos
                cad2 = CadenaImporte(True, RF!ti2facpr, Importe0)
                Sql = Sql & "," & cad2 & ","
                Sql = Sql & "NULL,'" & RF!Codmacta & "','FRAPRO',0)"
                If Not Importe0 Then
                    miConexion.Execute cad & Sql
                    contLinea = contLinea + 1
                End If
                
                'La de recargo  1-----------------
                If Not IsNull(RF!tr2facpr) Then
                         Sql = contLinea & "," & Cad3 & "," & DocConcAmp
                        'Importes, atencion importes negativos
                        cad2 = CadenaImporte(True, RF!tr2facpr, Importe0)
                        Sql = Sql & "," & cad2 & ","
                        Sql = Sql & "NULL,'" & RF!Codmacta & "','FRAPRO',0)"
                        If Not Importe0 Then
                            miConexion.Execute cad & Sql
                            contLinea = contLinea + 1
                        End If
                End If
                
                
                'MARZO 2015
                '=================================================================================
                'If DBLet(RF!extranje, "N") = 3 Then
                If NoSumaIVA Then
                    Cad3 = "cuentarr"
                    cad2 = DevuelveDesdeBDModulo("cuentare", "tiposiva", "codigiva", RF!tp2facpr, "N", Cad3)
                    
                    Cad3 = cad2 & "|" & Cad3 & "|"
                    
                    
                    Sql = contLinea & ",'" & RecuperaValor(Cad3, 1) & "'," & DocConcAmp
                    cad2 = CadenaImporte(False, RF!ti2facpr, Importe0)
                    Sql = Sql & "," & cad2 & ","
                    Sql = Sql & "NULL,'" & RF!Codmacta & "','FRAPRO',0)"
                    If Not Importe0 Then
                        miConexion.Execute cad & Sql
                        contLinea = contLinea + 1
                    End If
                   
                    If Not IsNull(RF!tr2facpr) Then
                         Sql = contLinea & "," & RecuperaValor(Cad3, 1) & "," & DocConcAmp
                        'Importes, atencion importes negativos
                        cad2 = CadenaImporte(False, RF!tr2facpr, Importe0)
                        Sql = Sql & "," & cad2 & ","
                        Sql = Sql & "NULL,'" & RF!Codmacta & "','FRAPRO',0)"
                        If Not Importe0 Then
                            miConexion.Execute cad & Sql
                            contLinea = contLinea + 1
                        End If
                    End If
                End If
                
                
                
                
            Else
                IntegraLaFacturaProv = "Error leyendo TIPO de IVA2: " & Rs!tp2facpr
                RF.Close
                Exit Function
            End If
            
        End If 'De IVA 2
        
        
        
         '------------------------------------------------------------
        'Cuentas de IVA. Cuenta 3
        A_Donde = "IVA 3"
        If Not IsNull(RF!tp3facpr) Then
            Cad3 = "cuentasr"
            cad2 = DevuelveDesdeBDModulo(ColumnaIVA, "tiposiva", "codigiva", RF!tp3facpr, "N", Cad3)
            If cad2 <> "" Then
                Sql = contLinea & ",'" & cad2 & "'," & DocConcAmp
                'Importes, atencion importes negativos
                cad2 = CadenaImporte(True, RF!ti3facpr, Importe0)
                Sql = Sql & "," & cad2 & ","
                Sql = Sql & "NULL,'" & RF!Codmacta & "','FRAPRO',0)"
                If Not Importe0 Then
                    miConexion.Execute cad & Sql
                    contLinea = contLinea + 1
                End If
                
                'La de recargo  1-----------------
                If Not IsNull(RF!tr3facpr) Then
                         Sql = contLinea & "," & Cad3 & "," & DocConcAmp
                        'Importes, atencion importes negativos
                        cad2 = CadenaImporte(True, RF!tr3facpr, Importe0)
                        Sql = Sql & "," & cad2 & ","
                        Sql = Sql & "NULL,'" & RF!Codmacta & "','FRAPRO',0)"
                        If Not Importe0 Then
                            miConexion.Execute cad & Sql
                            contLinea = contLinea + 1
                        End If
                End If
                
                'MARZO 2015
                '=================================================================================
                'If DBLet(RF!extranje, "N") = 3 Then
                If NoSumaIVA Then
                    Cad3 = "cuentarr"
                    cad2 = DevuelveDesdeBDModulo("cuentare", "tiposiva", "codigiva", RF!tp3facpr, "N", Cad3)
                    
                    Cad3 = cad2 & "|" & Cad3 & "|"
                    
                    
                    Sql = contLinea & ",'" & RecuperaValor(Cad3, 1) & "'," & DocConcAmp
                    cad2 = CadenaImporte(False, RF!ti3facpr, Importe0)
                    Sql = Sql & "," & cad2 & ","
                    Sql = Sql & "NULL,'" & RF!Codmacta & "','FRAPRO',0)"
                    If Not Importe0 Then
                        miConexion.Execute cad & Sql
                        contLinea = contLinea + 1
                    End If
                   
                    If Not IsNull(RF!tr3facpr) Then
                         Sql = contLinea & "," & RecuperaValor(Cad3, 1) & "," & DocConcAmp
                        'Importes, atencion importes negativos
                        cad2 = CadenaImporte(False, RF!tr3facpr, Importe0)
                        Sql = Sql & "," & cad2 & ","
                        Sql = Sql & "NULL,'" & RF!Codmacta & "','FRAPRO',0)"
                        If Not Importe0 Then
                            miConexion.Execute cad & Sql
                            contLinea = contLinea + 1
                        End If
                    End If
                End If
                
                
            Else
                IntegraLaFacturaProv = "Error leyendo TIPO de IVA 3: " & Rs!tp3facpr
                RF.Close
                Exit Function
            End If
            
        End If 'De IVA 3
        
    End If 'de nueva contabilizacion lineas
    
    
    
    
    '-------------------------------------
    ' RETENCION
    A_Donde = "Retencion"
    If DatosRetencion <> "" Then
        Sql = contLinea & ",'" & RecuperaValor(DatosRetencion, 1) & "'," & DocConcAmp
        'Importes, atencion importes negativos
        cad2 = CadenaImporte(False, CCur(RecuperaValor(DatosRetencion, 2)), Importe0)
        Sql = Sql & "," & cad2 & ","
        Sql = Sql & "NULL,NULL,'FRAPRO',0)"
       
        miConexion.Execute cad & Sql
        contLinea = contLinea + 1 'Es el contador de lineaapunteshco
    End If
    
    

    '------------------------------------------------------------
    'Las lineas de la factura. Para ello guardaremos algunos datos
    RF.Close
    
    
    
    A_Donde = "Leyendo lineas factura"
    
    If vParamAplic.ContabilidadNueva Then
        Sql = "Select factpro_lineas.codmacta codtbase,baseimpo impbaspr,codccost ,cuentas.codmacta"
        Sql = Sql & " FROM  factpro_lineas,cuentas WHERE  numserie='" & SerieFraPro
        Sql = Sql & "' AND   numregis= " & NumRegis & " AND anofactu=" & Anofac
        Sql = Sql & " AND  factpro_lineas.codmacta = Cuentas.codmacta"
    Else
        Sql = "Select linfactprov.* , codmacta FROM linfactprov,Cuentas WHERE "
        Sql = Sql & "  numregis= " & NumRegis & " AND anofacpr=" & Anofac
        Sql = Sql & " AND  linfactprov.codtbase = Cuentas.codmacta"
    End If
 
    
    RF.Open Sql, miConexion, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Para cada linea insertamos
    A_Donde = "Procesando lineas"
    cad2 = ""
    While Not RF.EOF
        'Importes, atencion importes negativos
        If cad2 = "" Then PrimeraContrapartida = RF!codtbase
        Sql = contLinea & ",'" & RF!codtbase & "'," & DocConcAmp
        cad2 = CadenaImporte(True, RF!impbaspr, Importe0)
        Sql = Sql & "," & cad2 & ","
        If IsNull(RF!CodCCost) Then
            cad2 = "NULL"
        Else
            cad2 = "'" & RF!CodCCost & "'"
        End If
        
        Sql = Sql & cad2 & ",'" & cuenta & "','FRAPRO',0)"
    
        miConexion.Execute cad & Sql
        contLinea = contLinea + 1 'Es el contador de lineaapunteshco
        
        RF.MoveNext
        If Not RF.EOF Then PrimeraContrapartida = ""
    Wend
    RF.Close
    
    
    'AHora viene lo bueno.  MARZO 2006
    'Si el valor fuera true YA lo habria insertado en la cabcera
    A_Donde = "Grabando CC en factura"
    If Not CCenFacturas Then
        If PrimeraContrapartida <> "" Then
            Sql = "UPDATE hlinapu SET ctacontr ='" & PrimeraContrapartida & "'"
            Sql = Sql & " WHERE numdiari = " & DiarioFacturas & " AND fechaent ='" & Format(FechaAsi, FormatoFecha) & "' and numasien = " & Numasiento
            Sql = Sql & " AND linliapu =1 " 'LA PRIMERA LINEA SIEMPRE ES LA DE LA CUENTA
            miConexion.Execute Sql  'Lo hacemos aqui para controlar el error y que no explote
        End If
    End If
    
    'Actualimos en factura, el nº de asiento
    Sql = "UPDATE " & IIf(vParamAplic.ContabilidadNueva, "factpro", "cabfactprov")
    Sql = Sql & " SET numdiari = " & DiarioFacturas & ", fechaent = '" & Format(FechaAsi, FormatoFecha) & "', numasien =" & Numasiento & " WHERE "
    cad2 = "anofacpr"
    If vParamAplic.ContabilidadNueva Then
        Sql = Sql & " numserie='" & SerieFraPro & "' AND "
        cad2 = "anofactu"
    End If
    Sql = Sql & " numregis = " & NumRegis
    Sql = Sql & " AND " & cad2 & "=" & Anofac
   
    miConexion.Execute Sql
    
    
    
    'MAYO 2011   Guarda tb el asiento en el que esta
    Sql = "UPDATE tmpinformes set importeb5 = " & Numasiento & " WHERE codusu = " & vUsu.Codigo & " AND codigo1=" & NumRegis
    ejecutar Sql, True 'Me da lo mismo que de error
    
    
    'Actualizaremos los saldos
    If Not vParamAplic.ContabilidadNueva Then
        A_Donde = "Calculando saldos"
        IntegraLaFacturaProv = CalcularLineasYSaldosFacturas(True)
        
    End If
    
    
    
    
    
    Exit Function
EIntegraLaFactura2:
    IntegraLaFacturaProv = A_Donde & vbCrLf & Err.Description & vbCrLf & Sql
    
    
End Function














'Funciones necesarias
Private Function DevuelveCentroCosteFactura(Cliente As Boolean, LaPrimeraContrapartida As String, numFac As Long, Anofac As Integer, numserie As String) As String
Dim r As ADODB.Recordset
Dim Sql As String
    DevuelveCentroCosteFactura = ""
    If Cliente Then
        
        Sql = "SELECT codccost,numlinea,codtbase FROM linfact"
        Sql = Sql & " WHERE numserie='" & numserie
        Sql = Sql & "' AND codfaccl= " & numFac
        Sql = Sql & " AND anofaccl=" & Anofac
        Sql = Sql & " AND not (codccost is null)"   'El primero k devuelva
        Sql = Sql & " ORDER BY numlinea"
    Else
        
        Sql = "SELECT codccost,numlinea,codtbase FROM linfactprov"
        Sql = Sql & " WHERE numregis = " & numFac
        Sql = Sql & " AND anofacpr=" & Anofac
        Sql = Sql & " AND not (codccost is null)"   'El primero k devuelva
        Sql = Sql & " ORDER BY numlinea"
    End If
    
    
    Set r = New ADODB.Recordset
    r.Open Sql, miConexion, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not r.EOF Then
        If Not IsNull(r.Fields(0)) Then DevuelveCentroCosteFactura = r.Fields(0)
        LaPrimeraContrapartida = r!codtbase
        r.MoveNext
        If Not r.EOF Then LaPrimeraContrapartida = ""
    End If
    r.Close
    Set r = Nothing
End Function

Private Function CadenaImporte(VaAlDebe As Boolean, ByRef Importe As Currency, ElImporteEsCero As Boolean) As String
Dim CadImporte As String

'Si va al debe, pero el importe es negativo entonces va al haber a no ser que la contabilidad admita importes negativos
    If Importe < 0 Then
        If Not Abononeg Then
            VaAlDebe = Not VaAlDebe
            Importe = Abs(Importe)
        End If
    End If
    ElImporteEsCero = (Importe = 0)
    CadImporte = TransformaComasPuntos(CStr(Importe))
    If VaAlDebe Then
        CadenaImporte = CadImporte & ",NULL"
    Else
        CadenaImporte = "NULL," & CadImporte
    End If
End Function




'-------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'//
'//
'//
'//
'Conseguiremos el contador.   0.- ERROR
Private Function ConseguirContador(EjercicioActual As Boolean, ByRef PosibleError As String) As Long
Dim RT As ADODB.Recordset
Dim c1 As Long
Dim f1 As Date

    On Error GoTo Err1
    'Abrimos bloqueando
    ConseguirContador = 0 'ERROR
    Sql = "Select * from contadores WHERE TipoRegi = '0' FOR UPDATE"
    c1 = 0
    Set RT = New ADODB.Recordset
    RT.Open Sql, miConexion, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RT.EOF Then
        If EjercicioActual Then
            c1 = RT!Contado1
        Else
            c1 = RT!Contado2
        End If
    Else
        PosibleError = "No se ecuentra Contador( 0)"
    End If
    RT.Close
    If c1 = 0 Then Exit Function
    c1 = c1 + 1
    
    
    

    

        Sql = "Select numasien from hcabapu where numasien = " & c1
        
        'Las fechas
        If EjercicioActual Then
            f1 = DateAdd("yyyy", -1, FechaFinEjercicio)
            Sql = Sql & " AND fechaent > " & DBSet(f1, "F") 'Mayor estricto
            Sql = Sql & " AND fechaent <= " & DBSet(FechaFinEjercicio, "F")
        Else
            f1 = DateAdd("yyyy", 1, FechaFinEjercicio)
            Sql = Sql & " AND fechaent > " & DBSet(DateAdd("yyyy", 1, FechaFinEjercicio), "F")
            Sql = Sql & " AND fechaent <= " & DBSet(DateAdd("yyyy", 1, f1), "F")
        End If

        
        RT.Open Sql, miConexion, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RT.EOF Then
            'OK. Todo bien
        Else
            
            'Error. Saldremos con el error
            PosibleError = "Ya existe el asiento: " & c1 & "    No se ha podido contabilizar desde el registro de facturas"
            Sql = ""
        End If
        RT.Close
    
    
    If Sql = "" Then
        'Ha habido un error . Ya exise asiento
        c1 = 0
        Exit Function
    End If
    
    'Actualizamos el contador
    Sql = "UPDATE contadores set "
    If EjercicioActual Then
        Sql = Sql & " contado1="
    Else
        Sql = Sql & " contado2="
    End If
    Sql = Sql & c1
    Sql = Sql & " WHERE TipoRegi = '0'"
    miConexion.Execute Sql
    
    ConseguirContador = c1
    
    Exit Function
Err1:
    PosibleError = "Error: " & Err.Number & " : " & Err.Description
    Err.Clear
End Function



'--------------------------------------------------------------------
'//
'//
'//     Calcula los saldos del asiento desde las facturas
'//     Estoes, el asiento esta ya en hco, con lo cual las tablas son de hco
'       Integrar:   TRUE mete en hsaldos
'                   FALSE descontabilziar
'
Private Function CalcularLineasYSaldosFacturas(Integrar As Boolean) As String
    Dim Reparto As Boolean
    Dim RL As Recordset
    Set RL = New ADODB.Recordset
    
    On Error GoTo ECalcularLineasYSaldosFacturas2
    
    CalcularLineasYSaldosFacturas = ""
    
    '------------------------------------------
    'SALDOS
    Sql = "SELECT hlinapu.timporteD AS SD, hlinapu.timporteH AS SH, hlinapu.codmacta"
    Sql = Sql & " From hlinapu"
    Sql = Sql & " WHERE (((hlinapu.numdiari)= " & DiarioFacturas
    Sql = Sql & ") AND ((hlinapu.fechaent)='" & Format(FechaAsi, FormatoFecha) & "'"
    Sql = Sql & ") AND ((hlinapu.numasien)=" & Numasiento
    Sql = Sql & "));"
    
   
    Set RL = New ADODB.Recordset
    RL.Open Sql, miConexion, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        cuenta = RL!Codmacta
        If IsNull(RL!sD) Then
            ImporteD = 0
        Else
            'ImporteD = RL!tImporteD
            ImporteD = RL!sD
        End If
        If IsNull(RL!sH) Then
            ImporteH = 0
        Else
            'ImporteH = RL!tImporteH
            ImporteH = RL!sH
        End If
        
        
        If Not Integrar Then
            ImporteD = -ImporteD
            ImporteH = -ImporteH
        End If
        
        
        If Not CalcularSaldos() Then
            CalcularLineasYSaldosFacturas = "Error calculando saldos "
            RL.Close
            Exit Function
        End If
        
        'Sig
        RL.MoveNext
    Wend
    RL.Close
    
    If Not Autocoste Then
        'NO tiene analitica
        CalcularLineasYSaldosFacturas = ""
        Exit Function
    End If
    
    
    '------------------------------------------
    '       ANALITICA
    Sql = "SELECT hlinapu.timporteD AS SD, hlinapu.timporteH AS SH, hlinapu.codmacta,"
    Sql = Sql & " hlinapu.fechaent, hlinapu.numdiari, hlinapu.numasien, hlinapu.codccost,idsubcos"
    Sql = Sql & " From hlinapu,cabccost WHERE cabccost.codccost=hlinapu.codccost"
    Sql = Sql & " AND hlinapu.numdiari =" & DiarioFacturas
    Sql = Sql & " AND hlinapu.fechaent='" & FechaAsi & "'"
    Sql = Sql & " AND hlinapu.numasien=" & Numasiento
    Sql = Sql & " AND hlinapu.codccost Is Not Null;"
    RL.Open Sql, miConexion, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        cuenta = RL!Codmacta
        CCost = RL!CodCCost
        ImporteD = DBLet(RL!sD, "N")
        ImporteH = DBLet(RL!sH, "N")
        Reparto = (RL!idsubcos = 1)
        If Not CalcularSaldosAnal Then
            CalcularLineasYSaldosFacturas = "Error Saldos analitica"
            Exit Function
        End If
        'Sig
        
        If Reparto Then
            If Not HacerReparto(True) Then
                RL.Close
                CalcularLineasYSaldosFacturas = "Error reparto analitica"
                Exit Function
            End If
        End If
        RL.MoveNext
    Wend
    RL.Close
    CalcularLineasYSaldosFacturas = ""  'Todo bien
    
    Exit Function
ECalcularLineasYSaldosFacturas2:
    CalcularLineasYSaldosFacturas = Err.Description
End Function


'-------------------------------------------------------
'-------------------------------------------------------
'ANALITICA
'-------------------------------------------------------
'-------------------------------------------------------

'Integrar TRUE  meter en hsaldos
'         FALSE quitar apunte
Private Function CalcularSaldos() As Boolean
    Dim i As Integer
    CalcularSaldos = False
    For i = vEmpresa.numnivel To 1 Step -1
        If Not CalcularSaldos1Nivel(i) Then Exit Function
    Next i
    CalcularSaldos = True
End Function


Private Function CalcularSaldos1Nivel(Nivel As Integer) As Boolean
    Dim ImpD As Double
    Dim ImpH As Double
    Dim TD As String
    Dim TH As String
    Dim cta As String
    Dim i As Integer
    
    
    CalcularSaldos1Nivel = False
    i = DigitosNivel(Nivel)
    If i <= 0 Then Exit Function
    
    cta = Mid(cuenta, 1, i)
    Sql = "Select Impmesde,impmesha from hsaldos where "
    Sql = Sql & " Codmacta = '" & cta & "' AND Anopsald = " & Year(FechaAsi) & " AND mespsald = " & Month(FechaAsi)
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, miConexion, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        i = 0   'Nuevo
        ImpD = 0
        ImpH = 0
    Else
        i = 1
        ImpD = Rs.Fields(0)
        ImpH = Rs.Fields(1)
    End If
    Rs.Close
    
    'Acumulamos
    ImpD = ImpD + ImporteD
    ImpH = ImpH + ImporteH
    
    TD = TransformaComasPuntos(CStr(ImpD))
    TH = TransformaComasPuntos(CStr(ImpH))
    If i = 0 Then
        'Nueva insercion
        Sql = "INSERT INTO hsaldos VALUES('" & cta & "'," & Year(FechaAsi) & "," & Month(FechaAsi) & "," & TD & "," & TH & ")"
        Else
        Sql = "UPDATE hsaldos SET Impmesde=" & TD & ", Impmesha = " & TH
        Sql = Sql & " WHERE Codmacta = '" & cta & "' AND Anopsald = " & Year(FechaAsi) & " AND mespsald = " & Month(FechaAsi)
    End If
    miConexion.Execute Sql
    CalcularSaldos1Nivel = True
End Function

Private Function HacerReparto(Actualizar As Boolean) As Boolean
Dim RR As ADODB.Recordset
Dim AD As Currency
Dim AH As Currency
Dim TD As Currency
Dim TH As Currency
Dim b As Boolean

    HacerReparto = False
    TD = ImporteD
    TH = ImporteH
    AD = 0
    AH = 0
    Set RR = New ADODB.Recordset
    Sql = "Select * from linccost WHERE codccost = '" & CCost & "'"
    RR.Open Sql, miConexion, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RR.EOF
        'Cargamos los porcentajes
        CCost = RR!subccost
        ImporteD = (RR!porccost) / 100
        ImporteH = ImporteD
        'Importe porcentajeado
        ImporteD = Round(ImporteD * TD, 2)
        ImporteH = Round(ImporteH * TH, 2)
        'Movemos al sguiente
        RR.MoveNext
        'Por si acaso los decimales quedan sueltos entonces
        'Los valores para el ultimo subcentro de reaparto se obtienen por diferencias
        'con el acumulado
        If RR.EOF Then
            ImporteD = TD - AD
            ImporteH = TH - AH
        Else
            'Acumulo
            AD = AD + ImporteD
            AH = AH + ImporteH
        End If
        If Actualizar Then
            b = CalcularSaldosAnal
        Else
            'NO DESACTUALIZAMOS. Es del copiar pegar
            ''''''''B = CalcularSaldosAnalDesactualizar
        End If
        If Not b Then
            RR.Close
            Exit Function
        End If
    Wend
    RR.Close
    HacerReparto = True
End Function


'Para la ampliacion que va a llevar cada linea de asiento
'Los datos de parametros estan cargados en el REcodset: RS
Private Function ASignarConceptosFacturas() As Boolean
Dim Aux As String
    
    On Error GoTo EASignarConceptosFacturas
    
    ASignarConceptosFacturas = False
    
    'FACTURAS PROVEEDORES
    Aux = "tipoconce"
    Sql = DevuelveDesdeBDModulo("nomconce", "conceptos", "codconce", CStr(Rs!concefpr), "N", Aux)
    If Sql = "" And Aux = "tipoconce" Then Exit Function
    AmpliacionFacurasPro = Rs!concefpr & "|" & Sql & "|"
        
    Aux = "tipoconce"
    Sql = DevuelveDesdeBDModulo("nomconce", "conceptos", "codconce", CStr(Rs!conceapr), "N", Aux)
    If Sql = "" And Aux = "tipoconce" Then Exit Function
    AmpliacionFacurasPro = AmpliacionFacurasPro & Rs!concefpr & "|" & Sql & "|"
    
    'FACTURAS CLIENTES
    Aux = "tipoconce"
    Sql = DevuelveDesdeBDModulo("nomconce", "conceptos", "codconce", CStr(Rs!concefcl), "N", Aux)
    If Sql = "" And Aux = "tipoconce" Then Exit Function
    AmpliacionFacurasCli = Rs!concefcl & "|" & Sql & "|"
        
    Aux = "tipoconce"
    Sql = DevuelveDesdeBDModulo("nomconce", "conceptos", "codconce", CStr(Rs!conceacl), "N", Aux)
    If Sql = "" And Aux = "tipoconce" Then Exit Function
    AmpliacionFacurasCli = AmpliacionFacurasCli & Rs!concefcl & "|" & Sql & "|"
    
    
    
    
    ASignarConceptosFacturas = True
    Exit Function
EASignarConceptosFacturas:
    MsgBox "ASignar Conceptos Facturas" & vbCrLf & Err.Description & vbCrLf & "Aux: " & Aux, vbExclamation
End Function




'-------------------------------------------------------
'-------------------------------------------------------
'ANALITICA
'-------------------------------------------------------
'-------------------------------------------------------

Private Function CalcularSaldosAnal() As Boolean
    
    CalcularSaldosAnal = CalcularSaldos1NivelAnal(NumNiveles)

End Function


Private Function CalcularSaldos1NivelAnal(Nivel As Integer) As Boolean
    Dim ImpD As Currency
    Dim ImpH As Currency
    Dim TD As String
    Dim TH As String
    Dim cta As String
    Dim i As Integer
    
    
    CalcularSaldos1NivelAnal = False
    i = DigitosNivel(Nivel)
    If i < 0 Then Exit Function
    
    cta = Mid(cuenta, 1, i)
    Sql = "Select debccost,habccost from hsaldosanal where "
    Sql = Sql & " codccost='" & CCost & "' AND"
    Sql = Sql & " Codmacta = '" & cta & "' AND anoccost = " & Year(FechaAsi) & " AND mesccost = " & Month(FechaAsi)
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, miConexion, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        i = 0   'Nuevo
        ImpD = 0
        ImpH = 0
    Else
        i = 1
        ImpD = Rs.Fields(0)
        ImpH = Rs.Fields(1)
    End If
    Rs.Close
    'Acumulamos
    ImpD = ImpD + ImporteD
    ImpH = ImpH + ImporteH
    TD = TransformaComasPuntos(CStr(ImpD))
    TH = TransformaComasPuntos(CStr(ImpH))
    If i = 0 Then
        'Nueva insercion
        Sql = "INSERT INTO hsaldosanal(codccost,codmacta,anoccost,mesccost,debccost,habccost)"
        Sql = Sql & " VALUES('" & CCost & "','" & cta & "'," & Year(FechaAsi) & "," & Month(FechaAsi) & "," & TD & "," & TH & ")"
        Else
        Sql = "UPDATE hsaldosanal SET debccost=" & TD & ", habccost = " & TH
        Sql = Sql & " WHERE Codmacta = '" & cta & "' AND Anoccost = " & Year(FechaAsi) & " AND mesccost = " & Month(FechaAsi)
        Sql = Sql & " AND codccost = '" & CCost & "';"
    End If
    miConexion.Execute Sql
    CalcularSaldos1NivelAnal = True
End Function


Private Function DigitosNivel(Nivel) As Integer
    If Nivel = NumNiveles Then
        'Utimo nivel
        
        DigitosNivel = vEmpresa.DigitosUltimoNivel
    Else
        DigitosNivel = DigitosNiveles(Nivel - 1)
    End If
End Function


Private Function DevuelveDesdeBDModulo(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef OtroCampo As String) As String
    Dim Rs As Recordset
    Dim cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBDModulo = ""
    cad = "Select " & kCampo
    If OtroCampo <> "" Then cad = cad & ", " & OtroCampo
    cad = cad & " FROM " & Ktabla
    cad = cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        cad = cad & ValorCodigo
    'Case "T", "F"
    Case Else
        cad = cad & "'" & ValorCodigo & "'"
    
    End Select
    

    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    

    Rs.Open cad, miConexion, adOpenForwardOnly, adLockOptimistic, adCmdText

    
    If Not Rs.EOF Then
        DevuelveDesdeBDModulo = DBLet(Rs.Fields(0))
        If OtroCampo <> "" Then OtroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EDevuelveDesdeBD:
        MsgBox "Devuelve DesdeBD." & vbCrLf & Err.Description, vbExclamation
End Function

Public Function TieneErrores() As Boolean
    TieneErrores = False
    If mvarRealizarContabilizacion Then
        If ColErrores.Count > 0 Then TieneErrores = True
    End If
End Function


Public Sub MuestraErroresContabilizacion()
Dim i As Integer
    
    Sql = ""
    For i = 1 To ColErrores.Count
        Sql = Sql & ColErrores.item(i) & vbCrLf
    Next i
    MsgBox Sql, vbQuestion

End Sub

Private Sub Class_Initialize()
    Set ColErrores = Nothing
    Set miConexion = Nothing
    Set Rs = Nothing
End Sub



'*************************  Veremos cuantas facturas va a pasar. Y podremos ver si tiene los asientos libres
Public Function PreComprobacionNumeroAsiento(Fecha, TotaAsientosGenerar As Integer) As Boolean
Dim c1 As Long
Dim f1 As Date
Dim EjercicioActual As Boolean

    On Error GoTo EPreComprobacionNumeroAsiento

    PreComprobacionNumeroAsiento = False

    Sql = "Select * from contadores WHERE TipoRegi = '0'"
    c1 = 0
    EjercicioActual = Fecha <= FechaFinEjercicio
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, miConexion, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        c1 = 0
        MsgBox "No se ecuentra el contador de asientos (0)", vbExclamation
    
    Else
        If EjercicioActual Then
            c1 = Rs!Contado1
        Else
            c1 = Rs!Contado2
        End If
        c1 = c1 + 1
        
    End If
    Rs.Close
    'No existe el contador
    If c1 = 0 Then Exit Function
    
    'Vemos si existe un asiento entre el numero de contador y los numeros consecutivos de asiento
    Sql = "Select numasien from hcabapu where numasien >= " & c1 & " AND numasien <= " & c1 + TotaAsientosGenerar - 1
    If EjercicioActual Then
        f1 = DateAdd("yyyy", -1, FechaFinEjercicio)
        Sql = Sql & " AND fechaent > " & DBSet(f1, "F") 'Mayor estricto
        Sql = Sql & " AND fechaent <= " & DBSet(FechaFinEjercicio, "F")
    Else
        f1 = DateAdd("yyyy", 1, FechaFinEjercicio)
        Sql = Sql & " AND fechaent > " & DBSet(DateAdd("yyyy", 1, FechaFinEjercicio), "F")
        Sql = Sql & " AND fechaent <= " & DBSet(DateAdd("yyyy", 1, f1), "F")
    End If
    Rs.Open Sql, miConexion, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    If Rs.EOF Then
        'Perfecto. NO deberian dar errores los contadores
        PreComprobacionNumeroAsiento = True
    Else
        'Tenemos problemas.
        Sql = "Ariconta. Error contadores." & vbCrLf
        If TotaAsientosGenerar = 1 Then
            Sql = Sql & "Ya existe el asiento " & c1
        Else
            cuenta = ""
            While Not Rs.EOF
                cuenta = cuenta & ", " & Rs!numasien
                Rs.MoveNext
            Wend
            cuenta = Mid(cuenta, 2)
            Sql = Sql & "Contadores: " & c1 & vbCrLf
            Sql = Sql & "Nº asientos:" & TotaAsientosGenerar & vbCrLf
            Sql = Sql & "Asiento ya existentes: " & cuenta
            cuenta = ""
            
        End If
        Sql = Sql & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then PreComprobacionNumeroAsiento = True
    End If
    Rs.Close
    Exit Function
EPreComprobacionNumeroAsiento:
    MsgBox "Precomprobar ARICONTA." & vbCrLf & Err.Description, vbExclamation
End Function



'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'
'   Eliminar factura cliente en contabilidad
'
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
Public Function EliminarFRACLIcontab(EliminaElCobro As Boolean, FechaFac As Date) As Boolean
Dim RF As ADODB.Recordset
Dim Sql As String
Dim Existe As Boolean

    On Error GoTo EEliminarFRACLIcontab
    EliminarFRACLIcontab = False
    
    
    
  
    

    
    'cobros
    If EliminaElCobro Then
        Sql = " numserie='" & Serie & "' AND codfaccl=" & NumeroFactura
        Sql = Sql & " AND fecfaccl='" & Format(FechaFac, FormatoFecha) & "'"
        miConexion.Execute "Delete from scobro WHERE " & Sql
    End If
    
    Set RF = New ADODB.Recordset
    Sql = "SELECT numasien,fechaent,numdiari,codfaccl FROM cabfact"
    Sql = Sql & " WHERE numserie='" & Serie
    Sql = Sql & "' AND codfaccl= " & NumeroFactura
    Sql = Sql & " AND anofaccl= " & Anofac
    RF.Open Sql, miConexion, adOpenForwardOnly, adLockPessimistic, adCmdText
    DiarioFacturas = -1
    Existe = False
    If Not RF.EOF Then
        Existe = True
        If Not IsNull(RF!numdiari) Then
            DiarioFacturas = RF!numdiari
            FechaAsi = RF!FechaEnt
            Numasiento = RF!numasien
        End If
    End If
    RF.Close
    
    If Existe Then
    
    
        If DiarioFacturas > 0 Then
            'OK. HAY que eliminar el apunte
    
            Sql = CalcularLineasYSaldosFacturas(False)
            If Sql <> "" Then
                MsgBox Sql, vbExclamation
                Exit Function
            End If
            
            Sql = " WHERE numdiari= " & DiarioFacturas
            Sql = Sql & " AND fechaent='" & Format(FechaAsi, FormatoFecha) & "'"
            Sql = Sql & " AND numasien=" & Numasiento
            
            miConexion.Execute "DELETE FROM hlinapu " & Sql
            miConexion.Execute "DELETE FROM hcabapu " & Sql
        End If
    
    
    
    
        Sql = " numserie='" & Serie & "' AND codfaccl=" & NumeroFactura
        Sql = Sql & " AND anofaccl = " & Anofac

        'Lineas
        miConexion.Execute "Delete from linfact WHERE " & Sql

        'cabecera
        miConexion.Execute "Delete from cabfact WHERE " & Sql

        
    End If
    
    EliminarFRACLIcontab = True
    Exit Function
EEliminarFRACLIcontab:
    MuestraError Err.Number, "Eliminar FRACLI contab"
End Function
