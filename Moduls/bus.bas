Attribute VB_Name = "bus"


'NOTA: en este m�dul, adem�s, n'hi han funcions generals que no siguen de formularis (molt b�)
Option Explicit


Public Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Definicion Conexi�n a BASE DE DATOS
'---------------------------------------------------
'Conexi�n a la BD Ariagro de la empresa
Public conn As ADODB.Connection

'Conexi�n a la BD de Usuarios
Public ConnUsuarios As ADODB.Connection

'Conexi�n a la BD de Contabilidad de la empresa conectada
Public ConnConta As ADODB.Connection

'Conexi�n a la BD de Contabilidad de otra empresa distinta a la conectada
Public ConnAuxCon As ADODB.Connection

'Conexi�n a la BD de Aridoc de la empresa conectada
Public ConnAridoc As ADODB.Connection

'Conexi�n a la BD de Ariges si tiene suministros
Public ConnAriges As ADODB.Connection

'Conexi�n a la BD Ariagro de la campa�a anterior
Public ConnCAnt As ADODB.Connection

'Conexion a la base de datos de indefa
Public ConnIndefa As ADODB.Connection

'Conexion a la base de datos de sqlserver de castelduc
Public CnnSqlServer As ADODB.Connection

'Conexion a la base de datos de dbf de Monasterios
Public CnnSqlMonast As ADODB.Connection

'Conexi�n a la BD de App
Public ConnApp As ADODB.Connection


'[Monica] 06/09/2010: su�stituida esta constante por el parametro vParamAplic.Faneca
Public Const cFaneca As Single = 0.0833 ' hectareas


'Que conexion a base de datos se va a utilizar
Public Const cAgro As Byte = 1 'trabajaremos con conn (conexion a BD Ariagro)
Public Const cConta As Byte = 2 'trabajaremos con connConta (cxion a BD Contabilidad)
Public Const cAridoc As Byte = 3 'trabajaremos con connAridoc (cxion a BD Aridoc)
Public Const cAriges As Byte = 4 'trabajaremos con connAriges (conexion a BD Suministros)
Public Const cAPP As Byte = 5 'trabajaremos con connApp (cxion a BD Aripush)

'LOG de acciones relevantes
Public LOG As cLOG   'Se instancia , se ejecuta LOG.insertar y se elimina :LOG=nothing   Ver ejemplo borre facturas


'Definicion de clases de la aplicaci�n
'-----------------------------------------------------
Public vEmpresa As Cempresa  'Los datos de la empresa
Public vParam As Cparametros  'Parametros Generales de la Empresa (nombre, direc.,...
Public vParamAplic As CParamAplic   'parametros de la aplicacion
Public vSesion As CSesion   'Los datos del usuario que hizo login

Public vUsu As Usuario  'Datos usuario
Public vConfig As Configuracion

Public miRsAux As ADODB.Recordset

Public Const vbFPTransferencia = 1

Public ProcesoCorrecto As Boolean

'Definicion de FORMATOS
'---------------------------------------------------
Public FormatoFecha As String
Public FormatoFechaHora As String
Public FormatoHora As String
Public FormatoImporte As String 'Decimal(12,2)
Public FormatoPrecio As String 'Decimal(8,3)
Public FormatoCantidad As String 'Decimal(10,2)
Public FormatoPorcen As String 'Decimal(5,2) 'Porcentajes
Public FormatoExp As String  'Expedientes
Public FormatoDescuento As String 'Decimal(4,2)

Public FormatoDec10d2 As String 'Decimal(10,2)
Public FormatoDec10d3 As String 'Decimal(10,3)
Public FormatoDec5d4 As String 'Decimal(5,4)
Public FormatoDec8d4 As String 'Decimal(8,4)
Public FormatoDec6d4 As String 'Decimal(6,4)
Public FormatoDec8d6 As String 'Decimal(8,6)
Public FormatoDec6d2 As String 'Decimal(6,2)
Public FormatoDec10d4 As String 'Decimal(10,4)
Public FormatoDec6d3 As String 'Decimal(6,3)

Public FIni As String
Public FFin As String

'Public FormatoKms As String 'Decimal(8,4)


Public teclaBuscar As Integer 'llamada desde prismaticos

Public Const vbLightBlue = &HFEEFDA
Public Const vbErrorColor = &HDFE1FF      '&HFFFFC0
Public Const vbMoreLightBlue = &HFEFBD8   ' azul clarito

Public CadenaDesdeOtroForm As String

'Global para n� de registro eliminado
Public NumRegElim  As Long

'publica para almacenar control cambios en registros de formularios
'se utiliza en InsertarCambios
Public CadenaCambio As String
Public ValorAnterior As String

Public MensError As String
Public FormularioOK As String

'Para algunos campos de texto suletos controlarlos
'Public miTag As CTag

'Variable para saber si se ha actualizado algun asiento
'Public AlgunAsientoActualizado As Boolean
'Public TieneIntegracionesPendientes As Boolean

'Public miRsAux As ADODB.Recordset

Public AnchoLogin As String  'Para fijar los anchos de columna

' **** DATOS DEL LOGIN ****
'Public CodEmple As Integer
'Public codAgenc As Integer
'Public codEmpre As Integer
'Public codGrupo As Integer
'Public claEmpre As Integer
'Public TipEmple As Integer
'Public areEmple As Integer
' *************************

Public ardDB As BaseDatos ' este es la base de datos que soportar� aridoc

Public dbAriagro As BaseDatos ' base de datos para la grabacion del chivato
Public ObsFactura As String ' Observaciones de la factura de anticipo/liquidacion


Public Const SerieFraPro = "1"
Public Const SerieFraPro2 = "2"
Public vvTrabajadores As String
Public DireccionAyuda As String

Public vCadBusqueda As String


Public ContabilizadoOk As Boolean

Public ResultadoFechaContaOK As Byte
Public MensajeFechaOkConta As String

' constante a partir de los cuales son socios de la otra cooperativa
' coopic y picassent
Public Const cMaxSocio = 100000
Public Const cMaxCampo = 10000000
Public Const cMaxCapa = 1000

Dim frmLect As frmPOZLecturasMonast
Public EsMonasterios As Boolean
Public Semaforo As Byte

'[Monica]08/05//2018: cambio de los mensajes
Public RespuestaMsgBox As Integer  'Para el msgbox del codejock


Public ValorDevuelto As Long


'Inicio Aplicaci�n
Public Sub Main()

'     If App.PrevInstance Then
'        MsgBox "Ariagro ya se esta ejecutando", vbExclamation
'        End
'     End If
     
'cambiarlo por lo de abajo, cuando quiera quitar el login

       AbrirConexionMonasterios



       Load frmIdentifica
       CadenaDesdeOtroForm = ""
       

       'Necesitaremos el archivo arifon.dat
       frmIdentifica.Show vbModal

       If CadenaDesdeOtroForm = "" Then
            'NO se ha identificado
            Set conn = Nothing
            End
       End If

       CadenaDesdeOtroForm = ""
       frmLogin.Show vbModal
       If CadenaDesdeOtroForm = "" Then
            'No ha seleccionado nonguna empresa
            Set conn = Nothing
            End
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
'++++
'    Set vConfig = New Configuracion
'    If vConfig.Leer = 1 Then
'
'         MsgBox "MAL CONFIGURADO", vbCritical
'         End
'         Exit Sub
'    End If
'
'     AbrirConexionUsuarios
'
'    Set vUsu = New Usuario
'    vUsu.Leer "root"
'    vUsu.CadenaConexion = "ariagro1"
'++++



'        LeerEmpresa 'Carga los Datos de la empresa
        'Carga los Datos B�sicos de la empresa
        LeerDatosEmpresa



        'Cerramos la conexion
        conn.Close

        'Abre la conexi�n a BDatos:Ariagro
        If AbrirConexion() = False Then
            MsgBox "La aplicaci�n no puede continuar sin acceso a los datos. ", vbCritical
            End
        Else
            'Carga Parametros Generales y Contables de la empresa
            LeerParametros
        End If
                
        '[Monica]16/10/2017: para el caso de Monasterios control de que si la fecha fin de campa�a es anterior o igual a la de
        '                    a hoy salimos de la aplicacion
        If vParamAplic.Cooperativa = 17 Then
            If CDate(vParam.FecFinCam) <= Now Then
                MsgBox "Error en las fechas de campa�a. La aplicaci�n no se ejecutar�.", vbExclamation
                End
            End If
        End If
                
                
                
' solo en el mantenimiento de contadores
'[Monica]03/12/2012: De momento pongo aqui la conexion a indefa
'        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
'            If AbrirConexionIndefa() = False Then
'                MsgBox "La aplicaci�n no puede continuar sin acceso a los datos de indefa. ", vbCritical
'                End
'            End If
'        End If
                
'--monica: aqui la conexion a contabilidad se har� en el momento dependiendo de la seccion
'        'Abrir conexi�n a la BDatos de Contabilidad para acceder a
'        'Tablas: Cuentas, Tipos IVA
'        If vParamAplic.NumeroConta <> 0 Then
'            If AbrirConexionConta() = False Then
'                MsgBox "La aplicaci�n no puede continuar sin acceso a los datos. ", vbCritical
'                End
'            End If
'        End If
'
'        'Carga los Niveles de cuentas de Contabilidad de la empresa y las fechasINICIO y FIN
'        If vParamAplic.NumeroConta <> 0 Then LeerNivelesEmpresa



        'Gestionar el nombre del PC para la asignacion de PC en el entorno de red
        GestionaPC

        'Otras acciones
        OtrasAcciones
        
        Screen.MousePointer = vbHourglass
        Load MDIppal
        Screen.MousePointer = vbDefault
        MDIppal.Show
     
     
     
        '[Monica]08/09/2017: para el caso de que se ejecute el programa de lecturas en la tableta
        If vParamAplic.Cooperativa = 17 And vUsu.Nivel = 2 Then
            Set frmLect = New frmPOZLecturasMonast
            frmLect.Show vbModal
            Set frmLect = Nothing
            
        Else
     
            If vParamAplic.ContabilidadNueva And (vUsu.Nivel = 0 Or vUsu.Nivel = 1) Then FrasPendientesContabilizar True
     
        End If
     
'     'obric la conexio
'    If AbrirConexionAriagro("root", "aritel") = False Then
'        MsgBox "La aplicaci�n no puede continuar sin acceso a los datos. ", vbCritical
'        End
'    End If
'
'    Load frmIdentifica
'    'CadenaDesdeOtroForm = ""
'
'    'Necesitaremos el archivo login.dat
'    frmIdentifica.Show
    
End Sub






Private Sub prueba()
Dim Rs As ADODB.Recordset
Dim SQL As String

    SQL = "select * from leccont.dbf"
    
    
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, CnnSqlMonast, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    While Not Rs.EOF
'
'
'        Rs.MoveNext
'    Wend
'    Set Rs = Nothing
End Sub


Public Sub FrasPendientesContabilizar(EsRecoleccion As Boolean)
Dim SQL As String
Dim Sql2 As String
Dim SqlBd As String
Dim Rs As ADODB.Recordset
Dim SqlValues As String
Dim SQLinsert As String
Dim RsBd As ADODB.Recordset
Dim BBDD As String

Dim frmMens As frmMensajes

Dim vSeccion As CSeccion

    On Error GoTo eFrasPendientesContabilizar



    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Set vSeccion = Nothing
            Exit Sub
        End If
        
'[Monica]29/06/2017: sale lo que haya en esta campa�a
'        If vParamAplic.Cooperativa <> 12 And vParamAplic.Cooperativa <> 9 And vParamAplic.Cooperativa <> 14 Then
'            If vEmpresa.BDAriagro <> "ariagro" Then Exit Sub '@@@
'        End If
    
        SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
        conn.Execute SQL
    
        BBDD = vEmpresa.BDAriagro
        
        SQLinsert = "insert into " & BBDD & ".tmpinformes (codusu, nombre1,codigo1,nombre2,fecha1, text1, nombre3, importe1) "
    
    
    If EsRecoleccion Then

        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        
            SQL = " select " & vUsu.Codigo & ",concat(rrecibpozos.codtipom,' ',usuarios.stipom.nomtipom)  tipofact,5, '','','" & BBDD & "' aa, count(*), sum(totalfact) from " & BBDD & ".rrecibpozos, " & BBDD & ".rsocios, usuarios.stipom where contabilizado = 0 and rrecibpozos.codsocio = rsocios.codsocio "
            SQL = SQL & " and rrecibpozos.codtipom = stipom.codtipom "
            If vEmpresa.TieneSII Then
                SQL = SQL & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
            End If
            SQL = SQL & " group by 1,2,3,4,5,6 order by 1,2,3,4,5,6"
        
        Else

            SQL = " select " & vUsu.Codigo & ",'Facturas ADV' tipofact, 0, concat(advfacturas.codtipom,right(concat('0000000',numfactu),7)),fecfactu,'" & BBDD & "' aa, nomsocio, totalfac  from " & BBDD & ".advfacturas where intconta = 0 "
            If vEmpresa.TieneSII Then
                SQL = SQL & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
            End If
            SQL = SQL & " union "
            SQL = SQL & " select " & vUsu.Codigo & ",'Fras Varias Cliente' tipofact,1, concat(fvarcabfact.codtipom,right(concat('0000000',numfactu),7)),fecfactu,'" & BBDD & "' aa, rsocios.nomsocio, totalfac from " & BBDD & ".fvarcabfact, " & BBDD & ".rsocios  where intconta = 0 and fvarcabfact.codsocio = rsocios.codsocio "
            If vEmpresa.TieneSII Then
                SQL = SQL & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
            End If
            SQL = SQL & " union "
            SQL = SQL & " select " & vUsu.Codigo & ",'Fras Varias Cliente' tipofact,1, concat(fvarcabfact.codtipom,right(concat('0000000',numfactu),7)),fecfactu,'" & BBDD & "' aa, clientes.nomclien, totalfac from " & BBDD & ".fvarcabfact, " & BBDD & ".clientes  where intconta = 0 and fvarcabfact.codclien = clientes.codclien "
            If vEmpresa.TieneSII Then
                SQL = SQL & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
            End If
            
            SQL = SQL & " union "
            SQL = SQL & " select " & vUsu.Codigo & ",'Fras Varias Proveedor' tipofact,2, concat(fvarcabfactpro.codtipom,right(concat('0000000',numfactu),7)),fecfactu,'" & BBDD & "' aa, rsocios.nomsocio, totalfac from " & BBDD & ".fvarcabfactpro, " & BBDD & ".rsocios where intconta = 0 and fvarcabfactpro.codsocio = rsocios.codsocio "
            If vEmpresa.TieneSII Then
                SQL = SQL & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
            End If
            SQL = SQL & " union "
            SQL = SQL & " select " & vUsu.Codigo & ",'Facturas Socio' tipofact,3, concat(rfactsoc.codtipom,right(concat('0000000',numfactu),7)),fecfactu,'" & BBDD & "' aa, rsocios.nomsocio, totalfac from " & BBDD & ".rfactsoc, " & BBDD & ".rsocios where contabilizado = 0 and rfactsoc.codsocio = rsocios.codsocio"
            If vEmpresa.TieneSII Then
                SQL = SQL & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
            End If
            SQL = SQL & " union "
            SQL = SQL & " select " & vUsu.Codigo & ",'Fras Transportistas' tipofact,4, concat(codtipom,right(concat('0000000',numfactu),7)),fecfactu,'" & BBDD & "' aa, rtransporte.nomtrans, totalfac from " & BBDD & ".rfacttra, " & BBDD & ".rtransporte where contabilizado = 0 and rfacttra.codtrans = rtransporte.codtrans "
            If vEmpresa.TieneSII Then
                SQL = SQL & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
            End If
            SQL = SQL & " union "
            SQL = SQL & " select " & vUsu.Codigo & ",'Facturas Pozos' tipofact,5, concat(rrecibpozos.codtipom,right(concat('0000000',numfactu),7)),fecfactu,'" & BBDD & "' aa, rsocios.nomsocio, totalfact from " & BBDD & ".rrecibpozos, " & BBDD & ".rsocios where contabilizado = 0 and rrecibpozos.codsocio = rsocios.codsocio "
            If vEmpresa.TieneSII Then
                SQL = SQL & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
            End If
            SQL = SQL & " union "
            SQL = SQL & " select " & vUsu.Codigo & ",'Facturas Retirada' tipofact,6, concat(rbodfacturas.codtipom,right(concat('0000000',numfactu),7)),fecfactu,'" & BBDD & "' aa, rsocios.nomsocio, totalfac from " & BBDD & ".rbodfacturas, " & BBDD & ".rsocios where intconta = 0 and rbodfacturas.codsocio = rsocios.codsocio "
            If vEmpresa.TieneSII Then
                SQL = SQL & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
            End If
        
        
        
        End If
                
        '[Monica]04/03/2019: eliminamos el blqueo de tablas que producia, poniendolo en un insert values
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        SqlValues = ""
        While Not Rs.EOF
            SqlValues = SqlValues & ",(" & DBSet(Rs.Fields(0), "N") & "," & DBSet(Rs.Fields(1), "T") & "," & DBSet(Rs.Fields(2), "N") & "," & DBSet(Rs.Fields(3), "T") & "," & DBSet(Rs.Fields(4), "F") & "," & DBSet(Rs.Fields(5), "T") & "," & DBSet(Rs.Fields(6), "T") & "," & DBSet(Rs.Fields(7), "N") & ")"
            
            Rs.MoveNext
        Wend
        Set Rs = Nothing

        If SqlValues <> "" Then
            SqlValues = " values " & Mid(SqlValues, 2)
            conn.Execute SQLinsert & SqlValues
        End If
'        conn.Execute SQLinsert & Sql
            
    Else
    
        SQL = " select " & vUsu.Codigo & ",'Facturas Cliente' tipofact, 0,concat(facturas.codtipom,right(concat('0000000',numfactu),7)),fecfactu,'" & BBDD & "' aa, nomclien, totalfac from " & BBDD & ".facturas, " & BBDD & ".clientes  where intconta = 0 and facturas.codclien = clientes.codclien "
        If vEmpresa.TieneSII Then
            SQL = SQL & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
        End If
        SQL = SQL & " union "
        SQL = SQL & " select " & vUsu.Codigo & ",'Facturas Cliente a Socios' tipofact, 1,concat(facturassocio.codtipom,right(concat('0000000',numfactu),7)),fecfactu,'" & BBDD & "' aa, nomsocio, totalfac from " & BBDD & ".facturassocio, " & BBDD & ".rsocios where intconta = 0 and facturassocio.codsocio = rsocios.codsocio "
        If vEmpresa.TieneSII Then
            SQL = SQL & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
        End If
            
        SQL = SQL & " union "
        SQL = SQL & " select " & vUsu.Codigo & ",'Facturas Proveedor' tipofact, 2,numfactu,fecrecep,'" & BBDD & "' aa, nomprove, totalfac  from " & BBDD & ".scafpc where intconta = 0 "
        If vEmpresa.TieneSII Then
            SQL = SQL & " and fecrecep >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecrecep <= " & DBSet(DateAdd("d", -1, Now), "F")
        End If
        SQL = SQL & " union "
        SQL = SQL & " select " & vUsu.Codigo & ",'Facturas Transportistas' tipofact,3, numfactu,fecrecep,'" & BBDD & "' aa, nomtrans, totalfac from " & BBDD & ".tcafpc where intconta = 0 "
        If vEmpresa.TieneSII Then
            SQL = SQL & " and fecrecep >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecrecep <= " & DBSet(DateAdd("d", -1, Now), "F")
        End If
        SQL = SQL & " union "
        SQL = SQL & " select " & vUsu.Codigo & ",'Facturas Compras' tipofact,4, numfactu,fecrecep,'" & BBDD & "' aa, proveedor.nomprove, totalfac from " & BBDD & ".facturascom, " & BBDD & ".proveedor  where intconta = 0 and facturascom.codprove = proveedor.codprove "
        If vEmpresa.TieneSII Then
            SQL = SQL & " and fecrecep >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecrecep <= " & DBSet(DateAdd("d", -1, Now), "F")
        End If
        

        '[Monica]04/03/2019: eliminamos el blqueo de tablas que producia, poniendolo en un insert values
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        SqlValues = ""
        While Not Rs.EOF
            SqlValues = SqlValues & ",(" & DBSet(Rs.Fields(0), "N") & "," & DBSet(Rs.Fields(1), "T") & "," & DBSet(Rs.Fields(2), "N") & "," & DBSet(Rs.Fields(3), "T") & "," & DBSet(Rs.Fields(4), "F") & "," & DBSet(Rs.Fields(5), "T") & "," & DBSet(Rs.Fields(6), "T") & "," & DBSet(Rs.Fields(7), "N") & ")"
            
            Rs.MoveNext
        Wend
        Set Rs = Nothing

        If SqlValues <> "" Then
            SqlValues = " values " & Mid(SqlValues, 2)
            conn.Execute SQLinsert & SqlValues
        End If

'        conn.Execute SQLinsert & Sql
    End If
        
'[Monica]29/06/2017: quito lo de la campa�a anterior
'
'        If vParamAplic.Cooperativa <> 12 And vParamAplic.Cooperativa <> 9 And vParamAplic.Cooperativa <> 14 And vParamAplic.Cooperativa <> 16 Then
'
'
'            Dim vCampAnt As CCampAnt
'
'            Set vCampAnt = New CCampAnt
'
'        ' si solo tenemos que buscar en la campa�a anterior
'            If vCampAnt.Leer = 0 Then
'
'                SqlBd = "SHOW DATABASES like 'ariagro' "
'                Set RsBd = New ADODB.Recordset
'                RsBd.Open SqlBd, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'                While Not RsBd.EOF
'    '               If Trim(DBLet(RsBd.Fields(0).Value)) <> vEmpresa.BDAriagro And Trim(DBLet(RsBd.Fields(0).Value)) <> "" And InStr(1, DBLet(RsBd.Fields(0).Value), "ariagroutil") = 0 Then
'                    If Trim(DBLet(RsBd.Fields(0).Value)) = vCampAnt.BaseDatos Then
'
'                        Sql2 = Replace(Sql, BBDD, DBLet(RsBd.Fields(0).Value))
'
'                        conn.Execute SqlInsert & Sql2
'                    End If
'
'                    RsBd.MoveNext
'                Wend
'
'                Set RsBd = Nothing
'
'            End If
'
'            Set vCampAnt = Nothing
'
'        End If
        
        SQL = "select codusu,nombre1,codigo1,nombre2,nombre3,fecha1,text1,importe1  from tmpinformes where codusu = " & vUsu.Codigo '& " order by 6,5 "
        
        If TotalRegistrosConsulta(SQL) > 0 Then
            Set frmMens = New frmMensajes
            
            If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                frmMens.OpcionMensaje = 70
            Else
                frmMens.OpcionMensaje = 68
            End If
            frmMens.cadena = SQL
            frmMens.Show vbModal
        
            Set frmMens = Nothing
        End If
        
        CerrarConexionConta
        Exit Sub
        
    End If
    Set vSeccion = Nothing
    
eFrasPendientesContabilizar:
    MuestraError Err.Number, "Facturas Pendientes de Integrar a Contabilidad", Err.Description
End Sub


Public Function UltimaFechaCorrectaSII(DiasAVisoSII As Integer, FechaPresentacion As Date) As Date
Dim DiaSemanaPresen As Integer
Dim DiaSemanaUltimoDiaPresentar As Integer
Dim F As Date
Dim Resta As Integer
    
    If DiasAVisoSII > 5 Then
        UltimaFechaCorrectaSII = DateAdd("d", -DiasAVisoSII, FechaPresentacion)
    Else
        DiaSemanaPresen = WeekDay(FechaPresentacion, vbMonday)
        If DiaSemanaPresen >= 6 Then
            'Si presento el sabado o el domingo tengo mas dias
            If DiaSemanaPresen = 6 Then
                Resta = DiasAVisoSII
            Else
                Resta = DiasAVisoSII + 1
            End If
        Else
            F = DateAdd("d", -DiasAVisoSII, FechaPresentacion)
            DiaSemanaUltimoDiaPresentar = WeekDay(F, vbMonday)
            If DiaSemanaUltimoDiaPresentar > DiaSemanaPresen Then
                Resta = DiasAVisoSII + 2
            Else
                'Directamente la resta son 4
                Resta = DiasAVisoSII
            End If
        End If
        UltimaFechaCorrectaSII = DateAdd("d", -Resta, FechaPresentacion)
    End If
    UltimaFechaCorrectaSII = Format(UltimaFechaCorrectaSII, "dd/mm/yyyy")

End Function




Public Function ComprovaVersio() As Boolean
  
'    Dim RS2 As Recordset
'    Dim cad2 As String
'    Dim major_ul As Integer
'    Dim minor_ul As Integer
'    Dim revis_ul As Integer
'
'    ComprovaVersio = False
'
'    cad2 = "SELECT * FROM ulversio"
'
'    Set RS2 = New ADODB.Recordset
'    RS2.Open cad2, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    If Not RS2.EOF Then
'        major_ul = RS2.Fields!major_ul
'        minor_ul = RS2.Fields!minor_ul
'        revis_ul = RS2.Fields!revis_ul
'    Else
'        MsgBox "Error al consultar la �ltima versi�n disponible", vbCritical
'        'ulVersio = False
'        Exit Function
'    End If
'
'    RS2.Close
'    Set RS2 = Nothing
'
'    If (App.Major <> major_ul) Or (App.Minor <> minor_ul) Or (App.Revision <> revis_ul) Then
'        ComprovaVersio = True
'    End If
'
'    Exit Function
    
End Function

'espera els segon que li digam
Public Function espera(Segundos As Single)
    Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function


Public Function AbrirConexionConta() As Boolean
'Abre

Dim Cad As String
Dim BBDD As String

On Error GoTo EAbrirConexion

    
    AbrirConexionConta = False
    Set ConnConta = Nothing
    Set ConnConta = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnConta.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
    If vParamAplic.ContabilidadNueva Then
        BBDD = "ariconta" & vParamAplic.NumeroConta
    Else
        BBDD = "conta" & vParamAplic.NumeroConta
    End If
                        
    If vParamAplic.ServidorConta = "" Then vParamAplic.ServidorConta = vConfig.SERVER
    
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;"
    
    '[Monica]29/06/2018: cambiado, ahora volvemos a dar el odbc
    If vParamAplic.ContabilidadNueva Then
        Cad = Cad & "DATA SOURCE=Ariconta6;"
    Else
        Cad = Cad & "DATA SOURCE=vConta;"
    End If
    
    Cad = Cad & "DATABASE=" & BBDD & ";SERVER=" & vParamAplic.ServidorConta
    Cad = Cad & ";UID=" & vParamAplic.UsuarioConta
    Cad = Cad & ";PWD=" & vParamAplic.PasswordConta

'[Monica]29/06/2018:
'    '---- Laura: 29/09/2006
'    cad = cad & ";PORT=3306;OPTION=3;STMT=;"
'    '----
    '++monica: tema de vista
    Cad = Cad & ";Persist Security Info=true"
    
    
    ConnConta.ConnectionString = Cad
    ConnConta.Open
    ConnConta.Execute "Set AUTOCOMMIT = 1"
    AbrirConexionConta = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n.", Err.Description
End Function

Public Function AbrirConexionConta2(NumeroConta As Integer) As Boolean
'Abre
Dim BBDD As String
Dim Cad As String
On Error GoTo EAbrirConexion
    
    AbrirConexionConta2 = False
    Set ConnConta = Nothing
    Set ConnConta = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnConta.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
    
    If vParamAplic.ContabilidadNueva Then
        BBDD = "ariconta" & NumeroConta
    Else
        BBDD = "conta" & NumeroConta
    End If
                        
    If vParamAplic.ServidorConta = "" Then vParamAplic.ServidorConta = vConfig.SERVER
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;"
    
'[Monica]29/06/2018: indicamos el ODBC
    If vParamAplic.ContabilidadNueva Then
        Cad = Cad & "DATA SOURCE=Ariconta6;"
    Else
        Cad = Cad & "DATA SOURCE=vConta;"
    End If
    
    Cad = Cad & "DATABASE=" & BBDD & ";SERVER=" & vParamAplic.ServidorConta
    Cad = Cad & ";UID=" & vParamAplic.UsuarioConta
    Cad = Cad & ";PWD=" & vParamAplic.PasswordConta
    
'[Monica]29/06/2018: quitamos lo del puerto
'    '---- Laura: 29/09/2006
'    cad = cad & ";PORT=3306;OPTION=3;STMT=;"
'    '----
    '++monica: tema de vista
    Cad = Cad & ";Persist Security Info=true"
    
    
    ConnConta.ConnectionString = Cad
    ConnConta.Open
    ConnConta.Execute "Set AUTOCOMMIT = 1"
    AbrirConexionConta2 = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n.", Err.Description
End Function

Public Function AbrirConexionAridoc(Usuario As String, Pass As String) As Boolean
'Abre
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionAridoc = False
    Set ConnAridoc = Nothing
    Set ConnAridoc = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnAridoc.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
'    cad = "DSN=Aridoc;DESC=MySQL ODBC 3.51 Driver DSN;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
'    '++monica:tema del vista
'    cad = cad & "Persist Security Info=true"
'    '++
    
    '[Monica]29/06/2018:se a�ade el odbc DATA SOURCE=Aridoc
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE=Aridoc;DATABASE=aridoc;SERVER=" & vConfig.SERVER
    Cad = Cad & ";UID=" & vConfig.User
    Cad = Cad & ";PWD=" & vConfig.password
'++monica: tema del vista
    Cad = Cad & ";Persist Security Info=true"
 
    ConnAridoc.ConnectionString = Cad
    ConnAridoc.Open
    ConnAridoc.Execute "Set AUTOCOMMIT = 1"
    AbrirConexionAridoc = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n Aridoc.", Err.Description
End Function


Public Function AbrirConexionAriges() As Boolean
'Abre
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionAriges = False
    Set ConnAriges = Nothing
    Set ConnAriges = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnAriges.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
    '[Monica]24/02/2012: Por Moixent si no tienen servidor en parametros coge el de ConfigAgro
    If vParamAplic.ServidorConta = "" Then vParamAplic.ServidorConta = vConfig.SERVER
    
    '[Monica]29/06/2018: a�adido el DATA SOURCE=vODBC
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE=vAriagro;DATABASE=" & vParamAplic.BDAriges & ";SERVER=" & vParamAplic.ServidorConta
    Cad = Cad & ";UID=" & vParamAplic.UsuarioConta
    Cad = Cad & ";PWD=" & vParamAplic.PasswordConta
'    '---- Laura: 29/09/2006
'    cad = cad & ";PORT=3306;OPTION=3;STMT=;"
'    '----
    '++monica: tema de vista
    Cad = Cad & ";Persist Security Info=true"
    
    ConnAriges.ConnectionString = Cad
    ConnAriges.Open
    ConnAriges.Execute "Set AUTOCOMMIT = 1"
    AbrirConexionAriges = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n Ariges.", Err.Description
End Function







'Public Function AbrirConexionApp() As Boolean
''Abre
'Dim cad As String
'On Error GoTo EAbrirConexion
'
'
'    AbrirConexionApp = False
'    Set ConnApp = Nothing
'    Set ConnApp = New Connection
''    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
'    ConnApp.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
'
'    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=aripush;SERVER=" & vConfig.SERVER & ";"
'    cad = cad & ";UID=" & vConfig.User
'    cad = cad & ";PWD=" & vConfig.password
''++monica: tema del vista
'    cad = cad & ";Persist Security Info=true"
'
'    ConnAridoc.ConnectionString = cad
'    ConnAridoc.Open
'    ConnAridoc.Execute "Set AUTOCOMMIT = 1"
'    AbrirConexionApp = True
'    Exit Function
'
'EAbrirConexion:
'    MuestraError Err.Number, "Abrir conexi�n APP Ariagro.", Err.Description
'End Function
'
'

Public Function CerrarConexionConta()
  'Cerramos la conexion con BD: Contabilidad
  On Error Resume Next
   ConnConta.Close
   If Err.Number <> 0 Then Err.Clear
End Function

Public Function CerrarConexionConta2()
  'Cerramos la conexion con BD: Contabilidad
  On Error Resume Next
   ConnConta.Close
   If Err.Number <> 0 Then Err.Clear
End Function


Public Function CerrarConexionUsuarios()
  'Cerramos la conexion con BD: Usuarios
  On Error Resume Next
   conn.Close
   If Err.Number <> 0 Then Err.Clear
End Function

Public Function CerrarConexionAridoc()
  'Cerramos la conexion con BD: Aridoc
  On Error Resume Next
   ConnAridoc.Close
   If Err.Number <> 0 Then Err.Clear
End Function


Public Function CerrarConexionAriges()
  'Cerramos la conexion con BD: Ariges (suministros)
  On Error Resume Next
   ConnAriges.Close
   If Err.Number <> 0 Then Err.Clear
End Function



Public Function CerrarConexionCampAnterior()
  'Cerramos la conexion con BD: Usuarios
  On Error Resume Next
   ConnCAnt.Close
   If Err.Number <> 0 Then Err.Clear
End Function


Public Function CerrarConexionIndefa()
'Cerramos la conexion con BD: Indefa (pozos)
On Error Resume Next
   ConnIndefa.Close
   If Err.Number <> 0 Then Err.Clear
End Function

Public Function CerrarConexionSqlSERVER()
  'Cerramos la conexion con BD: sqlserver que se utiliza en Castelduc
  On Error Resume Next
   CnnSqlServer.Close
   If Err.Number <> 0 Then Err.Clear
End Function

Public Function CerrarConexionApp()
  'Cerramos la conexion con BD: Aripush
  On Error Resume Next
   ConnApp.Close
   If Err.Number <> 0 Then Err.Clear
End Function



Public Function LeerDatosEmpresa()
 'Crea instancia de la clase Cempresa con los valores en
 'Tabla: ArigesEmpresa
 'BDatos: Usuarios
 
        Set vEmpresa = New Cempresa
        If vEmpresa.LeerDatos = 1 Then
            MsgBox "No se han podido cargar datos empresa (BD:usuarios). Debe configurar la aplicaci�n.", vbExclamation
            Set vEmpresa = Nothing
        End If
            
End Function


Public Sub LeerParametros()
'Crea instancia de la clase Cempresa con los valores en
'Tabla: Empresas
'BDatos: PTours y Conta
 Dim devuelve As String
 
    'Parametros Generales
    Set vParam = New Cparametros
    If vParam.Leer() = 1 Then
        devuelve = "No se han podido cargar los Par�metros Generales.(empresas)" & vbCrLf
        MsgBox devuelve & " Debe configurar la aplicaci�n.", vbExclamation
        Set vParam = Nothing
    End If
    
    ' ### [Monica] 06/09/2006
    ' a�adido
    Set vParamAplic = New CParamAplic
    If vParamAplic.Leer = 1 Then
        MsgBox "No se han podido cargar los Par�metros de la Aplicaci�n(sparam). Debe configurar la aplicaci�n.", vbExclamation

        Set vParamAplic = Nothing
        Exit Sub
    End If

End Sub


Public Function PonerDatosPpal()
'    If Not vEmpresa Is Nothing Then
'        MDIppal.Caption = "AriAgro" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  Empresa: " & vEmpresa.nomEmpre
'    End If
    
    If vParam Is Nothing Then
        MDIppal.Caption = "AriAgro - Gesti�n Comercial" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & " FALTA CONFIGURAR"
    Else
        MDIppal.Caption = "AriAgro - Gesti�n Comercial" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  Empresa: " & vParam.NombreEmpresa & _
                  " - Campa�a: " & vParam.FecIniCam & " - " & vParam.FecFinCam & "   -  Usuario: " & vUsu.Nombre
    End If

    
    If Err.Number <> 0 Then MuestraError Err.Description, "Poniendo datos de la pantalla principal", Err.Description
End Function

    

Public Sub MuestraError(numero As Long, Optional cadena As String, Optional Desc As String)
    Dim Cad As String
    Dim Aux As String
    
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    Cad = "Se ha producido un error: " & vbCrLf
    If cadena <> "" Then
        Cad = Cad & vbCrLf & cadena & vbCrLf & vbCrLf
    End If
    'Numeros de errores que contolamos
    If conn.Errors.Count > 0 Then
        ControlamosError Aux
        conn.Errors.Clear
    Else
        Aux = ""
    End If
    If Aux <> "" Then Desc = Aux
    If Desc <> "" Then Cad = Cad & vbCrLf & Desc & vbCrLf & vbCrLf
    If Aux = "" Then Cad = Cad & "N�mero: " & numero & vbCrLf & "Descripci�n: " & Error(numero)
    MsgBox Cad, vbExclamation
End Sub

Public Function DBSet(vData As Variant, Tipo As String, Optional EsNulo As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
Dim Cad As String

        If IsNull(vData) Then
            DBSet = ValorNulo
            Exit Function
        End If

        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    If vData = "" Then
                        If EsNulo = "N" Then
                            DBSet = "''"
                        Else
                            DBSet = ValorNulo
                        End If
                    Else
                        Cad = (CStr(vData))
                        NombreSQL Cad
                        DBSet = "'" & Cad & "'"
                    End If
                    
                Case "N"    'Numero
                    If vData = "" Or vData = 0 Then
                        If EsNulo <> "" Then
                            If EsNulo = "S" Then
                                DBSet = ValorNulo
                            Else
                                DBSet = 0
                            End If
                        Else
                            DBSet = 0
                        End If
                    Else
                        Cad = CStr(ImporteFormateado(CStr(vData)))
                        DBSet = TransformaComasPuntos(Cad)
                    End If
                    
                Case "F"    'Fecha
'                     '==David
''                    DBLet = "0:00:00"
'                     '==Laura
                    If vData = "" Then
                        If EsNulo = "S" Then
                            DBSet = ValorNulo
                        Else
                            DBSet = "'1900-01-01'"
                        End If
                    Else
                        DBSet = "'" & Format(vData, FormatoFecha) & "'"
                    End If
                    
                Case "FH" 'Fecha/Hora
                    If vData = "" Then
                        If EsNulo = "S" Then DBSet = ValorNulo
                    Else
                        DBSet = "'" & Format(vData, "yyyy-mm-dd hh:mm:ss") & "'"
                    End If
                    
                Case "H" 'Hora
                    If vData = "" Then
                    Else
                        DBSet = "'" & Format(vData, "hh:mm:ss") & "'"
                    End If
                    
                Case "B"  'Boolean
                    If vData Then
                        DBSet = 1
                    Else
                        DBSet = 0
                    End If
            End Select
        End If
End Function

Public Function DBLetMemo(vData As Variant) As Variant
    On Error Resume Next
    
    DBLetMemo = vData
    
    
    
    If Err.Number <> 0 Then
        Err.Clear
        DBLetMemo = ""
    End If
End Function



Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
'Para cuando recupera Datos de la BD
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    DBLet = ""
                Case "N"    'Numero
                    DBLet = 0
                Case "F"    'Fecha
                     '==David
'                    DBLet = "0:00:00"
                     '==Laura
'                     DBLet = "0000-00-00"
                      DBLet = ""
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo antes de bloquear
'   Prepara la conexion para bloquear
Public Sub PreparaBloquear()
    conn.Execute "commit"
    conn.Execute "set autocommit=0"
End Sub

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo despues de un bloque
'   Prepara la conexion para bloquear
Public Sub TerminaBloquear()
    conn.Execute "commit"
    conn.Execute "set autocommit=1"
End Sub

'///////////////////////////////////////////////////////////////
'
'   Cogemos un numero formateado: 1.256.256,98  y deevolvemos 1256256,98
'   Tiene que venir num�rico
Public Function ImporteFormateado(Importe As String) As Currency
Dim i As Integer

    If Importe = "" Then
        ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            i = InStr(1, Importe, ".")
            If i > 0 Then Importe = Mid(Importe, 1, i - 1) & Mid(Importe, i + 1)
        Loop Until i = 0
        ImporteFormateado = Importe
    End If
End Function

' ### [Monica] 11/09/2006
Public Function ImporteSinFormato(cadena As String) As String
Dim i As Integer
'Quitamos puntos
Do
    i = InStr(1, cadena, ".")
    If i > 0 Then cadena = Mid(cadena, 1, i - 1) & Mid(cadena, i + 1)
Loop Until i = 0
ImporteSinFormato = TransformaPuntosComas(cadena)
End Function



'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(cadena As String) As String
Dim i As Integer
    Do
        i = InStr(1, cadena, ",")
        If i > 0 Then
            cadena = Mid(cadena, 1, i - 1) & "." & Mid(cadena, i + 1)
        End If
    Loop Until i = 0
    TransformaComasPuntos = cadena
End Function

'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef cadena As String)
Dim J As Integer
Dim i As Integer
Dim Aux As String
    J = 1
    Do
        i = InStr(J, cadena, "'")
        If i > 0 Then
            Aux = Mid(cadena, 1, i - 1) & "\"
            cadena = Aux & Mid(cadena, i)
            J = i + 2
        End If
    Loop Until i = 0
End Sub

Public Function EsFechaOKString(ByRef T As String) As Boolean
Dim Cad As String
    
    Cad = T
    If InStr(1, Cad, "/") = 0 Then
        If Len(T) = 8 Then
            Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5)
        Else
            If Len(T) = 6 Then Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5)
        End If
    End If
    If IsDate(Cad) Then
        EsFechaOKString = True
        T = Format(Cad, "dd/mm/yyyy")
    Else
        EsFechaOKString = False
    End If
End Function

Public Function DevNombreSQL(cadena As String) As String
Dim J As Integer
Dim i As Integer
Dim Aux As String
    J = 1
    Do
        i = InStr(J, cadena, "'")
        If i > 0 Then
            Aux = Mid(cadena, 1, i - 1) & "\"
            cadena = Aux & Mid(cadena, i)
            J = i + 2
        End If
    Loop Until i = 0
    DevNombreSQL = cadena
End Function


Public Function DevuelveDesdeBD(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef OtroCampo As String) As String
    Dim Rs As Recordset
    Dim Cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    Cad = "Select " & kCampo
    If OtroCampo <> "" Then Cad = Cad & ", " & OtroCampo
    Cad = Cad & " FROM " & Ktabla
    Cad = Cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        Cad = Cad & ValorCodigo
    Case "T", "F"
        Cad = Cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
    
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        DevuelveDesdeBD = DBLet(Rs.Fields(0))
        If OtroCampo <> "" Then OtroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EDevuelveDesdeBD:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function



''Este metodo sustituye a DevuelveDesdeBD
''Funciona para claves primarias formadas por 2 campos
'Public Function DevuelveDesdeBDnew(vBD As Byte, Ktabla As String, kCampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef otroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String) As String
''IN: vBD --> Base de Datos a la que se accede
'Dim RS As Recordset
'Dim cad As String
'Dim Aux As String
'
'On Error GoTo EDevuelveDesdeBDnew
'    DevuelveDesdeBDnew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
'    cad = "Select " & kCampo
'    If otroCampo <> "" Then cad = cad & ", " & otroCampo
'    cad = cad & " FROM " & Ktabla
'    cad = cad & " WHERE " & Kcodigo1 & " = "
'    If tipo1 = "" Then tipo1 = "N"
'    Select Case tipo1
'        Case "N"
'            'No hacemos nada
'            If IsNumeric(valorCodigo1) Then
'                cad = cad & Val(valorCodigo1)
'            Else
'                MsgBox "El campo debe ser num�rico.", vbExclamation
'                DevuelveDesdeBDnew = "Error"
'                Exit Function
'            End If
'        Case "T", "F"
'            cad = cad & "'" & valorCodigo1 & "'"
'        Case Else
'            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
'            Exit Function
'    End Select
'
'    If KCodigo2 <> "" Then
'        cad = cad & " AND " & KCodigo2 & " = "
'        If tipo2 = "" Then tipo2 = "N"
'        Select Case tipo2
'        Case "N"
'            'No hacemos nada
'            If ValorCodigo2 = "" Then
'                cad = cad & "-1"
'            Else
'                cad = cad & Val(ValorCodigo2)
'            End If
'        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
'        Case "F"
'            cad = cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
'        Case Else
'            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
'            Exit Function
'        End Select
'    End If
'
'
'    'Creamos el sql
'    Set RS = New ADODB.Recordset
'
'    Select Case vBD
'        Case cAgro 'vBD=1: PlannerTours
'            RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'        Case cConta 'BD 2: Contabilidad
'            RS.Open cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
'        Case 3 'vBD=3: contabilidad distinta a la de la empresa conectada
'            RS.Open cad, ConnAuxCon, adOpenForwardOnly, adLockOptimistic, adCmdText
'    End Select
''    If vBD = cAgro Then 'vBD=1: PlannerTours
''        RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
''    ElseIf vBD = cConta Then  'BD 2: Contabilidad
''        RS.Open cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
''    End If
'
'    If Not RS.EOF Then
'        DevuelveDesdeBDnew = DBLet(RS.Fields(0))
'        If otroCampo <> "" Then otroCampo = DBLet(RS.Fields(1))
'    End If
'    RS.Close
'    Set RS = Nothing
'    Exit Function
'
'EDevuelveDesdeBDnew:
'        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
'End Function


'LAURA
'Este metodo sustituye a DevuelveDesdeBD
'Funciona para claves primarias formadas por 3 campos
Public Function DevuelveDesdeBDNew(vBD As Byte, Ktabla As String, kCampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef OtroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String, Optional KCodigo3 As String, Optional ValorCodigo3 As String, Optional tipo3 As String) As String
'IN: vBD --> Base de Datos a la que se accede
Dim Rs As Recordset
Dim Cad As String
Dim Aux As String
    
On Error GoTo EDevuelveDesdeBDnew
    DevuelveDesdeBDNew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
    Cad = "Select " & kCampo
    If OtroCampo <> "" Then Cad = Cad & ", " & OtroCampo
    Cad = Cad & " FROM " & Ktabla
    If Kcodigo1 <> "" Then
        Cad = Cad & " WHERE " & Kcodigo1 & " = "
        If tipo1 = "" Then tipo1 = "N"
    Select Case tipo1
        Case "N"
            'No hacemos nada
            Cad = Cad & Val(valorCodigo1)
        '[Monica]28/11/2018: nuevo tipo de datos T1 para no concatenar asteriscos
        Case "T", "T1"
            Cad = Cad & DBSet(valorCodigo1, "T")
        Case "F"
            Cad = Cad & DBSet(valorCodigo1, "F")
        Case Else
            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
            Exit Function
    End Select
    End If
    
    If KCodigo2 <> "" Then
        Cad = Cad & " AND " & KCodigo2 & " = "
        If tipo2 = "" Then tipo2 = "N"
        Select Case tipo2
        Case "N"
            'No hacemos nada
            If ValorCodigo2 = "" Then
                Cad = Cad & "-1"
            Else
                Cad = Cad & Val(ValorCodigo2)
            End If
            
        '[Monica]28/11/2018: nuevo tipo de datos T1 para no concatenar asteriscos
        Case "T", "T1"
'            cad = cad & "'" & ValorCodigo2 & "'"
            Cad = Cad & DBSet(ValorCodigo2, "T")
        Case "F"
            Cad = Cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    If KCodigo3 <> "" Then
        Cad = Cad & " AND " & KCodigo3 & " = "
        If tipo3 = "" Then tipo3 = "N"
        Select Case tipo3
        Case "N"
            'No hacemos nada
            If ValorCodigo3 = "" Then
                Cad = Cad & "-1"
            Else
                Cad = Cad & Val(ValorCodigo3)
            End If
        '[Monica]28/11/2018: nuevo tipo de datos para no concatenar asteriscos
        Case "T", "T1"
            Cad = Cad & "'" & ValorCodigo3 & "'"
        Case "F"
            Cad = Cad & "'" & Format(ValorCodigo3, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo3 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    
    Select Case vBD
        Case cAgro 'BD 1: Ariagro
            Rs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Case cConta 'BD 2: Conta
            Rs.Open Cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
        Case cAridoc 'BD 3: Aridoc
            Rs.Open Cad, ConnAridoc, adOpenForwardOnly, adLockOptimistic, adCmdText
        Case cAriges 'BD 4: Ariges (suministros)
            Rs.Open Cad, ConnAriges, adOpenForwardOnly, adLockOptimistic, adCmdText
        Case cAPP 'BD 5: Aripush (app de ariagro)
            Rs.Open Cad, ConnApp, adOpenForwardOnly, adLockOptimistic, adCmdText
        
    End Select
    
    If Not Rs.EOF Then
        DevuelveDesdeBDNew = DBLet(Rs.Fields(0))
        If OtroCampo <> "" Then OtroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
EDevuelveDesdeBDnew:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function




'CESAR
Public Function DevuelveDesdeBDnew2(kBD As Integer, kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional num As Byte, Optional ByRef OtroCampo As String) As String
Dim Rs As Recordset
Dim Cad As String
Dim Aux As String
Dim v_aux As Integer
Dim campo As String
Dim Valor As String
Dim tip As String

On Error GoTo EDevuelveDesdeBDnew2
DevuelveDesdeBDnew2 = ""

Cad = "Select " & kCampo
If OtroCampo <> "" Then Cad = Cad & ", " & OtroCampo
Cad = Cad & " FROM " & Ktabla

If Kcodigo <> "" Then Cad = Cad & " where "

For v_aux = 1 To num
    campo = RecuperaValor(Kcodigo, v_aux)
    Valor = RecuperaValor(ValorCodigo, v_aux)
    tip = RecuperaValor(Tipo, v_aux)
        
    Cad = Cad & campo & "="
    If tip = "" Then Tipo = "N"
    
    Select Case tip
            Case "N"
                'No hacemos nada
                Cad = Cad & Valor
            'Monica]28/11/2018: nuevo tipo de datos para no concatenar asteriscos
            Case "T", "T1", "F"
                Cad = Cad & "'" & Valor & "'"
            Case Else
                MsgBox "Tipo : " & tip & " no definido", vbExclamation
            Exit Function
    End Select
    
    If v_aux < num Then Cad = Cad & " AND "
  Next v_aux

'Creamos el sql
Set Rs = New ADODB.Recordset
Select Case kBD
    Case 1
        Rs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
End Select

If Not Rs.EOF Then
    DevuelveDesdeBDnew2 = DBLet(Rs.Fields(0))
    If OtroCampo <> "" Then OtroCampo = DBLet(Rs.Fields(1))
Else
     If OtroCampo <> "" Then OtroCampo = ""
End If
Rs.Close
Set Rs = Nothing
Exit Function
EDevuelveDesdeBDnew2:
    MuestraError Err.Number, "Devuelve DesdeBDnew2.", Err.Description
End Function


Public Function EsEntero(Texto As String) As Boolean
Dim i As Integer
Dim C As Integer
Dim L As Integer
Dim res As Boolean

    res = True
    EsEntero = False

    If Not IsNumeric(Texto) Then
        res = False
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            i = InStr(L, Texto, ".")
            If i > 0 Then
                L = i + 1
                C = C + 1
            End If
        Loop Until i = 0
        If C > 1 Then res = False
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                i = InStr(L, Texto, ",")
                If i > 0 Then
                    L = i + 1
                    C = C + 1
                End If
            Loop Until i = 0
            If C > 1 Then res = False
        End If
        
    End If
        EsEntero = res
End Function

Public Function TransformaPuntosComas(cadena As String) As String
    Dim i As Integer
    Do
        i = InStr(1, cadena, ".")
        If i > 0 Then
            cadena = Mid(cadena, 1, i - 1) & "," & Mid(cadena, i + 1)
        End If
        Loop Until i = 0
    TransformaPuntosComas = cadena
End Function

Public Sub InicializarFormatos()
    FormatoFecha = "yyyy-mm-dd"
    FormatoHora = "hh:mm:ss"
'    FormatoFechaHora = "yyyy-mm-dd hh:mm:ss"
    FormatoImporte = "#,###,###,##0.00"  'Decimal(12,2)
    FormatoPrecio = "##,##0.000"  'Decimal(8,3) antes decimal(10,4)
'    FormatoCantidad = "##,###,##0.00"   'Decimal(10,2)
    FormatoPorcen = "##0.00" 'Decima(5,2) para porcentajes
    
    FormatoDec10d2 = "##,###,##0.00"   'Decimal(10,2)
    FormatoDec10d3 = "##,###,##0.000"   'Decimal(10,3)
    FormatoDec5d4 = "0.0000"   'Decimal(5,4)
    FormatoDec8d4 = "###0.0000" ' Decimal(8,4)
    FormatoDec6d4 = "#0.0000" ' Decimal(6,4)
    FormatoDec8d6 = "#0.000000" ' Decimal(8,6)
    FormatoDec6d2 = "#,##0.00" ' decimal(6,2)
    FormatoDec10d4 = "###,##0.0000" ' decimal(10,4)
    FormatoDec6d3 = "##0.000" ' decimal(6,3)
    
    FormatoExp = "0000000000"
'    FormatoKms = "#,##0.00##" 'Decimal(8,4)
End Sub


Public Sub AccionesCerrar()
'cosas que se deben hacen cuando finaliza la aplicacion
    On Error Resume Next
    
    'cerrar clases q estan abiertas durante la ejecucion
    Set vEmpresa = Nothing
    Set vSesion = Nothing
    
'    Set vParam = Nothing
'    Set vParamAplic = Nothing
'    Set vParamConta = Nothing
    
    
    'Cerrar Conexiones a bases de datos
    conn.Close
    ConnConta.Close
    Set conn = Nothing
    Set ConnConta = Nothing
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function OtrosPCsContraAplicacion() As String
Dim MiRS As Recordset
Dim Cad As String
Dim Equipo As String

    Set MiRS = New ADODB.Recordset
    Cad = "show processlist"
    MiRS.Open Cad, conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not MiRS.EOF
        If MiRS.Fields(3) = vUsu.CadenaConexion Then
            Equipo = MiRS.Fields(2)
            'Primero quitamos los dos puntos del puerot
            NumRegElim = InStr(1, Equipo, ":")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            'El punto del dominio
            NumRegElim = InStr(1, Equipo, ".")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            Equipo = UCase(Equipo)
            
            If Equipo <> vUsu.PC Then
                    If Equipo <> "LOCALHOST" Then
                        If InStr(1, Cad, Equipo & "|") = 0 Then Cad = Cad & Equipo & "|"
                    End If
            End If
        End If
        'Siguiente
        MiRS.MoveNext
    Wend
    NumRegElim = 0
    MiRS.Close
    Set MiRS = Nothing
    OtrosPCsContraAplicacion = Cad
End Function


Public Function UsuariosConectados() As Boolean
Dim i As Integer
Dim Cad As String
Dim metag As String
Dim SQL As String
Cad = OtrosPCsContraAplicacion
UsuariosConectados = False
If Cad <> "" Then
    UsuariosConectados = True
    i = 1
    metag = "Los siguientes PC's est�n conectados a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
    Do
        SQL = RecuperaValor(Cad, i)
        If SQL <> "" Then
            metag = metag & "    - " & SQL & vbCrLf
            i = i + 1
        End If
    Loop Until SQL = ""
    MsgBox metag, vbExclamation
End If
End Function

'Usuario As String, Pass As String --> Directamente el usuario
Public Function AbrirConexion() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexion = False
    Set conn = Nothing
    Set conn = New Connection
    'Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
    '[Monica]29/06/2018:se a�ade el odbc DATA SOURCE=vAriagro
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE=vAriagro;DATABASE=" & vUsu.CadenaConexion & ";SERVER=" & vConfig.SERVER & ";"
    Cad = Cad & ";UID=" & vConfig.User
    Cad = Cad & ";PWD=" & vConfig.password
'++monica: tema del vista
    Cad = Cad & ";Persist Security Info=true"
    
    
'    cad = "DSN=vAriagro;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & vUsu.CadenaConexion & ";UID=" & vConfig.User & ";PASSWORD=" & vConfig.password & ";PORT=3306;OPTION=3;STMT=;"
'    cad = cad & ";Persist Security Info=true"
'
    
    conn.ConnectionString = Cad
    conn.Open
    conn.Execute "Set AUTOCOMMIT = 1"
    AbrirConexion = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n.", Err.Description
End Function


Public Function AbrirConexionCampAnterior(BaseDatos As String) As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionCampAnterior = False
    Set ConnCAnt = Nothing
    Set ConnCAnt = New Connection
    'Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnCAnt.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
'    cad = "DSN=vAriagro;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & BaseDatos & ";UID=" & vConfig.User & ";PASSWORD=" & vConfig.password & ";PORT=3306;OPTION=3;STMT=;"
'    cad = cad & ";Persist Security Info=true"
    
    '[Monica]29/06/2018: a�adido el DATA SOURCE=vODBC
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE=vAriagro;DATABASE=" & BaseDatos & ";SERVER=" & vConfig.SERVER
    Cad = Cad & ";UID=" & vConfig.User
    Cad = Cad & ";PWD=" & vConfig.password
'++monica: tema del vista
    Cad = Cad & ";Persist Security Info=true"
    
    ConnCAnt.ConnectionString = Cad
    ConnCAnt.Open
    ConnCAnt.Execute "Set AUTOCOMMIT = 1"
    
    AbrirConexionCampAnterior = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n Campa�a Anterior.", Err.Description
End Function


Public Function AbrirConexionIndefa() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionIndefa = False
    Set ConnIndefa = Nothing
    Set ConnIndefa = New Connection
    'Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnIndefa.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
                       
'    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & vUsu.CadenaConexion & ";SERVER=" & vConfig.SERVER & ";"
'    cad = cad & ";UID=" & vConfig.User
'    cad = cad & ";PWD=" & vConfig.password
''++monica: tema del vista
'    cad = cad & ";Persist Security Info=true"
    
    
    
'Conexion con la tabla de datos de Indefa
    Cad = "Provider=MSDASQL.1;Persist Security Info=false;User ID=escalona;Data Source=vIndefa"
    Cad = Cad & ";Persist Security Info=true"

    ConnIndefa.ConnectionString = Cad
    ConnIndefa.Open
    
    If vParamAplic.Cooperativa = 8 Then
    
    Else
        ConnIndefa.Execute "Set AUTOCOMMIT = 1"
    End If
    AbrirConexionIndefa = True
    Exit Function
EAbrirConexion:
'    MuestraError Err.Number, "Abrir conexi�n.", Err.Description
End Function

Public Function AbrirConexionSqlSERVER(Servidor As Byte) As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion
    
    AbrirConexionSqlSERVER = False
    Set CnnSqlServer = Nothing
    Set CnnSqlServer = New Connection
    
    'ProductionV50;"
    If Servidor = 0 Then
        Cad = "Driver={SQL Server};Server=" & Trim(vParamAplic.SqlServer) & ",1433;Database=" & Trim(vParamAplic.SqlServerBD) & ";" & _
                "Uid=client;Pwd=client;"
    Else
        Cad = "Driver={SQL Server};Server=" & Trim(vParamAplic.SqlServer1) & ",1433;Database=" & Trim(vParamAplic.SqlServerBD) & ";" & _
                "Uid=client;Pwd=client;"
    End If

    CnnSqlServer.ConnectionString = Cad
    CnnSqlServer.Open
    AbrirConexionSqlSERVER = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n SqlSERVER.", Err.Description
End Function


Public Function LeerEmpresaParametros()
    'Abrimos la empresa
    Set vEmpresa = New Cempresa
    If vEmpresa.LeerDatos = 1 Then
        MsgBox "No se han podido cargar datos empresa. Debe configurar la aplicaci�n.", vbExclamation
        Set vEmpresa = Nothing
    End If
    
    If Not (vEmpresa Is Nothing) Then 'And Not (vParamAplic Is Nothing) Then
        CadenaDesdeOtroForm = ""
    End If
        
End Function

Private Sub GestionaPC()
CadenaDesdeOtroForm = ComputerName
If CadenaDesdeOtroForm <> "" Then
    FormatoFecha = DevuelveDesdeBD("codpc", "usuarios.pcs", "nompc", CadenaDesdeOtroForm, "T")
    If FormatoFecha = "" Then
        NumRegElim = 0
        FormatoFecha = "Select max(codpc) from usuarios.pcs"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open FormatoFecha, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            NumRegElim = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        NumRegElim = NumRegElim + 1
        '[Monica] 07/10/2009 > 32000 pq daba error en el numero de pcs activo
        If NumRegElim > 32000 Then '30 Then
            MsgBox "Error en numero de PC's activos. Demasiados PC en BD. Llame a soporte t�cnico.", vbCritical
            End
        End If
        FormatoFecha = "INSERT INTO usuarios.pcs (codpc, nompc) VALUES (" & NumRegElim & ", '" & CadenaDesdeOtroForm & "')"
        conn.Execute FormatoFecha
    End If
End If
End Sub


Private Sub OtrasAcciones()
On Error Resume Next

    FormatoFecha = "yyyy-mm-dd"
    FormatoFechaHora = "yyyy-mm-dd hh:mm:ss"
    FormatoImporte = "#,###,###,##0.00"
    FormatoCantidad = "##,###,##0.00"   'Decimal(10,2)
    FormatoDescuento = "#0.00" 'Decima(4,2)

    teclaBuscar = 43

    DireccionAyuda = "http://help-ariagro.ariadnasw.com/"


    InicializarFormatos

    'Borramos uno de los archivos temporales
    If Dir(App.Path & "\ErrActua.txt") <> "" Then Kill App.Path & "\ErrActua.txt"
    
    
    'Borramos tmp bloqueos
    'Borramos temporal
    CadenaDesdeOtroForm = OtrosPCsContraContabiliad
    NumRegElim = Len(CadenaDesdeOtroForm)
    If NumRegElim = 0 Then
        CadenaDesdeOtroForm = ""
    Else
        CadenaDesdeOtroForm = " WHERE codusu = " & vUsu.Codigo
    End If
    conn.Execute "Delete from zBloqueos " & CadenaDesdeOtroForm
    CadenaDesdeOtroForm = ""
    NumRegElim = 0
    
    
End Sub

Public Function OtrosPCsContraContabiliad() As String
Dim MiRS As Recordset
Dim Cad As String
Dim Equipo As String

    Set MiRS = New ADODB.Recordset
    Cad = "show processlist"
    MiRS.Open Cad, conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not MiRS.EOF
        If MiRS.Fields(3) = vUsu.CadenaConexion Then
            Equipo = MiRS.Fields(2)
            'Primero quitamos los dos puntos del puerot
            NumRegElim = InStr(1, Equipo, ":")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            'El punto del dominio
            NumRegElim = InStr(1, Equipo, ".")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            Equipo = UCase(Equipo)
            
            If Equipo <> vUsu.PC Then
                    If Equipo <> "LOCALHOST" Then
                        If InStr(1, Cad, Equipo & "|") = 0 Then Cad = Cad & Equipo & "|"
                    End If
            End If
        End If
        'Siguiente
        MiRS.MoveNext
    Wend
    NumRegElim = 0
    MiRS.Close
    Set MiRS = Nothing
    OtrosPCsContraContabiliad = Cad
End Function

Public Function ComprobarEmpresaBloqueada(CodUsu As Long, ByRef Empresa As String) As Boolean
Dim Cad As String

ComprobarEmpresaBloqueada = False

'Antes de nada, borramos las entradas de usuario, por si hubiera kedado algo
conn.Execute "Delete from usuarios.vbloqbd where codusu=" & CodUsu

'Ahora comprobamos k nadie bloquea la BD
Cad = DevuelveDesdeBD("codusu", "usuarios.vbloqbd", "conta", Empresa, "T")
If Cad <> "" Then
    'En teoria esta bloqueada. Puedo comprobar k no se haya kedado el bloqueo a medias
    
    Set miRsAux = New ADODB.Recordset
    Cad = "show processlist"
    miRsAux.Open Cad, conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not miRsAux.EOF
        If miRsAux.Fields(3) = Empresa Then
            Cad = miRsAux.Fields(2)
            miRsAux.MoveLast
        End If
    
        'Siguiente
        miRsAux.MoveNext
    Wend
    
    If Cad = "" Then
        'Nadie esta utilizando la aplicacion, luego se puede borrar la tabla
        conn.Execute "Delete from usuarios.vbloqbd where conta ='" & Empresa & "'"
        
    Else
        MsgBox "BD bloqueada.", vbCritical
        ComprobarEmpresaBloqueada = True
    End If
End If

conn.Execute "commit"
End Function

Public Function AbrirConexionUsuarios() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionUsuarios = False
    Set conn = Nothing
    Set conn = New Connection
    'Conn.CursorLocation = adUseClient
    conn.CursorLocation = adUseServer
    
'    cad = "DSN=vUsuarios;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=usuarios;"
'    cad = cad & "SERVER=" & vConfig.SERVER & ";UID=" & vConfig.User & ";PASSWORD=" & vConfig.password & ";PORT=3306;OPTION=3;STMT=;"
    
    '[Monica]29/06/2018: metemos referencia al odbc DATA SOURCE=vUsuarios
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE=vUsuarios;DATABASE=usuarios;SERVER=" & vConfig.SERVER

    Cad = Cad & ";UID=" & vConfig.User
    Cad = Cad & ";PWD=" & vConfig.password

    '++monica: tema del vista
    Cad = Cad & ";Persist Security Info=true"
    '++
    
    conn.ConnectionString = Cad
    conn.Open
    AbrirConexionUsuarios = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n usuarios.", Err.Description
End Function


Public Sub AbrirConexionMonasterios()
Dim Cad As String
Dim SQL As String
Dim Rs As ADODB.Recordset

    On Error Resume Next
    
    Set conn = Nothing
    Set conn = New Connection
    'Conn.CursorLocation = adUseClient
    conn.CursorLocation = adUseServer
    
'    cad = "DSN=vUsuarios;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=usuarios;"
'    cad = cad & "SERVER=" & vConfig.SERVER & ";UID=" & vConfig.User & ";PASSWORD=" & vConfig.password & ";PORT=3306;OPTION=3;STMT=;"
    
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;"
    '[Monica]29/06/2018: ponemos de nuevo el odbc
    Cad = Cad & "DATA SOURCE=vAriagro;DATABASE=ariagro1;SERVER=localhost" '& vConfig.SERVER
    Cad = Cad & ";UID=root" ' & vConfig.User
    Cad = Cad & ";PWD=aritel" '& vConfig.password

    '++monica: tema del vista
    Cad = Cad & ";Persist Security Info=true"
    '++
    
    conn.ConnectionString = Cad
    conn.Open
    
    EsMonasterios = False
    
    SQL = "select cooperativa from rparam "
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        EsMonasterios = (DBLet(Rs.Fields(0), "N") = 17)
    End If
    
    conn.Close
    Set conn = Nothing
End Sub





Public Sub CommitConexion()
On Error Resume Next
    conn.Execute "Commit"
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function LeerNivelesEmpresa()
 'Crea instancia de la clase Cempresa con los valores en
 'Tabla: Empresa
 'BDatos: Conta
        
        If vEmpresa.LeerNiveles = 1 Then
            MsgBox "No se han podido cargar los niveles de la contabilidad de la empresa. Debe configurar la aplicaci�n.", vbExclamation
            Set vEmpresa = Nothing
        End If
            
End Function

'--------------------------------------------------------------------
'-------------------------------------------------------------------
'Para el envio de los mails
Public Function PrepararCarpetasEnvioMail(Optional NoBorrar As Boolean) As Boolean
    On Error GoTo EPrepararCarpetasEnvioMail
    PrepararCarpetasEnvioMail = False

    If Dir(App.Path & "\temp", vbDirectory) = "" Then
        MkDir App.Path & "\temp"
    Else
        If Not NoBorrar Then
            If Dir(App.Path & "\temp\*.*", vbArchive) <> "" Then Kill App.Path & "\temp\*.*"
        End If
    End If


    PrepararCarpetasEnvioMail = True
    Exit Function
EPrepararCarpetasEnvioMail:
    MuestraError Err.Number, "", "Preparar Carpetas Envio Mail "
End Function


'------------------------------------------------------------------
'   Comprobara si una daterminada fecha esta o no en los ejercicios
'   contables (actual y siguiente)
'   Dando un O: SI. Correcto. Ok
'            1: Inferior
'            2: Superior

Public Function EsFechaOKConta(Fecha As Date) As Byte
Dim F2 As Date

    If vEmpresa.FechaIni > Fecha Then
        EsFechaOKConta = 1
    Else
        F2 = DateAdd("yyyy", 1, vEmpresa.FechaFin)
        If Fecha > F2 Then
            EsFechaOKConta = 2
        Else
            'OK. Dentro de los ejercicios contables
            EsFechaOKConta = 0
        End If
    End If
    '[Monica]20/06/2017: de david
    If EsFechaOKConta = 0 Then
        'Si tiene SII
        If vParamAplic.ContabilidadNueva Then
            If vEmpresa.TieneSII Then
                '[Monica]06/10/2017: a�adida la segunda condicion: fecha > vEmpresa.SIIFechaInicio
                '                    fallaba cuando la fecha es anterior a la declaracion del SII.
                '                    Caso de Coopic con una factura interna
                'If DateDiff("d", Fecha, Now) > vEmpresa.SIIDiasAviso And Fecha > vEmpresa.SIIFechaInicio Then
                '[Monica]19/02/2018: fines de semana
                If Fecha < UltimaFechaCorrectaSII(vEmpresa.SIIDiasAviso, Now) Then
                    MensajeFechaOkConta = "Fecha fuera de periodo de comunicaci�n SII."
                    'LLEVA SII y han trascurrido los dias
                    If vUsu.Nivel = 0 Then
                        If MsgBox(MensajeFechaOkConta & vbCrLf & "�Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then
                            EsFechaOKConta = 4
                        End If
                    Else
                        'NO tienen nivel
                        EsFechaOKConta = 5
                    End If
                End If
            End If
        End If
    Else
        MensajeFechaOkConta = "Fuera de ejercicios contables"
    End If

End Function


Public Function ejecutar(ByRef SQL As String, OcultarMsg As Boolean) As Boolean
    On Error Resume Next
    conn.Execute SQL
    If Err.Number <> 0 Then
        If Not OcultarMsg Then MuestraError Err.Number, Err.Description, SQL
        ejecutar = False
    Else
        ejecutar = True
    End If
End Function

