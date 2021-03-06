Attribute VB_Name = "busAriLecturas"


'NOTA: en este m�dul, adem�s, n'hi han funcions generals que no siguen de formularis (molt b�)
Option Explicit

Public Const ValorNulo = "Null"
Public NombreCheck As String

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



'[Monica] 06/09/2010: sustituida esta constante por el parametro vParamAplic.Faneca
Public Const cFaneca As Single = 0.0833 ' hectareas





'Que conexion a base de datos se va a utilizar
Public Const cAgro As Byte = 1 'trabajaremos con conn (conexion a BD Ariagro)
Public Const cConta As Byte = 2 'trabajaremos con connConta (cxion a BD Contabilidad)
Public Const cAridoc As Byte = 3 'trabajaremos con connAridoc (cxion a BD Aridoc)
Public Const cAriges As Byte = 4 'trabajaremos con connAriges (conexion a BD Suministros)

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


Dim frmLect As frmPOZLecturasMonast
Public EsMonasterios As Boolean
Public Semaforo As Byte

'Inicio Aplicaci�n
Public Sub Main()
Dim OK As Byte

'     If App.PrevInstance Then
'        MsgBox "Ariagro ya se esta ejecutando", vbExclamation
'        End
'     End If
     
'cambiarlo por lo de abajo, cuando quiera quitar el login

'**********
        Set vConfig = New Configuracion
        If vConfig.Leer = 1 Then

             MsgBox "MAL CONFIGURADO", vbCritical
             End
             Exit Sub
        End If
        
        
        
         'Abrimos conexion para comprobar el usuario
         'Luego, en funcion del nivel de usuario que tenga cerraremos la conexion
         'y la abriremos con usuario-codigo ajustado a su nivel
         If AbrirConexionUsuarios() = False Then
             MsgBox "La aplicaci�n no puede continuar sin acceso a los datos. ", vbCritical
             End
         End If

        Set vUsu = New Usuario
        
        If vUsu.Leer("Lectura") = 0 Then
'[Monica]12/12/2017: no debemos mirar la contrase�a
            'Con exito
'            If vUsu.PasswdPROPIO = "l" Then
'                OK = 0
'            Else
'                OK = 1
'            End If
            OK = 0
        Else
            OK = 2
        End If
        
        If OK <> 0 Then
            MsgBox "Usuario-Clave Incorrecto de lecturas", vbExclamation
            End
        Else
            vUsu.CadenaConexion = "ariagro1"
        End If


        Set vEmpresa = New Cempresa
        If vEmpresa.LeerDatos = 1 Then
            MsgBox "No se han podido cargar datos empresa (BD:usuarios). Debe configurar la aplicaci�n.", vbExclamation
            Set vEmpresa = Nothing
        End If



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
                
                

        'Gestionar el nombre del PC para la asignacion de PC en el entorno de red
        GestionaPC

        'Otras acciones
        OtrasAcciones

        ' *** per als iconos XP ***
        GetIconsFromLibrary App.Path & "\iconos.dll", 1, 24
        GetIconsFromLibrary App.Path & "\iconos_BN.dll", 2, 24
        GetIconsFromLibrary App.Path & "\iconos_OM.dll", 3, 24
        
        GetIconsFromLibrary App.Path & "\iconosAriagroRec.dll", 4, 24
         
     
        '[Monica]08/09/2017: para el caso de que se ejecute el programa de lecturas en la tableta
        If vParamAplic.Cooperativa = 17 Then
            Set frmLect = New frmPOZLecturasMonast
            frmLect.Show 'vbModal
            'Set frmLect = Nothing
        End If


'**********


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
    
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & BBDD & ";SERVER=" & vParamAplic.ServidorConta & ";"
    
    Cad = Cad & ";UID=" & vParamAplic.UsuarioConta
    Cad = Cad & ";PWD=" & vParamAplic.PasswordConta
    '---- Laura: 29/09/2006
    Cad = Cad & ";PORT=3306;OPTION=3;STMT=;"
    '----
    '++monica: tema de vista
    Cad = Cad & "Persist Security Info=true"
    
    
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
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & BBDD & ";SERVER=" & vParamAplic.ServidorConta & ";"
    
    Cad = Cad & ";UID=" & vParamAplic.UsuarioConta
    Cad = Cad & ";PWD=" & vParamAplic.PasswordConta
    '---- Laura: 29/09/2006
    Cad = Cad & ";PORT=3306;OPTION=3;STMT=;"
    '----
    '++monica: tema de vista
    Cad = Cad & "Persist Security Info=true"
    
    
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

    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=aridoc;SERVER=" & vConfig.SERVER & ";"
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
    
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & vParamAplic.BDAriges & ";SERVER=" & vParamAplic.ServidorConta & ";"
    Cad = Cad & ";UID=" & vParamAplic.UsuarioConta
    Cad = Cad & ";PWD=" & vParamAplic.PasswordConta
    '---- Laura: 29/09/2006
    Cad = Cad & ";PORT=3306;OPTION=3;STMT=;"
    '----
    '++monica: tema de vista
    Cad = Cad & "Persist Security Info=true"
    
    ConnAriges.ConnectionString = Cad
    ConnAriges.Open
    ConnAriges.Execute "Set AUTOCOMMIT = 1"
    AbrirConexionAriges = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n Ariges.", Err.Description
End Function




Public Function AbrirConexionAuxCon(Empresa As String, Usuario As String, Pass As String) As Boolean
Dim Cad As String
Dim nomConta As String 'nombre de la BD de la contabilidad
Dim serConta As String 'servidor donde esta la BD de la contabilidad
On Error GoTo EAbrirConexion

    AbrirConexionAuxCon = False

    Set ConnAuxCon = Nothing
    Set ConnAuxCon = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnAuxCon.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

    'Obtener la BD de contabilidad
'    SQL = "select bdaconta FROM paramcon WHERE codempre=" & codEmpre
    serConta = "serconta"
    nomConta = DevuelveDesdeBDNew(2, "sparam", "bdaconta", "codempre", Empresa, "N", serConta)
'    vEmpresa.BDConta = nomConta
    If nomConta <> "" Then
    '    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamConta.NumeroConta & ";SERVER=" & vParamConta.ServidorConta & ";"
    '    cad = cad & ";UID=" & vParamConta.UsuarioConta
    '    cad = cad & ";PWD=" & vParamConta.PasswordConta
    '    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta2;SERVER=david;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
        If serConta <> "" Then 'especificamos servidor
            Cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";SERVER=" & serConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
        Else 'por defecto cogera la BD del servidor que haya en el ODBC
            Cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
        End If
        ConnAuxCon.ConnectionString = Cad
        ConnAuxCon.Open
        ConnAuxCon.Execute "Set AUTOCOMMIT = 1"
        AbrirConexionAuxCon = True
    Else
        AbrirConexionAuxCon = False
    End If
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n Contabilidad.", Err.Description
End Function

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

Public Function DBSet(vData As Variant, tipo As String, Optional EsNulo As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
Dim Cad As String

        If IsNull(vData) Then
            DBSet = ValorNulo
            Exit Function
        End If

        If tipo <> "" Then
            Select Case tipo
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



Public Function DBLet(vData As Variant, Optional tipo As String) As Variant
'Para cuando recupera Datos de la BD
    If IsNull(vData) Then
        DBLet = ""
        If tipo <> "" Then
            Select Case tipo
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
Dim I As Integer

    If Importe = "" Then
        ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            I = InStr(1, Importe, ".")
            If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
        Loop Until I = 0
        ImporteFormateado = Importe
    End If
End Function

' ### [Monica] 11/09/2006
Public Function ImporteSinFormato(cadena As String) As String
Dim I As Integer
'Quitamos puntos
Do
    I = InStr(1, cadena, ".")
    If I > 0 Then cadena = Mid(cadena, 1, I - 1) & Mid(cadena, I + 1)
Loop Until I = 0
ImporteSinFormato = TransformaPuntosComas(cadena)
End Function



'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(cadena As String) As String
Dim I As Integer
    Do
        I = InStr(1, cadena, ",")
        If I > 0 Then
            cadena = Mid(cadena, 1, I - 1) & "." & Mid(cadena, I + 1)
        End If
    Loop Until I = 0
    TransformaComasPuntos = cadena
End Function

'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef cadena As String)
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, cadena, "'")
        If I > 0 Then
            Aux = Mid(cadena, 1, I - 1) & "\"
            cadena = Aux & Mid(cadena, I)
            J = I + 2
        End If
    Loop Until I = 0
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
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, cadena, "'")
        If I > 0 Then
            Aux = Mid(cadena, 1, I - 1) & "\"
            cadena = Aux & Mid(cadena, I)
            J = I + 2
        End If
    Loop Until I = 0
    DevNombreSQL = cadena
End Function


Public Function DevuelveDesdeBD(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional tipo As String, Optional ByRef OtroCampo As String) As String
    Dim Rs As Recordset
    Dim Cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    Cad = "Select " & kCampo
    If OtroCampo <> "" Then Cad = Cad & ", " & OtroCampo
    Cad = Cad & " FROM " & Ktabla
    Cad = Cad & " WHERE " & Kcodigo & " = "
    If tipo = "" Then tipo = "N"
    Select Case tipo
    Case "N"
        'No hacemos nada
        Cad = Cad & ValorCodigo
    Case "T", "F"
        Cad = Cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & tipo & " no definido", vbExclamation
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
        Case "T"
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
        Case "T"
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
        Case "T"
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
Public Function DevuelveDesdeBDnew2(kBD As Integer, kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional tipo As String, Optional num As Byte, Optional ByRef OtroCampo As String) As String
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
    tip = RecuperaValor(tipo, v_aux)
        
    Cad = Cad & campo & "="
    If tip = "" Then tipo = "N"
    
    Select Case tip
            Case "N"
                'No hacemos nada
                Cad = Cad & Valor
            Case "T", "F"
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
Dim I As Integer
Dim c As Integer
Dim L As Integer
Dim res As Boolean

    res = True
    EsEntero = False

    If Not IsNumeric(Texto) Then
        res = False
    Else
        'Vemos si ha puesto mas de un punto
        c = 0
        L = 1
        Do
            I = InStr(L, Texto, ".")
            If I > 0 Then
                L = I + 1
                c = c + 1
            End If
        Loop Until I = 0
        If c > 1 Then res = False
        
        'Si ha puesto mas de una coma y no tiene puntos
        If c = 0 Then
            L = 1
            Do
                I = InStr(L, Texto, ",")
                If I > 0 Then
                    L = I + 1
                    c = c + 1
                End If
            Loop Until I = 0
            If c > 1 Then res = False
        End If
        
    End If
        EsEntero = res
End Function

Public Function TransformaPuntosComas(cadena As String) As String
    Dim I As Integer
    Do
        I = InStr(1, cadena, ".")
        If I > 0 Then
            cadena = Mid(cadena, 1, I - 1) & "," & Mid(cadena, I + 1)
        End If
        Loop Until I = 0
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
Dim I As Integer
Dim Cad As String
Dim metag As String
Dim SQL As String
Cad = OtrosPCsContraAplicacion
UsuariosConectados = False
If Cad <> "" Then
    UsuariosConectados = True
    I = 1
    metag = "Los siguientes PC's est�n conectados a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
    Do
        SQL = RecuperaValor(Cad, I)
        If SQL <> "" Then
            metag = metag & "    - " & SQL & vbCrLf
            I = I + 1
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
                        
                       
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & vUsu.CadenaConexion & ";SERVER=" & vConfig.SERVER & ";"
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
    
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & BaseDatos & ";SERVER=" & vConfig.SERVER & ";"
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
    
    If Servidor = 0 Then
        Cad = "Driver={SQL Server};Server=" & Trim(vParamAplic.SqlServer) & ",1433;Database=ProductionV50;" & _
                "Uid=client;Pwd=client;"
    Else
        Cad = "Driver={SQL Server};Server=" & Trim(vParamAplic.SqlServer1) & ",1433;Database=ProductionV50;" & _
                "Uid=client;Pwd=client;"
    End If

    CnnSqlServer.ConnectionString = Cad
    CnnSqlServer.Open
    AbrirConexionSqlSERVER = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n SqlSERVER.", Err.Description
End Function


Public Function AbrirConexionSqlMonasterios(Servidor As Byte) As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion
    
    AbrirConexionSqlMonasterios = False
    Set CnnSqlMonast = Nothing
    Set CnnSqlMonast = New Connection
    

'    cad = "Provider=MSDASQL.1;Persist Security Info=false;User ID=escalona;Data Source=vMonasterios"
'    cad = cad & ";Persist Security Info=true"
'
'
'    cad = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BBDD_Monasterios;"
'    cad = cad & "Extended Properties=dBASE 5.0;User ID=Admin;Password=;"

'   cad = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=c:\BBDD_Monasterios\;"

   Cad = "DSN=vMonast1;"
   Cad = Cad & ";Persist Security Info=true"
    
    
'11/09/2017
'    Dim strDB As String
'    strDB = "utxera.mdb"
'
'    cad = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDB & ";Persist Security Info=False"
' funciona pero con mdb
'******
'    cad = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\programas\ariagrorec;Extended Properties=dBASE IV;User ID=Admin;Password="
'   no funciona: external table is not in the expected format
'******

'    cad = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=C:\programas\ariagrorec;"
'   no funciona: external table is not in the expected format
'****

    CnnSqlMonast.ConnectionString = Cad
    CnnSqlMonast.Open
    AbrirConexionSqlMonasterios = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n Sql LibreOffice.", Err.Description
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

Public Function ComprobarEmpresaBloqueada(Codusu As Long, ByRef Empresa As String) As Boolean
Dim Cad As String

ComprobarEmpresaBloqueada = False

'Antes de nada, borramos las entradas de usuario, por si hubiera kedado algo
conn.Execute "Delete from usuarios.vbloqbd where codusu=" & Codusu

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
    
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=usuarios;SERVER=" & vConfig.SERVER

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
    
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=ariagro1;SERVER=localhost" '& vConfig.SERVER
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
                If DateDiff("d", Fecha, Now) > vEmpresa.SIIDiasAviso And Fecha > vEmpresa.SIIFechaInicio Then
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


Public Function ComprobarCero(Valor As String) As String
    If Valor = "" Then
        ComprobarCero = "0"
    Else
        ComprobarCero = Valor
    End If
End Function


Public Function DevuelveValor(vSQL As String) As Variant
'Devuelve el valor de la SQL
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    DevuelveValor = 0
    If Not Rs.EOF Then
        ' antes RS.Fields(0).Value > 0
        If Not IsNull(Rs.Fields(0).Value) Then DevuelveValor = Rs.Fields(0).Value   'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        DevuelveValor = 0
        Err.Clear
    End If
End Function

Public Function RecuperaValor(ByRef cadena As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim cont As Integer
Dim Cad As String

    I = 0
    cont = 1
    Cad = ""
    Do
        J = I + 1
        I = InStr(J, cadena, "|")
        If I > 0 Then
            If cont = Orden Then
                Cad = Mid(cadena, J, I - J)
                I = Len(cadena) 'Para salir del bucle
                Else
                    cont = cont + 1
            End If
        End If
    Loop Until I = 0
    RecuperaValor = Cad
End Function

Public Sub PonerOpcionesMenuGeneral(ByRef formulario As Form)
Dim I As Integer
Dim J As Integer
'Dim bol As Boolean

On Error GoTo EPonerOpcionesMenuGeneral
'bol = vSesion.Nivel < 2

'A�adir, modificar y borrar deshabilitados si no nivel
With formulario
    For I = 1 To .Toolbar1.Buttons.Count
        If .Toolbar1.Buttons(I).Tag <> "" Then
            J = Val(.Toolbar1.Buttons(I).Tag)
            If J < vUsu.Nivel Then
                .Toolbar1.Buttons(I).Enabled = False
            End If
        End If
    Next I
End With

Exit Sub
EPonerOpcionesMenuGeneral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub


Public Sub PonerFrameVisible(ByRef vFrame As Frame, visible As Boolean, H As Integer, W As Integer)
'Pone el Frame Visible y Ajustado al Formulario, y visualiza los controles
    
        vFrame.visible = visible
        If visible = True Then
            'Ajustar Tama�o del Frame para ajustar tama�o de Formulario al del Frame
            vFrame.Top = -90
            vFrame.Left = 0
            vFrame.Width = W
            vFrame.Height = H
        End If
End Sub


Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String, ByVal op As Integer, ByVal tam As Integer)
    Dim I As Integer
    Dim tRes As ResType, iCount As Integer
        
    opcio = op
    tamany = tam
    ghmodule = LoadLibraryEx(sLibraryFilePath, 0, DONT_RESOLVE_DLL_REFERENCES)

    If ghmodule = 0 Then
        MsgBox "Invalid library file.", vbCritical
        Exit Sub
    End If
        
    For tRes = RT_FIRST To RT_LAST
        DoEvents
        EnumResourceNames ghmodule, tRes, AddressOf EnumResNameProc, 0
    Next
    FreeLibrary ghmodule
             
End Sub
