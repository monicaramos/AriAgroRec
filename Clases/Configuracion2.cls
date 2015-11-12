VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Configuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'#Generador DAVID(Ariadna)



'-------------------------------------------------------
'-------------------------------------------------------
'Declaracion de las variables
'---- Servidor donde esta MYSQL
Private mvarSERVER As String 'Variable local
'---- Numero de contabilidades maximo

Private mvarUser As String 'Variable local
'---- General
Private mvarpassword As String 'Variable local
'--------
Private mvarPedirPasswd As Boolean

'---- Carpeta Integraciones
Private mvarIntegraciones As String

'---- Fichero de estructura de la base de datos para crear nueva campaña
Private mvarEstructura As String


'---- Fichero de estructura del explorador
Private mvarExplorador As String


'Variables de este modulo
Dim NF As Integer
Dim Mens As String



'------------------------------------------------------
'------------------------------------------------------
'      Propiedades del modulo de configuracion
'------------------------------------------------------
'------------------------------------------------------

'----
'---- Servidor donde esta MYSQL
'----
Public Property Let SERVER(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.SERVER = 5
     mvarSERVER = vData
End Property

Public Property Get SERVER() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.SERVER
     SERVER = mvarSERVER
End Property








'----
'---- Usuario general BD
'----
Public Property Let User(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.User = 5
     mvarUser = vData
End Property

Public Property Get User() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.User
     User = mvarUser
End Property




'----
'---- General
'----
Public Property Let password(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.password = 5
     mvarpassword = vData
End Property

Public Property Get password() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     password = mvarpassword
End Property


'---------------------------
' Pedir password


Public Property Let PedirPasswd(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.SERVER = 5
     mvarPedirPasswd = vData
End Property

Public Property Get PedirPasswd() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.SERVER
     PedirPasswd = mvarPedirPasswd
End Property

'----
'---- Integraciones
'----
Public Property Let Integraciones(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.SERVER = 5
     mvarIntegraciones = vData
End Property

Public Property Get Integraciones() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.SERVER
     Integraciones = mvarIntegraciones
End Property

'----
'----
'----
Public Property Let FichGene(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.SERVER = 5
     mvarEstructura = vData
End Property

Public Property Get FichGene() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.SERVER
     FichGene = mvarEstructura
End Property


'----
'----
'----
Public Property Let Explorador(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.SERVER = 5
     mvarExplorador = vData
End Property

Public Property Get Explorador() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.SERVER
     Explorador = mvarExplorador
End Property



'//////////////////////////////////////////////////////
'             LEER  Y  GUARDAR
'//////////////////////////////////////////////////////
'
'
'
'
'
'
'------------------------------------------------------
'     LEER
'
Public Function Leer() As Byte
    '--------------------------------------------------
    '   Leer:
    '   Busca un fichero de nombre Config.ini y lee de él los valores
    '   de parametrización para la aplicación. El fichero se busca
    '   en el directorio desde el que se ejecuta la aplicación.
    '--------------------------------------------------
    Dim fichero As String ' Nombre completo del fichero a leer
    Dim LinAnt As String
    Dim Lin As String
    Dim NomParam As String ' Para leer el nombre del parámetro
    Dim ValParam As String ' Para leer el valor del parámetro
    Dim NFile As Integer
    
    On Error GoTo Err1
    fichero = App.Path & "\ConfigAgroRec.ini"
    NFile = FreeFile
    Open fichero For Input As #NFile   ' Abre el archivo.
    Do While Not EOF(NFile) ' Repite el bucle hasta el final del archivo.
        Line Input #NFile, Lin
        Select Case Lin
           Case "[SERVER]", "[User]", "[password]", "[PEDIR]", "[INTEG]", "[FichGene]", "[FichExplo]"
               
               
               ValParam = ""
               NomParam = Lin

           Case "<-------->"
               Select Case NomParam
                    Case "[SERVER]": mvarSERVER = ValParam
                    Case "[User]": mvarUser = ValParam
                    Case "[password]": mvarpassword = ValParam
                    Case "[PEDIR]":
                    Case "[INTEG]": mvarIntegraciones = ValParam
                    Case "[FichGene]": mvarEstructura = ValParam
                    Case "[FichExplo]": mvarExplorador = ValParam
                    
                End Select
            Case Else
                If ValParam = "" Then
                    ValParam = Lin
                Else
                    ValParam = ValParam & vbCrLf & Lin
                End If
        End Select
    Loop
    Close #NFile    ' Cierra el archivo.
    
    'Desencriptamos password
    Desencripta (True)
    Desencripta (False)
    Leer = 0
    Exit Function
Err1:
   Leer = 1
End Function




'------------------------------------------------------
'     GUARDAR

Public Function Grabar() As Byte
    '--------------------------------------------------
    '   GuardaConfig:
    '   Guardamos en el fichero Config.ini los valores de los
    '   parámetros de la aplicación
    '--------------------------------------------------
    Dim fichero As String ' Nombre completo del fichero a leer
    Dim NFile As Integer

    On Error GoTo Err1
    fichero = App.Path & "\ConfigAgroRec.ini"
    NFile = FreeFile
    Open fichero For Output As #NFile   ' Abre el archivo.
    Print #NFile, "[SERVER]"
    Print #NFile, mvarSERVER
    Print #NFile, "<-------->"
    Print #NFile, "[User]"
    Print #NFile, Encripta(mvarUser)
    Print #NFile, "<-------->"
    Print #NFile, "[password]"
    Print #NFile, Encripta(mvarpassword)
    Print #NFile, "<-------->"
    Print #NFile, "[PEDIR]"
    Print #NFile, mvarPedirPasswd
    Print #NFile, "<-------->"
    Print #NFile, "[INTEG]"
    Print #NFile, mvarIntegraciones
    Print #NFile, "<-------->"
    Print #NFile, "[FichGene]"
    Print #NFile, mvarEstructura
    Print #NFile, "<-------->"
    Print #NFile, "[FichExplo]"
    Print #NFile, mvarExplorador
    Print #NFile, "<-------->"

    
    
    Close #NFile
    Grabar = 0
    Exit Function
Err1:
    Mens = "Error " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbCritical
    Grabar = 1
End Function


Private Sub Desencripta(ElPassword As Boolean)
Dim C As String
Dim i As Integer

If ElPassword Then
    C = mvarpassword
    Else
    C = mvarUser
End If
Mens = ""
For i = 1 To Len(C)
    If (i Mod 2) <> 0 Then
        'Solo valen los caracteres impares
        NF = Asc(Mid(C, i, 1))
        NF = NF + 1
        Mens = Mens & Chr(NF)
    End If
Next i

If ElPassword Then
    mvarpassword = Mens
    Else
    mvarUser = Mens
End If

End Sub



Private Function Encripta(ByRef C As String) As String
Dim i As Integer


Mens = ""
Randomize
For i = 1 To Len(C)
        'Solo valen los caracteres impares
        NF = Asc(Mid(C, i, 1))
        NF = NF - 1
        Mens = Mens & Chr(NF)
        'Y el aleatorio , entre 48 y 122  <- 74 ->
        NF = Int((74 * Rnd)) ' Genera valores aleatorios entre 0y74.
        NF = NF + 48
        Mens = Mens & Chr(NF)
Next i

Encripta = Mens

End Function
