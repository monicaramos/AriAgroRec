VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GRPTC_Vehiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarnomcamio As String 'copia local
Private mvarmatricul As String 'copia local

Public Function GenXML() As String
    ' Genera los datos XML tal y como los quiere GRUPOTEC
    Dim mXML As String
    mXML = "<?xml version=""1.0"" standalone=""yes""?>"
    mXML = mXML & "<DATAPACKET Version=""2.0"">"
    mXML = mXML & "<METADATA>"
    mXML = mXML & "<FIELDS>"
    mXML = mXML & "<FIELD attrname=""matricul"" fieldtype=""string"" WIDTH=""10""/>"
    mXML = mXML & "<FIELD attrname=""nomcamio"" fieldtype=""string"" WIDTH=""30""/>"
    mXML = mXML & "</FIELDS>"
    mXML = mXML & "</METADATA>"
    mXML = mXML & "<ROWDATA>"
        mXML = mXML & "<ROW "
            mXML = mXML & "matricul=""" & mvarmatricul & """ "
            mXML = mXML & "nomcamio=""" & mvarnomcamio & """ "
        mXML = mXML & " />"
    mXML = mXML & "</ROWDATA>"
    mXML = mXML & "</DATAPACKET>"
    GenXML = mXML
End Function


Public Property Let matricul(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.matricul = 5
    mvarmatricul = vData
End Property


Public Property Get matricul() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.matricul
    matricul = mvarmatricul
End Property



Public Property Let nomcamio(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.nomcamio = 5
    mvarnomcamio = vData
End Property


Public Property Get nomcamio() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.nomcamio
    nomcamio = mvarnomcamio
End Property



