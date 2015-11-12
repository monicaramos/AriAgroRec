VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GRPTC_Partida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarcodparti As Long 'copia local
Private mvarnomparti As String 'copia local

Public Function GenXML() As String
    ' Genera los datos XML tal y como los quiere GRUPOTEC
    Dim mXML As String
    mXML = "<?xml version=""1.0"" standalone=""yes""?>"
    mXML = mXML & "<DATAPACKET Version=""2.0"">"
    mXML = mXML & "<METADATA>"
    mXML = mXML & "<FIELDS>"
    mXML = mXML & "<FIELD attrname=""codparti"" fieldtype=""i4"" />"
    mXML = mXML & "<FIELD attrname=""nomparti"" fieldtype=""string"" WIDTH=""30""/>"
    mXML = mXML & "</FIELDS>"
    mXML = mXML & "</METADATA>"
    mXML = mXML & "<ROWDATA>"
        mXML = mXML & "<ROW "
            mXML = mXML & "codparti=""" & CStr(mvarcodparti) & """ "
            mXML = mXML & "nomparti=""" & mvarnomparti & """ "
        mXML = mXML & " />"
    mXML = mXML & "</ROWDATA>"
    mXML = mXML & "</DATAPACKET>"
    GenXML = mXML
End Function


Public Property Let nomparti(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.nomparti = 5
    mvarnomparti = vData
End Property


Public Property Get nomparti() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.nomparti
    nomparti = mvarnomparti
End Property



Public Property Let codparti(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codparti = 5
    mvarcodparti = vData
End Property


Public Property Get codparti() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codparti
    codparti = mvarcodparti
End Property



