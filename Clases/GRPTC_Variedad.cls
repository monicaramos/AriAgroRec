VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GRPTC_Variedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarcodvarie As Long 'copia local
Private mvarnomvarie As String 'copia local
Private mvarcodprodu As Long 'copia local

Public Function GenXML() As String
    ' Genera los datos XML tal y como los quiere GRUPOTEC
    Dim mXML As String
    mXML = "<?xml version=""1.0"" standalone=""yes""?>"
    mXML = mXML & "<DATAPACKET Version=""2.0"">"
    mXML = mXML & "<METADATA>"
    mXML = mXML & "<FIELDS>"
    mXML = mXML & "<FIELD attrname=""codvarie"" fieldtype=""i4"" />"
    mXML = mXML & "<FIELD attrname=""nomvarie"" fieldtype=""string"" WIDTH=""30""/>"
    mXML = mXML & "<FIELD attrname=""codprodu"" fieldtype=""i4"" />"
    mXML = mXML & "</FIELDS>"
    mXML = mXML & "</METADATA>"
    mXML = mXML & "<ROWDATA>"
        mXML = mXML & "<ROW "
            mXML = mXML & "codvarie=""" & CStr(mvarcodvarie) & """ "
            mXML = mXML & "nomvarie=""" & mvarnomvarie & """ "
            mXML = mXML & "codprodu=""" & CStr(mvarcodprodu) & """ "
        mXML = mXML & " />"
    mXML = mXML & "</ROWDATA>"
    mXML = mXML & "</DATAPACKET>"
    GenXML = mXML
End Function


Public Property Let codprodu(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codprodu = 5
    mvarcodprodu = vData
End Property


Public Property Get codprodu() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codprodu
    codprodu = mvarcodprodu
End Property



Public Property Let nomvarie(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.nomvarie = 5
    mvarnomvarie = vData
End Property


Public Property Get nomvarie() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.nomvarie
    nomvarie = mvarnomvarie
End Property



Public Property Let codvarie(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codvarie = 5
    mvarcodvarie = vData
End Property


Public Property Get codvarie() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codvarie
    codvarie = mvarcodvarie
End Property



