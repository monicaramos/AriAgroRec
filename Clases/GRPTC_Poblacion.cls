VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GRPTC_Poblacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarcodpobla As String 'copia local
Private mvardespobla As String 'copia local

Public Function GenXML() As String
    ' Genera los datos XML tal y como los quiere GRUPOTEC
    Dim mXML As String
    mXML = "<?xml version=""1.0"" standalone=""yes""?>"
    mXML = mXML & "<DATAPACKET Version=""2.0"">"
    mXML = mXML & "<METADATA>"
    mXML = mXML & "<FIELDS>"
    mXML = mXML & "<FIELD attrname=""codpobla"" fieldtype=""string"" WIDTH=""6""/>"
    mXML = mXML & "<FIELD attrname=""despobla"" fieldtype=""string"" WIDTH=""30""/>"
    mXML = mXML & "</FIELDS>"
    mXML = mXML & "</METADATA>"
    mXML = mXML & "<ROWDATA>"
        mXML = mXML & "<ROW "
            mXML = mXML & "codpobla=""" & mvarcodpobla & """ "
            mXML = mXML & "despobla=""" & mvardespobla & """ "
        mXML = mXML & " />"
    mXML = mXML & "</ROWDATA>"
    mXML = mXML & "</DATAPACKET>"
    GenXML = mXML
End Function



Public Property Let despobla(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.despobla = 5
    mvardespobla = vData
End Property


Public Property Get despobla() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.despobla
    despobla = mvardespobla
End Property



Public Property Let codpobla(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codpobla = 5
    mvarcodpobla = vData
End Property


Public Property Get codpobla() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codpobla
    codpobla = mvarcodpobla
End Property



