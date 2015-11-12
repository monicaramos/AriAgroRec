VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GRPTC_Socio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarcodsocio As Long 'copia local
Private mvarnifsocio As String 'copia local
Private mvarnomsocio As String 'copia local
Private mvardomsocio As String 'copia local
Private mvartelsocio As String 'copia local
Private mvarcodpobla As Long 'copia local

Public Function GenXML() As String
    ' Genera los datos XML tal y como los quiere GRUPOTEC
    Dim mXML As String
    mXML = "<?xml version=""1.0"" standalone=""yes""?>"
    mXML = mXML & "<DATAPACKET Version=""2.0"">"
    mXML = mXML & "<METADATA>"
    mXML = mXML & "<FIELDS>"
    mXML = mXML & "<FIELD attrname=""codsocio"" fieldtype=""i4"" />"
    mXML = mXML & "<FIELD attrname=""nifsocio"" fieldtype=""string"" WIDTH=""9""/>"
    mXML = mXML & "<FIELD attrname=""nomsocio"" fieldtype=""string"" WIDTH=""35""/>"
    mXML = mXML & "<FIELD attrname=""domsocio"" fieldtype=""string"" WIDTH=""35""/>"
    mXML = mXML & "<FIELD attrname=""telsocio"" fieldtype=""string"" WIDTH=""10""/>"
    mXML = mXML & "<FIELD attrname=""codpobla"" fieldtype=""i4"" />"
    mXML = mXML & "</FIELDS>"
    mXML = mXML & "</METADATA>"
    mXML = mXML & "<ROWDATA>"
        mXML = mXML & "<ROW "
            mXML = mXML & "codsocio=""" & CStr(mvarcodsocio) & """ "
            mXML = mXML & "nifsocio=""" & mvarnifsocio & """ "
            mXML = mXML & "nomsocio=""" & mvarnomsocio & """ "
            mXML = mXML & "domsocio=""" & mvardomsocio & """ "
            mXML = mXML & "telsocio=""" & mvartelsocio & """ "
            mXML = mXML & "codpobla=""" & CStr(mvarcodpobla) & """ "
        mXML = mXML & " />"
    mXML = mXML & "</ROWDATA>"
    mXML = mXML & "</DATAPACKET>"
    GenXML = mXML
End Function


Public Property Let codpobla(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codpobla = 5
    mvarcodpobla = vData
End Property


Public Property Get codpobla() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codpobla
    codpobla = mvarcodpobla
End Property



Public Property Let telsocio(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.telsocio = 5
    mvartelsocio = vData
End Property


Public Property Get telsocio() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.telsocio
    telsocio = mvartelsocio
End Property



Public Property Let domsocio(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.domsocio = 5
    mvardomsocio = vData
End Property


Public Property Get domsocio() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.domsocio
    domsocio = mvardomsocio
End Property



Public Property Let nomsocio(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.nomsocio = 5
    mvarnomsocio = vData
End Property


Public Property Get nomsocio() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.nomsocio
    nomsocio = mvarnomsocio
End Property



Public Property Let nifsocio(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.nifsocio = 5
    mvarnifsocio = vData
End Property


Public Property Get nifsocio() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.nifsocio
    nifsocio = mvarnifsocio
End Property



Public Property Let codsocio(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codsocio = 5
    mvarcodsocio = vData
End Property


Public Property Get codsocio() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codsocio
    codsocio = mvarcodsocio
End Property



