VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GRPTC_Campo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarcodsocio As Long 'copia local
Private mvarcodcampo As Long 'copia local
Private mvarcodprodu As Long 'copia local
Private mvarcodvarie As Long 'copia local
Private mvarcodparti As Long 'copia local
Private mvarhanegada As Single 'copia local
Private mvarpoligono As Integer 'copia local

Private mvarTipoProdu As Long 'copia local

'variables locales para almacenar los valores de las propiedades
Private mvarnumarbol As Integer 'copia local

Public Property Let numarbol(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.numarbol = 5
    mvarnumarbol = vData
End Property


Public Property Get numarbol() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.numarbol
    numarbol = mvarnumarbol
End Property




Public Function GenXML() As String
    ' Genera los datos XML tal y como los quiere GRUPOTEC
    Dim mXML As String
    mXML = "<?xml version=""1.0"" standalone=""yes""?>"
    mXML = mXML & "<DATAPACKET Version=""2.0"">"
    mXML = mXML & "<METADATA>"
    mXML = mXML & "<FIELDS>"
    mXML = mXML & "<FIELD attrname=""codsocio"" fieldtype=""i4"" />"
    '[Monica]07/01/2010 antes i2 en el campo
    mXML = mXML & "<FIELD attrname=""codcampo"" fieldtype=""i4"" />"
    mXML = mXML & "<FIELD attrname=""codprodu"" fieldtype=""i2"" />"
    mXML = mXML & "<FIELD attrname=""codvarie"" fieldtype=""i2"" />"
    mXML = mXML & "<FIELD attrname=""codparti"" fieldtype=""i2"" />"
    mXML = mXML & "<FIELD attrname=""hanegada"" fieldtype=""r8"" />"
    mXML = mXML & "<FIELD attrname=""numarbol"" fieldtype=""i2"" />"
    mXML = mXML & "<FIELD attrname=""poligono"" fieldtype=""i2"" />"
    
    If vParamAplic.Cooperativa = 16 Then
        mXML = mXML & "<FIELD attrname=""tipoproduccion"" fieldtype=""i2"" />"
    End If
    
    
    mXML = mXML & "</FIELDS>"
    mXML = mXML & "</METADATA>"
    mXML = mXML & "<ROWDATA>"
        mXML = mXML & "<ROW "
            mXML = mXML & "codsocio=""" & CStr(mvarcodsocio) & """ "
            mXML = mXML & "codcampo=""" & CStr(mvarcodcampo) & """ "
            mXML = mXML & "codprodu=""" & CStr(mvarcodprodu) & """ "
            mXML = mXML & "codvarie=""" & CStr(mvarcodvarie) & """ "
            mXML = mXML & "codparti=""" & CStr(mvarcodparti) & """ "
            mXML = mXML & "hanegada=""" & CStr(mvarhanegada) & """ "
            mXML = mXML & "numarbol=""" & CStr(mvarnumarbol) & """ "
            mXML = mXML & "poligono=""" & CStr(mvarpoligono) & """ "
            
            If vParamAplic.Cooperativa = 16 Then
                mXML = mXML & "tipoproducion=""" & CStr(mvarTipoProdu) & """ "
            End If
            
        mXML = mXML & " />"
    mXML = mXML & "</ROWDATA>"
    mXML = mXML & "</DATAPACKET>"
    GenXML = mXML
End Function

Public Property Let Poligono(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.poligono = 5
    mvarpoligono = vData
End Property


Public Property Get Poligono() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.poligono
    Poligono = mvarpoligono
End Property



Public Property Let hanegada(ByVal vData As Single)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.hanegada = 5
    mvarhanegada = vData
End Property


Public Property Get hanegada() As Single
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.hanegada
    hanegada = mvarhanegada
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



Public Property Let codcampo(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codcampo = 5
    mvarcodcampo = vData
End Property


Public Property Get codcampo() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codcampo
    codcampo = mvarcodcampo
End Property



Public Property Let Codsocio(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codsocio = 5
    mvarcodsocio = vData
End Property

Public Property Get Codsocio() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codsocio
    Codsocio = mvarcodsocio
End Property

Public Property Let TipoProdu(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codsocio = 5
    mvarTipoProdu = vData
End Property

Public Property Get TipoProdu() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codsocio
    TipoProdu = mvarTipoProdu
End Property

