Attribute VB_Name = "ModImprimir"
'convertix una posició d'un Adodc en una Selection Formula SF
Public Function POS2SF(ByRef ado As Adodc, ByRef formu As Form, Optional opcio As Integer, Optional nom_frame As String) As String
'si opcio = 1 OR opcio = 1 => funcionament normal
'si opcio = 2 => funcionament per a llínies (NOTA: el manteniment de llinies ha d'estar dins d'un frameAux)
    Dim cadSQL2 As String
    Dim nom_camp As String
    Dim Control As Object
    Dim mTag As CTag
    Dim I As Integer
    
    Set mTag = New CTag
    cadSQL2 = ""

    For Each Control In formu.Controls
        If Control.Tag <> "" Then
            mTag.Cargar Control
            
            'If (mTag.Cargado) And (mTag.EsClave) And (InStr(1, Control.Container.Name, "FrameAux")) = 0 Then 'el control es clau primaria i no forma part de les llínies
            If (mTag.Cargado) And (mTag.EsClave) Then
                If (((opcio = 0) Or (opcio = 1)) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    For I = 0 To ado.Recordset.Fields.Count - 1
                        If mTag.columna = ado.Recordset.Fields(I).Name Then
                        
                            If cadSQL2 = "" Then
                                cadSQL2 = "{" & mTag.Tabla & "." & mTag.columna & "} = "
                            Else
                                cadSQL2 = cadSQL2 & " AND {" & mTag.Tabla & "." & mTag.columna & "} = "
                            End If
                            
                            If mTag.TipoDato = "T" Then 'text
                                cadSQL2 = cadSQL2 & "'" & ado.Recordset.Fields(I).Value & "'"
                            ElseIf mTag.TipoDato = "N" Then 'integer i decimal
                                cadSQL2 = cadSQL2 & ado.Recordset.Fields(I).Value
                            ElseIf mTag.TipoDato = "F" Then 'fecha
                                'cadSQL2 = cadSQL2 & "'" & ado.Recordset.Fields(i).Value & "'"
                                cadSQL2 = cadSQL2 & Date2SF("'" & ado.Recordset.Fields(I).Value & "'")
                            End If
                            
                            Exit For
                        End If
                    Next I
                End If
            End If
        End If
    Next Control
    
    POS2SF = cadSQL2

End Function

'convertix un SQL a una Selection Formula SF
Public Function SQL2SF(cadSQL As String) As String
    Dim cadSQL2 As String
    Dim posP As Integer 'posició del Punt
    
    cadSQL2 = cadSQL
    
    cadSQL2 = Replace(cadSQL2, "AND (1=1)", "") 'lleva el AND (1=1)  NOTA: açò ha d'estar abans de traure els parentesi
    cadSQL2 = Replace(cadSQL2, "%", "*") 'canvia el % per un *
    'cadSQL2 = Replace(cadSQL2, "_", "?") 'canvia el _ per un ?
    cadSQL2 = Replace(cadSQL2, "(", "{") 'canvia el ( per una {
    cadSQL2 = Replace(cadSQL2, ")", "") 'lleva )
    
    '+-+- per a posar el } +-+-
    posP = 0
    Do
        If posP = 0 Then
            posP = InStr(cadSQL2, ".")
        Else
            posP = InStr(posP + 1, cadSQL2, ".")
        End If
        If posP > 0 Then cadSQL2 = Left(cadSQL2, posP - 1) & Replace(cadSQL2, " ", "} ", posP, 1)
    Loop Until (InStr(posP + 1, cadSQL2, ".") = 0)
    '+-+-+-+-+-+-+-+-+-+-+-+-+
    
    '[Monica]16/09/2013: en caso de estar descompensados las llaves
    If InStrVeces(1, cadSQL2, "{") <> InStrVeces(1, cadSQL2, "}") Then
        cadSQL2 = Replace(Replace(cadSQL2, "{", " "), "}", " ")
        
        cadSQL2 = RellenaLlaves(cadSQL2)
    
    End If
    
    
    'per a canviar el format de les dates
    cadSQL2 = Date2SF(CStr(cadSQL2))
    
    'per a canviar el format de les dates
    cadSQL2 = Time2SF(CStr(cadSQL2))
    
    'per a canviar els _ per ?
    cadSQL2 = Like2SF(CStr(cadSQL2))
    
    cadSQL2 = NULL2SF(CStr(cadSQL2))
    
    SQL2SF = cadSQL2
End Function

Private Function RellenaLlaves(Cadena As String) As String
Dim Cad1 As String
Dim Cad2 As String
Dim I As Integer
Dim J As Integer

    Cad1 = Trim(Cadena)
    I = 1
    Do
        Cad2 = ""
        J = InStr(1, Cad1, " ")
        If J <> 0 Then Cad2 = Mid(Cad1, I, J - 1)
        If Cad2 <> "" Then
            Cadena = Replace(Cadena, Cad2, "{" & Cad2 & "}")
            Cad1 = InStr(J + 1, Cad1)
            I = J + 1
        End If
    Loop Until Cad2 = ""

    RellenaLlaves = Cadena
    
End Function



Private Function InStrVeces(Pos As Integer, Cad As String, Cad2 As String) As Integer
Dim Cad3 As String
Dim I As Integer
Dim J As Integer
Dim NVeces As Integer

    NVeces = 0
    Cad3 = Cad
    I = Len(Cad)
    J = InStr(1, Cad3, Cad2)
    
    While J <> 0
        NVeces = NVeces + 1
        If J + 1 <= I Then
            Cad3 = Mid(Cad3, J + 1)
        End If
        J = InStr(J, Cad3, Cad2)
    Wend
    
    InStrVeces = NVeces
    
End Function



'convertix una data per a pasar-li-la a una Selection Formula SF
' funciona tant en format 2005-01-17 com en 17/01/2005
Public Function Date2SF(cadData As String) As String
' DAVIDV [08/11/2006]: Cambios a causa del mal funcionamiento ocasionado por los criterios
' de búsqueda que contienen los carácteres - y /, y q no son fechas.
    Dim data, n_data As String
    Dim Pos As Integer, pos_desde, pos_hasta As Integer

    Pos = InStr(1, cadData, "-")
    While Pos <> 0
        pos_desde = InStrRev(cadData, "'", Pos) + 1
        pos_hasta = InStr(Pos, cadData, "'")
        data = Mid(cadData, pos_desde, pos_hasta - pos_desde)
        If IsDate(data) Then
          n_data = "Date(" & Year(data) & "," & Month(data) & "," & Day(data) & ")"
          cadData = Replace(cadData, "'" & data & "'", n_data)
          '-- LAURA: 27/04/2007
          Pos = InStr(Pos + 1, cadData, "-")
          '--
        Else
          Pos = InStr(Pos + 1, cadData, "-")
        End If
    Wend
    
    Pos = InStr(1, cadData, "/")
    While Pos <> 0
        pos_desde = InStrRev(cadData, "'", Pos) + 1
        pos_hasta = InStr(Pos, cadData, "'")
        data = Mid(cadData, pos_desde, pos_hasta - pos_desde)
        If IsDate(data) Then
          n_data = "Date(" & Year(data) & "," & Month(data) & "," & Day(data) & ")"
          cadData = Replace(cadData, "'" & data & "'", n_data)
          '-- LAURA: 27/04/2007
          Pos = InStr(Pos + 1, cadData, "/")
          '--
        Else
          Pos = InStr(Pos + 1, cadData, "/")
        End If
    Wend

'    While InStr(cadData, "-") <> 0
''        data = Mid(cadData, InStr(cadData, "-") - 5, 12) 'pa llevar les ' '
''        n_data = "Date(" & Mid(data, InStr(data, "-") - 4, 4) & "," & Mid(data, InStr(data, "-") + 1, 2)
''        n_data = n_data & "," & Mid(data, InStr(data, "-") + 4, 2) & ")"
''        cadData = Replace(cadData, data, n_data)
'        data = Mid(cadData, InStr(cadData, "-") - 5, 12) 'pa llevar les ' '
'        'data = cadData
'        n_data = "Date(" & Mid(data, InStr(data, "-") - 4, 4) & "," & Mid(data, InStr(data, "-") + 1, 2)
'        n_data = n_data & "," & Mid(data, InStr(data, "-") + 4, 2) & ")"
'        cadData = Replace(cadData, data, n_data)
'    Wend
    
'    While InStr(cadData, "/") <> 0
''        data = Mid(cadData, InStr(cadData, "/") - 2, 10) 'pa llevar les ' '
''        n_data = "Date(" & Mid(data, InStr(data, "/") + 4, 4) & "," & Mid(data, InStr(data, "/") + 1, 2)
''        n_data = n_data & "," & Mid(data, InStr(data, "/") - 2, 2) & ")"
''        cadData = Replace(cadData, data, n_data)
'
'        data = Mid(cadData, InStr(cadData, "/") - 2, 10) 'pa llevar les ' '
'        'data = cadData
'        n_data = "Date(" & Mid(data, InStr(data, "/") + 4, 4) & "," & Mid(data, InStr(data, "/") + 1, 2)
'        n_data = n_data & "," & Mid(data, InStr(data, "/") - 2, 2) & ")"
'        cadData = Replace(cadData, data, n_data)
'    Wend
    
    Date2SF = cadData
    
End Function

' funció per a llevar els _ de la cadena
' només els lleva de lo que hi haja entre ' ' i després de LIKE
'per a que no canvie el _ dels noms del camps
Public Function Like2SF(Cadena As String) As String
    Dim cadLike As String
    Dim cadTemp As String
    
    cadLike = Cadena
    
    While InStr(cadLike, "LIKE") <> 0
        cadLike = Mid(cadLike, InStr(cadLike, "LIKE") + 5, Len(cadLike) - 1)
        cadTemp = Mid(cadLike, 1, InStr(2, cadLike, "'"))
        Cadena = Replace(Cadena, cadTemp, Replace(cadTemp, "_", "?"))
    Wend
    
    Like2SF = Cadena
    
End Function


' funció per a llevar els nulls de la cadena
' només els lleva de lo que hi haja entre ' ' i després de LIKE
'per a que no canvie el _ dels noms del camps
Public Function NULL2SF(Cadena As String) As String
    Dim cadLike As String
    Dim cadTemp As String
    
    cadLike = Cadena
    
    If InStr(cadLike, "is NULL") <> 0 Then
        cadLike = Mid(cadLike, 1, InStr(1, cadLike, "is NULL") - 1)
        cadTemp = "isnull(" & Trim(cadLike) & ")"
        '[Monica]12/02/2014: cuando habia un campo con null para buscar no seleccionaba nada mas
        cadTemp = cadTemp & Mid(Cadena, InStr(1, Cadena, "is NULL") + 7)
        
        Cadena = cadTemp
    End If
    
    NULL2SF = Cadena
    
End Function



Public Function Time2SF(cadData As String) As String
' DAVIDV [08/11/2006]: Cambios a causa del mal funcionamiento ocasionado por los criterios
' de búsqueda que contienen los carácteres - y /, y q no son fechas.
    Dim data, n_data As String
    Dim Pos As Integer, pos_desde, pos_hasta As Integer

    Pos = InStr(1, cadData, ":")
    While Pos <> 0
        pos_desde = InStrRev(cadData, "'", Pos) + 1
        pos_hasta = InStr(Pos, cadData, "'")
        data = Mid(cadData, pos_desde, pos_hasta - pos_desde)
        If IsDate(data) Then
          n_data = "Time(" & Hour(data) & "," & Minute(data) & "," & Second(data) & ")"
          cadData = Replace(cadData, "'" & data & "'", n_data)
          '-- LAURA: 27/04/2007
          Pos = InStr(Pos + 1, cadData, ":")
          '--
        Else
          Pos = InStr(Pos + 1, cadData, ":")
        End If
    Wend
    
    
'    While InStr(cadData, "-") <> 0
''        data = Mid(cadData, InStr(cadData, "-") - 5, 12) 'pa llevar les ' '
''        n_data = "Date(" & Mid(data, InStr(data, "-") - 4, 4) & "," & Mid(data, InStr(data, "-") + 1, 2)
''        n_data = n_data & "," & Mid(data, InStr(data, "-") + 4, 2) & ")"
''        cadData = Replace(cadData, data, n_data)
'        data = Mid(cadData, InStr(cadData, "-") - 5, 12) 'pa llevar les ' '
'        'data = cadData
'        n_data = "Date(" & Mid(data, InStr(data, "-") - 4, 4) & "," & Mid(data, InStr(data, "-") + 1, 2)
'        n_data = n_data & "," & Mid(data, InStr(data, "-") + 4, 2) & ")"
'        cadData = Replace(cadData, data, n_data)
'    Wend
    
'    While InStr(cadData, "/") <> 0
''        data = Mid(cadData, InStr(cadData, "/") - 2, 10) 'pa llevar les ' '
''        n_data = "Date(" & Mid(data, InStr(data, "/") + 4, 4) & "," & Mid(data, InStr(data, "/") + 1, 2)
''        n_data = n_data & "," & Mid(data, InStr(data, "/") - 2, 2) & ")"
''        cadData = Replace(cadData, data, n_data)
'
'        data = Mid(cadData, InStr(cadData, "/") - 2, 10) 'pa llevar les ' '
'        'data = cadData
'        n_data = "Date(" & Mid(data, InStr(data, "/") + 4, 4) & "," & Mid(data, InStr(data, "/") + 1, 2)
'        n_data = n_data & "," & Mid(data, InStr(data, "/") - 2, 2) & ")"
'        cadData = Replace(cadData, data, n_data)
'    Wend
    
    Time2SF = cadData
    
End Function

