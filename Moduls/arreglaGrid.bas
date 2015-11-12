Attribute VB_Name = "arreglaGrid"
Public Sub arregla(ByRef tots As String, ByRef grid As DataGrid, ByRef formu As Form)
    'Dim tots As String
    Dim camp As String
    Dim mens As String
    Dim difer As Integer
    Dim i As Integer
    Dim k As Integer
    Dim posi As Integer
    Dim posi2 As Integer
    Dim fil As Integer
    Dim c As Integer
    Dim o As Integer
    Dim A() As Variant 'per als 5 parametres
    'Dim grid As DataGrid
    Dim obj As Object
    Dim obj_ant As Object
    Dim primer As Boolean
    Dim TotalAncho As Integer
    
    grid.AllowRowSizing = False
    grid.RowHeight = 290
    
    '***********
    difer = 563 'dirència recomanda entre l'ample del Datagrid i la suma dels amples de les columnes
    '***********
    
    TotalAncho = 0
    primer = False
'    Set grid = DataGrid1 'nom del DataGrid
    fil = -1 'fila a -1
    c = -1 'columna del datagrid a 0
    'tots = "S|txtAux(0)|T|Código|700|;S|txtAux(1)|T|Descripción|3000|;"
    
    While (tots <> "") 'bucle per a recorrer els distins camps
        Set obj = Nothing
        Set obj_ant = Nothing
    
        fil = fil + 1
        'ReDim Preserve A(6, fil)
        ReDim Preserve A(5, fil)
        'fila i columna a 0 (NOTA: les files es numeren a partir d'1 i les columnes a partir de 0)
        posi = InStr(tots, ";") '1ª posicio del ;
        camp = Left(tots, posi - 1)
        tots = Right(tots, Len(tots) - posi) 'lleve el camp actual
        'For k = 0 To 5
        For k = 0 To 4
          posi2 = InStr(camp, "|") '1ª posició del |
          A(k, fil) = Left(camp, posi2 - 1)
          camp = Right(camp, Len(camp) - posi2) 'lleve l'argument actual
        Next k 'quan acabe el for tinc en A el camp actual
        
        'només incremente el nº de la columna si no es un boto
        If A(2, fil) <> "B" Then c = c + 1
        
        If A(0, fil) = "N" Then 'no visible
            grid.Columns(c).visible = False
            grid.Columns(c).Width = 0 'si no es visible, pose a 0 l'ample
        ElseIf A(0, fil) = "S" Then 'visible
            ' ********* CAPTION I WIDTH DE L'OBJECTE ************
            
            Select Case A(2, fil) 'tipo (T, C o B) (o CB=CheckBox ) (DT=DTPicker)
                Case "T"
                    grid.Columns(c).visible = True
                    If A(3, fil) <> "" Then grid.Columns(c).Caption = A(3, fil)
                    If A(4, fil) <> "" Then grid.Columns(c).Width = CInt(A(4, fil))
'                    If A(5, fil) <> "" Then
'                        grid.Columns(c).NumberFormat = A(5, fil)
'                    Else
'                        grid.Columns(c).NumberFormat = ""
'                    End If
                    TotalAncho = TotalAncho + CInt(A(4, fil))
                Case "C"
                    grid.Columns(c).visible = True
                    If A(3, fil) <> "" Then grid.Columns(c).Caption = A(3, fil)
                    If A(4, fil) <> "" Then grid.Columns(c).Width = CInt(A(4, fil)) - 10
'                    If A(5, fil) <> "" Then
'                        grid.Columns(c).NumberFormat = A(5, fil)
'                    Else
'                        grid.Columns(c).NumberFormat = ""
'                    End If
                    TotalAncho = TotalAncho + CInt(A(4, fil))
                Case "B"
                
               '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                Case "CB"
                    grid.Columns(c).visible = True
                    If A(3, fil) <> "" Then grid.Columns(c).Caption = A(3, fil)
                    If A(4, fil) <> "" Then grid.Columns(c).Width = CInt(A(4, fil))
                    TotalAncho = TotalAncho + CInt(A(4, fil))
               '===============================================
               '=== LAURA (14/07/06): añadir tipo DT=DTPicker
                Case "DT"
                    grid.Columns(c).visible = True
                    If A(3, fil) <> "" Then grid.Columns(c).Caption = A(3, fil)
                    If A(4, fil) <> "" Then grid.Columns(c).Width = CInt(A(4, fil))
                    TotalAncho = TotalAncho + CInt(A(4, fil))
                '==============================================
            End Select
                       
            ' ********* CARREGUE L'OBJECTE ************
            Set obj = eval(formu, CStr(A(1, fil)))
            
            ' ********* NUMBERFORMAT i ALIGNMENT DE L'OBJECTE ************
            If (A(2, fil) = "T") Or (A(2, fil) = "C") Or (A(2, fil) = "DT") Then 'el numberformat només es per a text o combo
                If obj.Tag <> "" Then
                    grid.Columns(c).NumberFormat = FormatoCampo2(obj)
                    If TipoCamp(obj) = "N" Then
                        If (A(2, fil) = "T") Then _
                            grid.Columns(c).Alignment = dbgRight ' el Alignment només per a Text
                        grid.Columns(c).NumberFormat = grid.Columns(c).NumberFormat & " "
                    End If
                Else
                    grid.Columns(c).NumberFormat = ""
                End If
            End If
            
            ' ********* WIDTH I LEFT DE L'OBJECTE ************
            Select Case A(2, fil) 'tipo (T, C o B)
                Case "T"
                    If Not primer Then 'es el primer objecte visible
                        obj.Width = grid.Columns(c).Width - 60
                        'obj.Width = grid.Columns(c).Width - 8
                        obj.Left = grid.Left + 340
                        'obj.Left = grid.Left + 308
                    Else
                        o = 0
                        While obj_ant Is Nothing
                            o = o + 1
                            If A(0, fil - o) = "S" Then
                                Set obj_ant = eval(formu, CStr(A(1, fil - o)))
                            End If
                        Wend
                        'en obj_ant tindré el 1r objecte per darrere que siga visible
                        Select Case A(2, fil - o)
                            Case "T" 'objecte anterior a text es text
                                obj.Width = grid.Columns(c).Width - 60
                                'obj.Width = grid.Columns(c).Width - 38
                                obj.Left = obj_ant.Left + obj_ant.Width + 60
                                'obj.Left = obj_ant.Left + obj_ant.Width + 38
                            Case "C" 'objecte anterior a text es combo
                                obj.Width = grid.Columns(c).Width - 60
                                obj.Left = obj_ant.Left + obj_ant.Width + 30
                            Case "B" 'objecte anterior a text es un boto
                                obj.Width = grid.Columns(c).Width - 60
                                obj.Left = obj_ant.Left + obj_ant.Width + 30
                                
                             '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                            Case "CB" 'anterior es CheckBox
                                obj.Width = grid.Columns(c).Width - 60
                                obj.Left = obj_ant.Left + obj_ant.Width + 60
                            '=== LAURA (14/07/06): añadir tipo DT=DTPicker
                            Case "DT" 'anterior es un DTPicker
                                obj.Width = grid.Columns(c).Width - 60
                                obj.Left = obj_ant.Left + obj_ant.Width + 60
                        End Select
                    End If
                    
                Case "C"
                    If Not primer Then 'es el primer objecte visible
                        obj.Width = grid.Columns(c).Width - 10
                        obj.Left = grid.Left + 320
                    Else
                        o = 0
                        While obj_ant Is Nothing
                            o = o + 1
                            If A(0, fil - o) = "S" Then
                                Set obj_ant = eval(formu, CStr(A(1, fil - o)))
                            End If
                        Wend
                        'en obj_ant tindré el 1r objecte per darrere que siga visible
                        Select Case A(2, fil - o)
                            Case "T" 'objecte anterior a combo es text
                                obj.Width = grid.Columns(c).Width - 20
                                obj.Left = obj_ant.Left + obj_ant.Width + 40
                            Case "C" 'objecte anterior a combo es combo
                                obj.Width = grid.Columns(c).Width
                                obj.Left = obj_ant.Left + obj_ant.Width
                            Case "B" 'objecte anterior a combo es un boto
                                ' *** FALTA PER A QUAN L'OBJECTE ANTERIOR A UN COMBO ES UN BOTO
'                                mens = "Falta programar en arreglaGrid per al cas que un l'objecte anterior a un ComboBox es un Button"
'                                MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & mens
                                '=== LAURA (14/09/06): añadir este caso
                                obj.Width = grid.Columns(c).Width
                                obj.Left = obj_ant.Left + obj_ant.Width + 10
                            
                            '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                            Case "CB" 'anterior es CheckBox (falta comprobar)
                                obj.Width = grid.Columns(c).Width
                                obj.Left = obj_ant.Left + obj_ant.Width
                        End Select
                    End If
                    
                Case "B"
                    If Not primer Then 'es el primer objecte visible
                        ' *** FALTA PER A QUAN UN BOTO ES EL PRIMER OBJECTE VISIBLE
                        mens = "Falta programar en arreglaGrid per al cas que un Button es el primer objete visible d'un Datagrid"
                        MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & mens
                    Else
                        o = 0
                        While obj_ant Is Nothing
                            o = o + 1
                            If A(0, fil - o) = "S" Then
                                Set obj_ant = eval(formu, CStr(A(1, fil - o)))
                            End If
                        Wend
                        'en obj_ant tindré el 1r objecte per darrere que siga visible
                        Select Case A(2, fil - o)
                            Case "T" 'objecte anterior a boto es text
                                obj_ant.Width = obj_ant.Width - obj.Width + 30 '1r faig més curt l'objecte de text
                                obj.Left = obj_ant.Left + obj_ant.Width
                                'obj.Left = obj_ant.Left + obj_ant.Width - obj.Width
                            Case "C" 'objecte anterior a boto es combo
                                ' *** FALTA PER A QUAN L'OBJECTE ANTERIOR A UN BOTO ES UN COMBO
                                mens = "Falta programar en arreglaGrid per al cas que un l'objecte anterior a un Button es un ComboBox"
                                MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & mens
                            Case "B" 'objecte anterior a combo es un boto
                                ' *** FALTA PER A QUAN L'OBJECTE ANTERIOR A UN BOTO ES UN BOTO
                                mens = "Falta programar en arreglaGrid per al cas que un l'objecte anterior a un Button es un Button"
                                MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & mens
                        End Select
                    End If
                    
                 '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                Case "CB"
                    If Not primer Then 'es el primer objecte visible
                        obj.Width = grid.Columns(c).Width - 10
                        obj.Left = grid.Left + 320
                    Else
                        o = 0
                        While obj_ant Is Nothing
                            o = o + 1
                            If A(0, fil - o) = "S" Then
                                Set obj_ant = eval(formu, CStr(A(1, fil - o)))
                            End If
                        Wend
                        'en obj_ant tindré el 1r objecte per darrere que siga visible
                        Select Case A(2, fil - o)
                            Case "T" 'objecte anterior a combo es text
                                obj.Width = grid.Columns(c).Width - (grid.Columns(c).Width / 3)
                                obj.Left = obj_ant.Left + obj_ant.Width + (grid.Columns(c).Width / 3) - 10
                            Case "C" 'objecte anterior a combo es combo
                                obj.Width = grid.Columns(c).Width
                                obj.Left = obj_ant.Left + obj_ant.Width
                            Case "B" 'objecte anterior a combo es un boto
                                ' *** FALTA PER A QUAN L'OBJECTE ANTERIOR A UN COMBO ES UN BOTO
'Laura: 140508
'                                mens = "Falta programar en arreglaGrid per al cas que un l'objecte anterior a un ComboBox es un Button"
'                                MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & mens
                                
                                obj.Width = grid.Columns(c).Width
                                obj.Left = obj_ant.Left + obj_ant.Width + 10
                                
                             '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                            Case "CB" 'anterior es un ChekBox
                                obj.Width = grid.Columns(c).Width - (grid.Columns(c).Width / 3)
                                obj.Left = obj_ant.Left + obj_ant.Width + (grid.Columns(c).Width / 3)
                        End Select
                    End If
                
                
                 '=== LAURA (14/07/06): añadir tipo DT=DTPicker
                Case "DT"
                    If Not primer Then 'es el primer objecte visible
                        obj.Width = grid.Columns(c).Width - 10
                        obj.Left = grid.Left + 320
                    Else
                        o = 0
                        While obj_ant Is Nothing
                            o = o + 1
                            If A(0, fil - o) = "S" Then
                                Set obj_ant = eval(formu, CStr(A(1, fil - o)))
                            End If
                        Wend
                        'en obj_ant tindré el 1r objecte per darrere que siga visible
                        Select Case A(2, fil - o)
                            Case "T" 'objecte anterior a combo es text
                                obj.Width = grid.Columns(c).Width - 40
                                obj.Left = obj_ant.Left + obj_ant.Width + 40
                            Case "C" 'objecte anterior a combo es combo
                                obj.Width = grid.Columns(c).Width
                                obj.Left = obj_ant.Left + obj_ant.Width
                            Case "B" 'objecte anterior a combo es un boto
                                ' *** FALTA PER A QUAN L'OBJECTE ANTERIOR A UN COMBO ES UN BOTO
                                mens = "Falta programar en arreglaGrid per al cas que un l'objecte anterior a un ComboBox es un Button"
                                MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & mens
                                
                             '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                            Case "CB" 'anterior es un ChekBox
                                obj.Width = grid.Columns(c).Width - (grid.Columns(c).Width / 3)
                                obj.Left = obj_ant.Left + obj_ant.Width + (grid.Columns(c).Width / 3)
                            Case "DT" 'anterior es un DTPicker
                                obj.Width = grid.Columns(c).Width - 40
                                obj.Left = obj_ant.Left + obj_ant.Width + 40
                        End Select
                    End If
                Case Else
                    MsgBox "No existix el tipo de control " & A(2, fil)
            End Select
            
        primer = True
        End If
                
    Wend

    'No permitir canviar tamany de columnes
    For i = 0 To grid.Columns.Count - 1
         grid.Columns(i).AllowSizing = False
    Next i

'    If grid.Width - TotalAncho <> difer Then
'        mens = "Es recomana que el total d'amples de les columnes per a este DataGrid siga de "
'        mens = mens & CStr(grid.Width - difer)
'        MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & mens
'    End If
End Sub

Public Function eval(ByRef formu As Form, nom_camp As String) As Control
Dim Ctrl As Control
Dim nom_camp2 As String
Dim nou_i As Integer
Dim j As Integer

    Set eval = Nothing
    j = InStr(1, nom_camp, "(")
    If j = 0 Then
        nou_i = -1
    Else
        nom_camp = Left(nom_camp, Len(nom_camp) - 1)
        nou_i = Val(Mid(nom_camp, j + 1))
        nom_camp = Left(nom_camp, j - 1)
    End If
    
    For Each Ctrl In formu.Controls
        If Ctrl.Name = nom_camp Then
            If nou_i >= 0 Then
                If nou_i = Ctrl.Index Then
                    j = 1 'coincidix el nom i l'index
                Else
                    j = 0 'coincidix el nom però no l'index
                End If
            Else
                j = 1 'coincidix el nom i no te index
            End If
        Else
            j = -1 'no coincidix el nom
        End If
        
        If j > 0 Then
            Set eval = Ctrl
            Exit For
        End If
    Next Ctrl
End Function
