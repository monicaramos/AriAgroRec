Attribute VB_Name = "ModFunciones"
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
'   En este modulo estan las funciones que recorren el form
'   usando el each for
'   Estas son
'
'   CamposSiguiente -> Nos devuelve el el text siguiente en
'           el orden del tabindex
'
'   CompForm -> Compara los valores con su tag
'
'   InsertarDesdeForm - > Crea el sql de insert e inserta
'
'   Limpiar -> Pone a "" todos los objetos text de un form
'
'   ObtenerBusqueda -> A partir de los text crea el sql a
'       partir del WHERE ( sin el).
'
'   ModifcarDesdeFormulario -> Opcion modificar. Genera el SQL
'
'   PonerDatosForma -> Pone los datos del RECORDSET en el form
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
Option Explicit


Public Const ValorNulo = "Null"
Public NombreCheck As String

Public Function CompForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    Dim HayCamposIncorrectos As Boolean
    Dim CampoIncorrecto As String

    HayCamposIncorrectos = False
    CampoIncorrecto = ""


    CompForm = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox And Control.visible = True Then
            Carga = mTag.Cargar(Control)
            If Carga = True Then
                Correcto = mTag.Comprobar(Control, True)
                If Not Correcto Then
                    Control.BackColor = vbErrorColor
                    HayCamposIncorrectos = True
                    CampoIncorrecto = Control.Name
                    If IsArray(Control) Then CampoIncorrecto = CampoIncorrecto & "(" & Control.Index & ")"
                Else
                    Control.BackColor = vbWhite
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                Carga = mTag.Cargar(Control)
                If Carga = False Then
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function

                Else
                    If mTag.Vacio = "N" And Control.ListIndex < 0 Then
'                            MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
'                            Exit Function
                        Control.BackColor = vbErrorColor
                        HayCamposIncorrectos = True
                        CampoIncorrecto = Control.Name
                        If IsArray(Control) Then CampoIncorrecto = CampoIncorrecto & "(" & Control.Index & ")"
                    Else
                        Control.BackColor = vbWhite

                    End If
                End If
            End If
        End If
    Next Control
    If HayCamposIncorrectos Then
        MsgBox "Revise datos obligatorios o incorrectos", vbExclamation
    End If
    CompForm = Not HayCamposIncorrectos
    
'    CompForm = True
End Function

'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function CompForm2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    Dim HayCamposIncorrectos As Boolean
    Dim CampoIncorrecto As String
    
    HayCamposIncorrectos = False
    CampoIncorrecto = ""

    CompForm2 = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox Then
            If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                Carga = mTag.Cargar(Control)
                If Carga = True Then
                    Correcto = mTag.Comprobar(Control, True)
                    If Not Correcto Then
                        Control.BackColor = vbErrorColor
                        HayCamposIncorrectos = True
                        CampoIncorrecto = Control.Name
                        If IsArray(Control) Then CampoIncorrecto = CampoIncorrecto & "(" & Control.Index & ")"
                    Else
                        If Control.Tag <> "" Then Control.BackColor = vbWhite
                    End If
                    
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.visible = True Then
                'Comprueba que los campos estan bien puestos
                If Control.Tag <> "" Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        Carga = mTag.Cargar(Control)
                        If Carga = False Then
                            MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                            Exit Function
        
                        Else
                            If mTag.Vacio = "N" And Control.ListIndex < 0 Then
'                                    MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
'                                    Exit Function
                                Control.BackColor = vbErrorColor
                                HayCamposIncorrectos = True
                                CampoIncorrecto = Control.Name
                                If IsArray(Control) Then CampoIncorrecto = CampoIncorrecto & "(" & Control.Index & ")"

                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next Control
    
    If HayCamposIncorrectos Then
        MsgBox "Revise datos obligatorios o incorrectos", vbExclamation
    End If
    CompForm2 = Not HayCamposIncorrectos
    
'    CompForm2 = True
End Function




'Public Function CampoSiguiente(ByRef formulario As Form, valor As Integer) As Control
'Dim Fin As Boolean
'Dim Control As Object
'
'On Error GoTo ECampoSiguiente
'
'    'Debug.Print "Llamada:  " & Valor
'    'Vemos cual es el siguiente
'    Do
'        valor = valor + 1
'        For Each Control In formulario.Controls
'            'Debug.Print "-> " & Control.Name & " - " & Control.TabIndex
'            'Si es texto monta esta parte de sql
'            If Control.TabIndex = valor Then
'                    Set CampoSiguiente = Control
'                    Fin = True
'                    Exit For
'            End If
'        Next Control
'        If Not Fin Then
'            valor = -1
'        End If
'    Loop Until Fin
'    Exit Function
'ECampoSiguiente:
'    Set CampoSiguiente = Nothing
'    Err.Clear
'End Function



'-----------------------------------
Public Function ValorParaSQL(Valor, ByRef vtag As CTag) As String
Dim Dev As String
Dim D As Single
Dim I As Integer
Dim V
    Dev = ""
    If Valor <> "" Then
        Select Case vtag.TipoDato
        Case "N"
            V = Valor
            If InStr(1, Valor, ",") Or InStr(1, Valor, ".") Then
                If InStr(1, Valor, ".") Then
                    'ABRIL 2004

                    'Ademas de la coma lleva puntos
                    V = ImporteFormateado(CStr(Valor))
                    Valor = V
                Else

                    V = CSng(Valor)
                    Valor = V
                End If
            Else

            End If
            Dev = TransformaComasPuntos(CStr(Valor))

        Case "F"
            Dev = "'" & Format(Valor, FormatoFecha) & "'"
            
        Case "H"
            Dev = "'" & Format(Valor, "hh:mm:ss") & "'"
        
        Case "FHH"
            Dev = DBSet(Valor, "FH")
            
        Case "FH"
            Dev = DBSet(Valor, "FH")
        Case "T"
            Dev = CStr(Valor)
            NombreSQL Dev
            Dev = "'" & Dev & "'"
        Case Else
            Dev = "'" & Valor & "'"
        End Select

    Else
        'Si se permiten nulos, la "" ponemos un NULL
        If vtag.Vacio = "S" Then
            Dev = ValorNulo
        End If
    End If
    ValorParaSQL = Dev
End Function


Public Function InsertarDesdeForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim cad As String
    
    On Error GoTo EInsertarF
    
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm = False
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                    
                        'Parte VALUES
                        cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.columna & ""
                If Control.Value = 1 Then
                    cad = "1"
                    Else
                    cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                Der = Der & cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.ListIndex = -1 Then
                        cad = ValorNulo
                    Else
                        cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    cad = "INSERT INTO " & mTag.tabla
    cad = cad & " (" & Izda & ") VALUES (" & Der & ");"
    
    conn.Execute cad, , adCmdText
    
    InsertarDesdeForm = True
Exit Function

EInsertarF:
    MuestraError Err.Number, "Inserta. " & Err.Description
End Function


Public Function InsertarDesdeForm2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim cad As String
    
    On Error GoTo EInsertarF
    
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm2 = False
    Der = ""
    Izda = ""
    
    For Each Control In formulario.Controls
    
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            If Izda <> "" Then Izda = Izda & ","
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.columna & ""
                        
                            'Parte VALUES
                            cad = ValorParaSQL(Control.Text, mTag)
                            If Der <> "" Then Der = Der & ","
                            Der = Der & cad
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Izda <> "" Then Izda = Izda & ","
                    'Access
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.Value = 1 Then
                        cad = "1"
                        Else
                        cad = "0"
                    End If
                    If Der <> "" Then Der = Der & ","
                    If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                    Der = Der & cad
                End If
            End If
            
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                        If Control.ListIndex = -1 Then
                            cad = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            cad = Control.ItemData(Control.ListIndex)
                        Else
                            cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
            
        'OPTION BUTTON
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            If Izda <> "" Then Izda = Izda & ","
                            Izda = Izda & "" & mTag.columna & ""
                            cad = Control.Index
                            If Der <> "" Then Der = Der & ","
                            Der = Der & cad
                        End If
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is DTPicker Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
'                        If Control.Value Then
'                            If Izda <> "" Then Izda = Izda & ","
'                            Izda = Izda & "" & mTag.columna & ""
'                            cad = Control.index
'                            If Der <> "" Then Der = Der & ","
'                            Der = Der & cad
'                        End If
                        If Izda <> "" Then Izda = Izda & ","
                        Izda = Izda & "" & mTag.columna & ""
                        
                        'Parte VALUES
                        If Control.visible Then
                            cad = ValorParaSQL(Control.Value, mTag)
                        Else
                            cad = ValorNulo
                        End If
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
        End If
        
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    cad = "INSERT INTO " & mTag.tabla
    cad = cad & " (" & Izda & ") VALUES (" & Der & ");"
    conn.Execute cad, , adCmdText
    
     ' ### [Monica] 18/12/2006
    CadenaCambio = cad
   
    InsertarDesdeForm2 = True
Exit Function

EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function


Public Function CadenaInsertarDesdeForm(ByRef formulario As Form) As String
'Equivale a InsertarDesdeForm, excepto que devuelve la candena SQL y hace el execute fuera de la función.
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    CadenaInsertarDesdeForm = ""
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.visible = True Then
                If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            If Izda <> "" Then Izda = Izda & ","
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.columna & ""
                        
                            'Parte VALUES
                            cad = ValorParaSQL(Control.Text, mTag)
                            If Der <> "" Then Der = Der & ","
                            Der = Der & cad
                        End If
                    End If
                End If
           End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.visible = True Then
                If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If Izda <> "" Then Izda = Izda & ","
                    'Access
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.Value = 1 Then
                        cad = "1"
                        Else
                        cad = "0"
                    End If
                    If Der <> "" Then Der = Der & ","
                    If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                    Der = Der & cad
                End If
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.visible = True Then
                If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                        If Control.ListIndex = -1 Then
                            cad = ValorNulo
                        Else
                            cad = Control.ItemData(Control.ListIndex)
                        End If
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    cad = "INSERT INTO " & mTag.tabla
    cad = cad & " (" & Izda & ") VALUES (" & Der & ");"
'    Conn.Execute cad, , adCmdText
    
    CadenaInsertarDesdeForm = cad
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function


Public Function CadenaModificarDesdeForm(ByRef formulario As Form, vTabla As String) As String
'Equivale a InsertarDesdeForm, excepto que devuelve la candena SQL y hace el execute fuera de la función.
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim cad As String
    
    On Error GoTo EInsertarF
    
    'Exit Function
    Set mTag = New CTag
    CadenaModificarDesdeForm = ""
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.visible = True Then
                If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If UCase(mTag.tabla) = UCase(vTabla) Then
                            If mTag.columna <> "" Then
                                If Izda <> "" Then Izda = Izda & ","
                                'Access
                                'Izda = Izda & "[" & mTag.Columna & "]"
                                Izda = Izda & "" & mTag.columna & "="
                            
                                cad = ValorParaSQL(Control.Text, mTag)
                                Izda = Izda & cad
                                '++
                                If mTag.EsClave Then
                                    If Der <> "" Then Der = Der & " AND "
                                    Der = Der & mTag.columna & "=" & cad
                                End If
                                    
                            End If
                        End If
                    End If
                End If
           End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.visible = True Then
                If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If UCase(mTag.tabla) = UCase(vTabla) Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & "="
                        If Control.Value = 1 Then
                            cad = "1"
                            Else
                            cad = "0"
                        End If
                        If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                        Izda = Izda & cad
                    
                        '++
                        If mTag.EsClave Then
                            If Der <> "" Then Der = Der & " AND "
                            Der = Der & mTag.columna & "=" & cad
                        End If
                    End If
                End If
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.visible = True Then
                If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If UCase(mTag.tabla) = UCase(vTabla) Then
                            If Izda <> "" Then Izda = Izda & ","
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.columna & "="
                            If Control.ListIndex = -1 Then
                                cad = ValorNulo
                            Else
                                cad = Control.ItemData(Control.ListIndex)
                            End If
                            Izda = Izda & cad
                        
                            '++
                            If mTag.EsClave Then
                                If Der <> "" Then Der = Der & " AND "
                                Der = Der & mTag.columna & "=" & cad
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    cad = "UPDATE " & vTabla & " SET "
    cad = cad & Izda & " WHERE " & Der & ""
'    Conn.Execute cad, , adCmdText
    
    CadenaModificarDesdeForm = cad
Exit Function
EInsertarF:
    MuestraError Err.Number, "Modificar. "
End Function






Public Function PonerCamposForma(ByRef formulario As Form, ByRef vData As Adodc) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim cad As String
Dim Valor As Variant
Dim campo As String  'Campo en la base de datos
Dim I As Integer


    Set mTag = New CTag
    PonerCamposForma = False

    For Each Control In formulario.Controls
        'TEXTO
        If (TypeOf Control Is TextBox) And (Control.visible = True) And UCase(Control.Name) = "TEXT1" Then
'                If TypeOf control Is TextBox Then

            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                    'Columna en la BD
                    If mTag.columna <> "" Then
                        campo = mTag.columna
                        If mTag.Vacio = "S" Then
                            Valor = DBLet(vData.Recordset.Fields(campo))
                        Else
                            Valor = vData.Recordset.Fields(campo)
                        End If
                        If mTag.Formato <> "" And CStr(Valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                cad = Format(Valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = cad
                            Else
                                Control.Text = Format(Valor, mTag.Formato)
                            End If
                        Else
                            Control.Text = Valor
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf (TypeOf Control Is CheckBox) And (Control.visible = True) Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    campo = mTag.columna
                    Valor = vData.Recordset.Fields(campo)
                    Else
                        Valor = 0
                End If
                If IsNull(Valor) Then Valor = 0
                Control.Value = Valor
            End If

         'COMBOBOX
         ElseIf (TypeOf Control Is ComboBox) And Control.visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    campo = mTag.columna
                    Valor = DBLet(vData.Recordset.Fields(campo))
                    '++MONICA: 15-01-2008 añadida la condicion de que el valor sea nulo
                    If IsNull(vData.Recordset.Fields(campo)) Then
                        Control.ListIndex = -1
                    Else
                    '++
                        I = 0
                        For I = 0 To Control.ListCount - 1
                            If Control.ItemData(I) = Val(Valor) Then
                                Control.ListIndex = I
                                Exit For
                            End If
                        Next I
                        If I = Control.ListCount Then Control.ListIndex = -1
                    '++MONICA: 15-01-2008 añadida la condicion de que el valor sea nulo
                    End If
                    '++
                End If 'de cargado
            End If 'de <>""
        End If
    Next Control

    'Veremos que tal
    PonerCamposForma = True
Exit Function
EPonerCamposForma:
    MuestraError Err.Number, "Poner campos formulario. "
End Function

'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function PonerCamposForma2(ByRef formulario As Form, ByRef vData As Adodc, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim cad As String
Dim Valor As Variant
Dim campo As String  'Campo en la base de datos
Dim I As Integer
    On Error GoTo EPonerCamposForma2
    
    Set mTag = New CTag
    PonerCamposForma2 = False
    For Each Control In formulario.Controls
        'TEXTO
        If (TypeOf Control Is TextBox) Then
            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    If mTag.Cargado Then
                        'Columna en la BD
                        If mTag.columna <> "" Then
                            campo = mTag.columna
                            If mTag.Vacio = "S" Then
                                Valor = DBLet(vData.Recordset.Fields(campo))
                            Else
                                Valor = vData.Recordset.Fields(campo)
                            End If
                            If mTag.Formato <> "" And CStr(Valor) <> "" Then
                                If mTag.TipoDato = "N" Then
                                    'Es numerico, entonces formatearemos y sustituiremos
                                    ' La coma por el punto
                                    cad = Format(Valor, mTag.Formato)
                                    'Antiguo
                                    'Control.Text = TransformaComasPuntos(cad)
                                    'nuevo
                                    Control.Text = cad
                                Else
                                    Control.Text = Format(Valor, mTag.Formato)
                                End If
                            Else
                                Control.Text = Valor
                            End If
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf (TypeOf Control Is CheckBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Columna en la BD
                        campo = mTag.columna
                        Valor = vData.Recordset.Fields(campo)
                    Else
                        Valor = 0
                    End If
                    If IsNull(Valor) Then Valor = 0
                    Control.Value = Valor
                End If
            End If

         'COMBOBOX
         ElseIf (TypeOf Control Is ComboBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        campo = mTag.columna
                        Valor = DBLet(vData.Recordset.Fields(campo))
                        I = 0
                        For I = 0 To Control.ListCount - 1
                            If Control.ItemData(I) = Val(Valor) Then
                                Control.ListIndex = I
                                Exit For
                            End If
                        Next I
                        If I = Control.ListCount Then Control.ListIndex = -1
                    End If 'de cargado
                End If
            End If 'de <>""
            
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Columna en la BD
                        campo = mTag.columna
                        Valor = vData.Recordset.Fields(campo)
                        If IsNull(Valor) Then Valor = 0
                        If Control.Index = Valor Then
                            Control.Value = True
                        Else
                            Control.Value = False
                        End If
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is DTPicker Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Columna en la BD
                        campo = mTag.columna
                        Valor = vData.Recordset.Fields(campo)
                        If IsNull(Valor) Then Valor = Now
                        Control.Value = Format(Valor, mTag.Formato)
                    End If
                End If
            End If
        End If
    Next Control

    'Veremos que tal
    PonerCamposForma2 = True
Exit Function
EPonerCamposForma2:
    MuestraError Err.Number, "Poner campos formulario 2. "
End Function


Public Function ForaGrid(ByRef formulari As Form, ByRef vGrid As DataGrid, Control As Object) As Boolean
Dim mTag As CTag
Dim cad As String
Dim Valor As Variant
Dim camp As String  'Camp en la BDA
Dim I As Integer

    Set mTag = New CTag
    ForaGrid = False

    If (TypeOf Control Is TextBox) Then 'text
        mTag.Cargar Control
        If Control.Tag <> "" Then
            If mTag.Cargado Then
                If mTag.columna <> "" Then
                    camp = mTag.columna
                    If mTag.Vacio = "S" Then
                        Valor = DBLet(vGrid.Columns(camp).Text)
                        'valor = DBLet(vGrid.Recordset.Fields(campo))
                    Else
                        'valor = vGrid.Columns!camp
                        Valor = vGrid.Columns(camp).Text
                    End If
                    If mTag.Formato <> "" And CStr(Valor) <> "" Then
                        If mTag.TipoDato = "N" Then
                            cad = Format(Valor, mTag.Formato)
                            Control.Text = cad
                        Else
                            Control.Text = Format(Valor, mTag.Formato)
                        End If
                    Else
                        Control.Text = Valor
                    End If
                End If
            End If
        End If

'        'CheckBOX
'        ElseIf (TypeOf Control Is CheckBox) Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado Then
'                        'Columna en la BD
'                        campo = mTag.columna
'                        valor = vData.Recordset.Fields(campo)
'                        Else
'                            valor = 0
'                    End If
'                    If IsNull(valor) Then valor = 0
'                    Control.Value = valor
'                End If
'            End If
'
'         'COMBOBOX
'         ElseIf (TypeOf Control Is ComboBox) Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado Then
'                        campo = mTag.columna
'                        valor = DBLet(vData.Recordset.Fields(campo))
'                        i = 0
'                        For i = 0 To Control.ListCount - 1
'                            If Control.ItemData(i) = Val(valor) Then
'                                Control.ListIndex = i
'                                Exit For
'                            End If
'                        Next i
'                        If i = Control.ListCount Then Control.ListIndex = -1
'                    End If 'de cargado
'                End If
'            End If 'de <>""
    End If

    'Veremos que tal
    ForaGrid = True
Exit Function
EPosarCampsGrid:
    MuestraError Err.Number, "Poner campos grid. "
End Function


'Public Function PonerCamposFormaFrame(ByRef formulario As Form, NomTxtBox As String, ByRef vData As Adodc, Optional NomCheck As String, Optional NomCombo As String) As Boolean
'Dim Control As Object
'Dim mTag As CTag
'Dim cad As String
'Dim valor As Variant
'Dim campo As String  'Campo en la base de datos
'Dim i As Integer
'
'    Set mTag = New CTag
'    PonerCamposFormaFrame = False
'
'
'        For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox And Control.Visible = True And Control.Name = NomTxtBox Then
'            'Comprobamos que tenga tag
'            mTag.Cargar Control
''            Debug.Print Control.Parent
'            If Control.Tag <> "" Then
'                If mTag.Cargado Then
'                    'Columna en la BD
'                    If mTag.Columna <> "" Then
'                        campo = mTag.Columna
'                        If mTag.Vacio = "S" Then
'                            valor = DBLet(vData.Recordset.Fields(campo))
'                        Else
'                            valor = vData.Recordset.Fields(campo)
'                        End If
'                        If mTag.Formato <> "" And CStr(valor) <> "" Then
'                            If mTag.TipoDato = "N" Then
'                                'Es numerico, entonces formatearemos y sustituiremos
'                                ' La coma por el punto
'                                cad = Format(valor, mTag.Formato)
'                                'Antiguo
'                                'Control.Text = TransformaComasPuntos(cad)
'                                'nuevo
'                                Control.Text = cad
'                            Else
'                                Control.Text = Format(valor, mTag.Formato)
'                            End If
'                        Else
'                            Control.Text = valor
'                        End If
'                    End If
'                End If
'            End If
'        'CheckBOX
'        ElseIf TypeOf Control Is CheckBox And Control.Visible = True And Control.Name = NomCheck Then
'            If Control.Tag <> "" Then
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    'Columna en la BD
'                    campo = mTag.Columna
'                    valor = vData.Recordset.Fields(campo)
'                    Else
'                        valor = 0
'                End If
'                Control.Value = valor
'            End If
'
'         'COMBOBOX
'         ElseIf TypeOf Control Is ComboBox And Control.Visible = True And Control.Name = NomCombo Then
'            If Control.Tag <> "" Then
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    campo = mTag.Columna
'                    valor = vData.Recordset.Fields(campo)
'                    i = 0
'                    For i = 0 To Control.ListCount - 1
'                        If Control.ItemData(i) = Val(valor) Then
'                            Control.ListIndex = i
'                            Exit For
'                        End If
'                    Next i
'                    If i = Control.ListCount Then Control.ListIndex = -1
'                End If 'de cargado
'            End If 'de <>""
'        End If
'
'    Next Control
'
'    'Veremos que tal
'    PonerCamposFormaFrame = True
'Exit Function
'EPonerCamposForma:
'    MuestraError Err.Number, "Poner campos formulario. "
'End Function


Private Function ObtenerMaximoMinimo(vSql As String, Optional vBD As Byte) As String
Dim Rs As Recordset
    ObtenerMaximoMinimo = ""
    Set Rs = New ADODB.Recordset
    If vBD = cConta Then
        Rs.Open vSql, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else
        Rs.Open vSql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    End If
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            ObtenerMaximoMinimo = CStr(Rs.Fields(0))
        End If
    End If
    Rs.Close
    Set Rs = Nothing
End Function


'====DAVID
'Public Function ObtenerBusqueda(ByRef formulario As Form) As String
'    Dim Control As Object
'    Dim Carga As Boolean
'    Dim mTag As CTag
'    Dim Aux As String
'    Dim cad As String
'    Dim SQL As String
'    Dim tabla As String
'    Dim RC As Byte
'
'    On Error GoTo EObtenerBusqueda
'
'    'Exit Function
'    Set mTag = New CTag
'    ObtenerBusqueda = ""
'    SQL = ""
'
'    'Recorremos los text en busca de ">>" o "<<"
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            Aux = Trim(Control.Text)
'            If Aux = ">>" Or Aux = "<<" Then
'                Carga = mTag.Cargar(Control)
'                If Carga Then
'                    If Aux = ">>" Then
'                        cad = " MAX(" & mTag.Columna & ")"
'                    Else
'                        cad = " MIN(" & mTag.Columna & ")"
'                    End If
'                    SQL = "Select " & cad & " from " & mTag.tabla
'                    SQL = ObtenerMaximoMinimo(SQL)
'                    Select Case mTag.TipoDato
'                    Case "N"
'                        SQL = mTag.tabla & "." & mTag.Columna & " = " & TransformaComasPuntos(SQL)
'                    Case "F"
'                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
'                    Case Else
'                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & SQL & "'"
'                    End Select
'                    SQL = "(" & SQL & ")"
'                End If
'            End If
'        End If
'    Next
'
'
'
'    'Recorremos los text en busca del NULL
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            Aux = Trim(Control.Text)
'            If UCase(Aux) = "NULL" Then
'                Carga = mTag.Cargar(Control)
'                If Carga Then
'
'                    SQL = mTag.tabla & "." & mTag.Columna & " is NULL"
'                    SQL = "(" & SQL & ")"
'                    Control.Text = ""
'                End If
'            End If
'        End If
'    Next
'
'
'
'    'Recorremos los textbox
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            'Cargamos el tag
'            Carga = mTag.Cargar(Control)
'            If Carga Then
'                If mTag.Cargado Then
'                    Aux = Trim(Control.Text)
'                    If Aux <> "" Then
'                        If mTag.tabla <> "" Then
'                            tabla = mTag.tabla & "."
'                        Else
'                            tabla = ""
'                        End If
'                    RC = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.Columna, Aux, cad)
'                    If RC = 0 Then
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'            Else
'                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
'                Exit Function
'            End If
'
'        'COMBO BOX
'        ElseIf TypeOf Control Is ComboBox Then
'            mTag.Cargar Control
'            If mTag.Cargado Then
'                If Control.ListIndex > -1 Then
'                    If mTag.TipoDato <> "T" Then
'                        cad = Control.ItemData(Control.ListIndex)
'                        cad = mTag.tabla & "." & mTag.Columna & " = " & cad
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    Else
'                        cad = Control.List(Control.ListIndex)
'                        cad = mTag.tabla & "." & mTag.Columna & " = '" & cad & "'"
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'
'
'        'CHECK
'        ElseIf TypeOf Control Is CheckBox Then
'            If Control.Tag <> "" Then
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    If Control.Value = 1 Then
'                        cad = mTag.tabla & "." & mTag.Columna & " = 1"
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'        End If
'
'
'    Next Control
'    ObtenerBusqueda = SQL
'Exit Function
'EObtenerBusqueda:
'    ObtenerBusqueda = ""
'    MuestraError Err.Number, "Obtener búsqueda. "
'End Function

Public Function ObtenerBusqueda(ByRef formulario As Form, Optional CHECK As String, Optional vBD As Byte, Optional cadWHERE As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim cad As String
    Dim Sql As String
    Dim tabla As String
    Dim RC As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda = ""
    Sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                If Control.Tag <> "" Then
                    Carga = mTag.Cargar(Control)
                    If Carga Then
                        If Aux = ">>" Then
                            cad = " MAX("
                        Else
                            cad = " MIN("
                        End If
                        'monica
                        Select Case mTag.TipoDato
                            Case "FHF"
                                cad = cad & "date(" & mTag.columna & "))"
                            Case "FHH"
                                cad = cad & "time(" & mTag.columna & "))"
                            Case Else
                                cad = cad & mTag.columna & ")"
                        End Select
                        
                        Sql = "Select " & cad & " from " & mTag.tabla
                        If cadWHERE <> "" Then Sql = Sql & " WHERE " & cadWHERE
                        Sql = ObtenerMaximoMinimo(Sql, vBD)
                        Select Case mTag.TipoDato
                        Case "N"
                            Sql = mTag.tabla & "." & mTag.columna & " = " & TransformaComasPuntos(Sql)
                        Case "F"
                            Sql = mTag.tabla & "." & mTag.columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Case "FHF"
                            Sql = "date(" & mTag.tabla & "." & mTag.columna & ") = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Case "FHH"
                            Sql = "time(" & mTag.tabla & "." & mTag.columna & ") = '" & Format(Sql, "hh:mm:ss") & "'"
                        Case Else
                            '[Monica]04/03/2013: quito las comillas
                            Sql = mTag.tabla & "." & mTag.columna & " = " & DBSet(Sql, "T") ' & "'"
                        End Select
                        Sql = "(" & Sql & ")"
                    End If
                End If
            End If
        End If
    Next


'++monica: lo he añadido del anterior obtenerbusqueda
    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then

                    Sql = mTag.tabla & "." & mTag.columna & " is NULL"
                    Sql = "(" & Sql & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next
 

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                'Cargamos el tag
                Carga = mTag.Cargar(Control)
                If Carga Then
'                    Debug.Print Control.Tag
                    Aux = Trim(Control.Text)
                    If Aux <> "" Then
                        If mTag.tabla <> "" Then
                            tabla = mTag.tabla & "."
                            Else
                            tabla = ""
                        End If
                        RC = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.columna, Aux, cad)
                        If RC = 0 Then
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex > -1 Then
                        If mTag.TipoDato = "N" Then
                            cad = Control.ItemData(Control.ListIndex)
                        Else
                            cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        cad = mTag.tabla & "." & mTag.columna & " = " & cad
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is CheckBox Then
            '=============== Añade: Laura, 15/04/05
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    Aux = ""
                    If CHECK <> "" Then
                        tabla = DBLet(Control.Index, "T")
                        If tabla <> "" Then tabla = "(" & tabla & ")"
                        tabla = Control.Name & tabla & "|"
                        If InStr(1, CHECK, tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    If Aux <> "" Then
'                    If Control.Value = 1 Then
                        cad = Control.Value
                        cad = mTag.tabla & "." & mTag.columna & " = " & cad
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            End If
            '===================
        End If
    Next Control
    ObtenerBusqueda = Sql
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda = ""
    MuestraError Err.Number, "Obtener búsqueda. " & vbCrLf & Err.Description
End Function

'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function ObtenerBusqueda2(ByRef formulario As Form, Optional CHECK As String, Optional opcio As Integer, Optional nom_frame As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim cad As String
    Dim Sql As String
    Dim tabla As String
    Dim RC As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda2 = ""
    Sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Aux = ">>" Then
                            cad = " MAX(" & mTag.columna & ")"
                        Else
                            cad = " MIN(" & mTag.columna & ")"
                        End If
                        Sql = "Select " & cad & " from " & mTag.tabla
                        Sql = ObtenerMaximoMinimo(Sql)
                        Select Case mTag.TipoDato
                        Case "N"
                            Sql = mTag.tabla & "." & mTag.columna & " = " & TransformaComasPuntos(Sql)
                        Case "F"
                            Sql = mTag.tabla & "." & mTag.columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Case Else
                            '[Monica]04/03/2013: quito las comillas y pongo el dbset
                            Sql = mTag.tabla & "." & mTag.columna & " = " & DBSet(Sql, "T") ' & "'"
                        End Select
                        Sql = "(" & Sql & ")"
                    End If
                End If
            End If
        End If
    Next

'++monica: lo he añadido del anterior obtenerbusqueda
    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then

                    Sql = mTag.tabla & "." & mTag.columna & " is NULL"
                    Sql = "(" & Sql & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
          If Control.Tag <> "" Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    Aux = Trim(Control.Text)
                    If Aux <> "" Then
                        If mTag.tabla <> "" Then
                            tabla = mTag.tabla & "."
                            Else
                            tabla = ""
                        End If
                        RC = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.columna, Aux, cad)
                        If RC = 0 Then
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        End If
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then ' +-+- 12/05/05: canvi de Cèsar, no te sentit passar-li un control que no té TAG +-+-
                mTag.Cargar Control
                If mTag.Cargado Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Control.ListIndex > -1 Then
                            cad = Control.ItemData(Control.ListIndex)
                            cad = mTag.tabla & "." & mTag.columna & " = " & cad
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                End If
            End If
            
         ElseIf TypeOf Control Is CheckBox Then
            '=============== Añade: Laura, 27/04/05
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    ' añadido 12022007
                    Aux = ""
                    If CHECK <> "" Then
                        tabla = DBLet(Control.Index, "T")
                        If tabla <> "" Then tabla = "(" & tabla & ")"
                        tabla = Control.Name & tabla & "|"
                        If InStr(1, CHECK, tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    If Aux <> "" Then
'                    If Control.Value = 1 Then
                        cad = Control.Value
                        cad = mTag.tabla & "." & mTag.columna & " = " & cad
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            End If
            '===================
        End If
    Next Control
    ObtenerBusqueda2 = Sql
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda2 = ""
    MuestraError Err.Number, "Obtener búsqueda. " & vbCrLf & Err.Description
End Function

'Añado Optional CHECK As String. Para poder realizar las busquedas con los checks
'monica corresponde al ObtenerBusqueda de laura
Public Function ObtenerBusqueda3(ByRef formulario As Form, paraRPT As Boolean, Optional CHECK As String) As String
Dim Control As Object
Dim Carga As Boolean
Dim mTag As CTag
Dim Aux As String
Dim cad As String
Dim Sql As String
Dim tabla As String, columna As String
Dim RC As Byte

    On Error GoTo EObtenerBusqueda3

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda3 = ""
    Sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.visible Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Aux = ">>" Then
                        If Not paraRPT Then
                            cad = " MAX(" & mTag.columna & ")"
                        Else
                            cad = " MAX({" & mTag.tabla & "." & mTag.columna & "})"
                        End If
                    Else
                        If Not paraRPT Then
                            cad = " MIN(" & mTag.columna & ")"
                        Else
                            cad = " MIN({" & mTag.tabla & "." & mTag.columna & "})"
                        End If
                    End If
                    If Not paraRPT Then
                        Sql = "Select " & cad & " from " & mTag.tabla
                    Else
                        Sql = "Select " & cad & " from {" & mTag.tabla & "}"
                    End If
                    Sql = ObtenerMaximoMinimo(Sql)
                    Select Case mTag.TipoDato
                    Case "N"
                        If Sql <> "" Then
                            If Not paraRPT Then
                                Sql = mTag.tabla & "." & mTag.columna & " = " & TransformaComasPuntos(Sql)
                            Else
                                Sql = "{" & mTag.tabla & "." & mTag.columna & "} = " & TransformaComasPuntos(Sql)
                            End If
                        End If
                    Case "F"
                        If Sql = "" Then Sql = "0000-00-00"
                        If Not paraRPT Then
                            Sql = mTag.tabla & "." & mTag.columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Else
                            Sql = "{" & mTag.tabla & "." & mTag.columna & "} = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        End If
                    Case Else
                        '[Monica]04/03/2013: quito comillas
                        If Not paraRPT Then
                            Sql = mTag.tabla & "." & mTag.columna & " = " & DBSet(Sql, "T") '& "'"
                        Else
                            Sql = "{" & mTag.tabla & "." & mTag.columna & "} = " & DBSet(Sql, "T") ' & "'"
                        End If
                    End Select
                End If
            End If
        End If
    Next

    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.visible Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Not paraRPT Then
                        Sql = mTag.tabla & "." & mTag.columna & " is NULL"
                    Else
                        Sql = "{" & mTag.tabla & "." & mTag.columna & "} is NULL"
                    End If
                    Sql = "(" & Sql & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.visible Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                If mTag.Cargado Then
                    Aux = Trim(Control.Text)
                    Aux = QuitarCaracterEnter(Aux) 'Si es multilinea quitar ENTER
                    If Aux <> "" Then
                        If mTag.tabla <> "" Then
                            If Not paraRPT Then
                                tabla = mTag.tabla & "."
                            Else
                                tabla = "{" & mTag.tabla & "."
                            End If
                        Else
                            tabla = ""
                        End If
                        If Not paraRPT Then
                            columna = mTag.columna
                        Else
                            columna = mTag.columna & "}"
                        End If
                    RC = SeparaCampoBusqueda3(mTag.TipoDato, tabla & columna, Aux, cad, paraRPT)
                    If RC = 0 Then
                        If Sql <> "" Then Sql = Sql & " AND "
                        If Not paraRPT Then
                            Sql = Sql & "(" & cad & ")"
                        Else
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                End If
            End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If

        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.visible Then
            mTag.Cargar Control
            If mTag.Cargado Then
                If Control.ListIndex > -1 Then
                    If mTag.TipoDato <> "T" Then
                        cad = Control.ItemData(Control.ListIndex)
                        If Not paraRPT Then
                            cad = mTag.tabla & "." & mTag.columna & " = " & cad
                        Else
                            cad = "{" & mTag.tabla & "." & mTag.columna & "} = " & cad
                        End If
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    Else
                        cad = Control.List(Control.ListIndex)
                        If Not paraRPT Then
                            cad = mTag.tabla & "." & mTag.columna & " = '" & cad & "'"
                        Else
                            cad = "{" & mTag.tabla & "." & mTag.columna & "} = '" & cad & "'"
                        End If
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            End If


        'CHECK
                'CHECK
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    
                    Aux = ""
                    If CHECK <> "" Then
                        CheckBusqueda Control
                        tabla = NombreCheck & "|"
                        If InStr(1, CHECK, tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    
                    If Aux <> "" Then
                        If Not paraRPT Then
                            cad = mTag.tabla & "." & mTag.columna
                        Else
                            cad = "{" & mTag.tabla & "." & mTag.columna & "} "
                        End If
                        
                        cad = cad & " = " & Aux
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If 'cargado
                End If '<>""
            End If
        End If
    
    Next Control
    ObtenerBusqueda3 = Sql
Exit Function
EObtenerBusqueda3:
    ObtenerBusqueda3 = ""
    MuestraError Err.Number, "Obtener búsqueda. "
End Function

Public Function ModificaDesdeFormulario(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        'Sea para el where o para el update esto lo necesito
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                             cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"

                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox And Control.visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        'Lo pondremos para el WHERE
                         If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                         cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                    Else
                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
'
'
'                   If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                   'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
'                   cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    conn.Execute Aux, , adCmdText

    ModificaDesdeFormulario = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function


Public Function ModificaDesdeFormulario2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario2 = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            'Sea para el where o para el update esto lo necesito
                            Aux = ValorParaSQL(Control.Text, mTag)
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                 If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                                 cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
    
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Control.Value = 1 Then
                        Aux = "TRUE"
                    Else
                        Aux = "FALSE"
                    End If
                    If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'Esta es para access
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If

        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.ListIndex = -1 Then
                            Aux = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            Aux = Control.ItemData(Control.ListIndex)
                        Else
                            Aux = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                             cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                        End If
'
'
'                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                        'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
'                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            Aux = Control.Index
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                              If mTag.EsClave Then
                                  'Lo pondremos para el WHERE
                                   If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                                   cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                              Else
                                  If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                  cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                              End If
                        End If
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is DTPicker Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
'                        If Control.Value Then
                         If mTag.columna <> "" Then
'                            Aux = Control.index
                            If Control.visible Then
                                Aux = ValorParaSQL(Control.Value, mTag)
                            Else
                                Aux = ValorNulo
                            End If
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                                cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    conn.Execute Aux, , adCmdText

    ' ### [Monica] 18/12/2006
    CadenaCambio = cadUPDATE

    ModificaDesdeFormulario2 = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar 2. " & Err.Description
End Function

Public Function ModificaDesdeFormulario1(ByRef formulario As Form, Opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario1
    ModificaDesdeFormulario1 = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is CommonDialog Then
        ElseIf TypeOf Control Is TextBox And Control.visible = True Then
            If (Opcion = 1 And Control.Name = "Text1") Or (Opcion = 3 And Control.Name = "txtAux") Then
            If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            'Sea para el where o para el update esto lo necesito
                            Aux = ValorParaSQL(Control.Text, mTag)
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                 If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                                 cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
    
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.visible Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox And Control.visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        ElseIf TypeOf Control Is OptionButton And Control.visible Then
            If Control.Enabled Then
                If Control.Value = True And Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        Aux = Control.Index
                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    conn.Execute Aux, , adCmdText

    ModificaDesdeFormulario1 = True
    Exit Function
    
EModificaDesdeFormulario1:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function

Public Sub FormateaCampo(vTex As TextBox)
    Dim mTag As CTag
    Dim cad As String
    On Error GoTo EFormateaCampo
    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        If vTex.Text <> "" Then
            If mTag.Formato <> "" Then
                cad = TransformaPuntosComas(vTex.Text)
                cad = Format(cad, mTag.Formato)
                vTex.Text = cad
            End If
        End If
    End If
EFormateaCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Sub


Public Function FormatoCampo(ByRef vTex As TextBox) As String
'Devuelve el formato del campo en el TAg: "0000"
Dim mTag As CTag
Dim cad As String
    
    On Error GoTo EFormatoCampo

    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        FormatoCampo = mTag.Formato
    End If
    
EFormatoCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


'Añade: CESAR
'Para utilizalo en el arreglaGrid
Public Function FormatoCampo2(ByRef objec As Object) As String
'Devuelve el formato del campo en el TAg: "0000"
Dim mTag As CTag
Dim cad As String

    On Error GoTo EFormatoCampo2

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        FormatoCampo2 = mTag.Formato
    End If
    
EFormatoCampo2:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


Public Function TipoCamp(ByRef objec As Object) As String
Dim mTag As CTag
Dim cad As String

    On Error GoTo ETipoCamp

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        TipoCamp = mTag.TipoDato
    End If

ETipoCamp:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef cadena As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim cont As Integer
Dim cad As String

    I = 0
    cont = 1
    cad = ""
    Do
        J = I + 1
        I = InStr(J, cadena, "|")
        If I > 0 Then
            If cont = Orden Then
                cad = Mid(cadena, J, I - J)
                I = Len(cadena) 'Para salir del bucle
                Else
                    cont = cont + 1
            End If
        End If
    Loop Until I = 0
    RecuperaValor = cad
End Function


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValorNew(ByRef cadena As String, Separador As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim cont As Integer
Dim cad As String

    I = 0
    cont = 1
    cad = ""
    Do
        J = I + 1
        I = InStr(J, cadena, Separador)
        If I > 0 Then
            If cont = Orden Then
                cad = Mid(cadena, J, I - J)
                I = Len(cadena) 'Para salir del bucle
                Else
                    cont = cont + 1
            End If
        End If
    Loop Until I = 0
    RecuperaValorNew = cad
End Function




'-----------------------------------------------------------------------
'Deshabilitar ciertas opciones del menu
'EN funcion del nivel de usuario
'Esto es a nivel general, cuando el Toolba es el mismo

'Para ello en el tag del button tendremos k poner un numero k nos diara hasta k nivel esta permitido

Public Sub PonerOpcionesMenuGeneral(ByRef formulario As Form)
Dim I As Integer
Dim J As Integer
'Dim bol As Boolean

On Error GoTo EPonerOpcionesMenuGeneral
'bol = vSesion.Nivel < 2

'Añadir, modificar y borrar deshabilitados si no nivel
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


Public Sub PonerModoMenuGral(ByRef formulario As Form, activo As Boolean)
Dim I As Integer
'Dim j As Integer

On Error GoTo PonerModoMenuGral

'Añadir, modificar y borrar deshabilitados si no Modo
    With formulario
        For I = 1 To .Toolbar1.Buttons.Count
            Select Case .Toolbar1.Buttons(I).ToolTipText
                Case "Nuevo"
                    .Toolbar1.Buttons(I).visible = Not .DeConsulta
                Case "Modificar", "Eliminar", "Imprimir"
                    .Toolbar1.Buttons(I).visible = Not .DeConsulta
                    .Toolbar1.Buttons(I).Enabled = activo
'                Case "Modificar"
'                Case "Eliminar"
'                Case "Imprimir"
            End Select
        Next I
        
        
        'El menu Visible
        .mnModificar.visible = Not .DeConsulta
        .mnEliminar.visible = Not .DeConsulta
        'El menu activo
        .mnModificar.Enabled = activo
        .mnEliminar.Enabled = activo
    End With
    
    
Exit Sub
PonerModoMenuGral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub

Public Sub PonerOpcionesMenuGeneralNew(formulario As Form)
Dim Control As Object
Dim I As Integer
Dim J As Integer
'Dim bol As Boolean

On Error GoTo EPonerOpcionesMenuGeneralNew
'bol = vSesion.Nivel < 2
'Añadir, modificar y borrar deshabilitados si no nivel
    For Each Control In formulario.Controls
'        Debug.Print Control.Name
        
        If Mid(Control.Name, 1, 2) = "mn" And Mid(Control.Name, 1, 7) <> "mnBarra" _
           And Control.Name <> "mnOpciones" Then
            J = Val(Control.HelpContextID)
            If J < vUsu.Nivel And J <> 0 Then
                Control.Enabled = False
            End If
        End If
    Next Control

Exit Sub
EPonerOpcionesMenuGeneralNew:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub



'Este modifica las claves prinipales y todo
'la sentenca del WHERE cod=1 and .. viene en claves
Public Function ModificaDesdeFormularioClaves(ByRef formulario As Form, Claves As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUPDATE As String
Dim I As Integer

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormularioClaves = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    cadWHERE = Claves
    'Construimos el SQL
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    conn.Execute Aux, , adCmdText

ModificaDesdeFormularioClaves = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function

Public Function BLOQUEADesdeFormulario(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQUEADesdeFormulario
    BLOQUEADesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.visible = True Then
                If Control.Tag <> "" Then
    
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Sea para el where o para el update esto lo necesito
                        
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                             cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                        End If
                    End If
                End If
            End If
        End If
    Next Control

    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "select * FROM " & mTag.tabla
        Aux = Aux & " WHERE " & cadWHERE & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario = True
    End If
    
EBLOQUEADesdeFormulario:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function


'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function BLOQUEADesdeFormulario2(ByRef formulario As Form, ByRef ado As Adodc, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim AntiguoCursor As Byte
Dim nomcamp As String

    On Error GoTo EBLOQUEADesdeFormulario2
    
    BLOQUEADesdeFormulario2 = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If (TypeOf Control Is TextBox) Or (TypeOf Control Is ComboBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Sea para el where o para el update esto lo necesito
                        'Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            Aux = ValorParaSQL(CStr(ado.Recordset.Fields(mTag.columna)), mTag)
                            'Lo pondremos para el WHERE
                             If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                             cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                        End If
                    End If
                End If
            End If
        End If
    Next Control

    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "select * FROM " & mTag.tabla
        Aux = Aux & " WHERE " & cadWHERE & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario2 = True
    End If
    
EBLOQUEADesdeFormulario2:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla 2"
'        BLOQUEADesdeFormulario2 = False
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function BloqueaRegistro(cadTabla As String, cadWHERE As String) As Boolean
Dim Aux As String

    On Error GoTo EBloqueaRegistro
        
    BloqueaRegistro = False
    Aux = "select * FROM " & cadTabla
    If cadWHERE <> "" Then Aux = Aux & " WHERE " & cadWHERE & " FOR UPDATE"

    'Intenteamos bloquear
    PreparaBloquear
    conn.Execute Aux, , adCmdText
    BloqueaRegistro = True

EBloqueaRegistro:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
End Function


Public Function BloqueaRegistroForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim AuxDef As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQ
    BloqueaRegistroForm = False
    Set mTag = New CTag
    Aux = ""
    AuxDef = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        Aux = ValorParaSQL(Control.Text, mTag)
                        AuxDef = AuxDef & Aux & "|"
                    End If
                End If
            End If
        End If
    Next Control

    If AuxDef = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
'        Aux = "Insert into zBloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & mTag.tabla
        Aux = Aux & "',""" & AuxDef & """)"
        conn.Execute Aux
        BloqueaRegistroForm = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If conn.Errors.Count > 0 Then
            If conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueaRegistroForm(ByRef TextBoxConTag As TextBox) As Boolean
Dim mTag As CTag
Dim Sql As String

'Solo me interesa la tabla
On Error Resume Next
    Set mTag = New CTag
    mTag.Cargar TextBoxConTag
    If mTag.Cargado Then
'        SQL = "DELETE from zBloqueos where codusu=" & vUsu.Codigo & " and tabla='" & mTag.tabla & "'"
        conn.Execute Sql
        If Err.Number <> 0 Then
            Err.Clear
        End If
    End If
    Set mTag = Nothing
End Function


'====================== LAURA

Public Function ComprobarCero(Valor As String) As String
    If Valor = "" Then
        ComprobarCero = "0"
    Else
        ComprobarCero = Valor
    End If
End Function

'Public Sub InsertarCambios(Tabla As String, ValorAnterior As String, Numalbar As String)
'Dim SQL As String
'Dim Sql2 As String
'
'    SQL = CadenaCambio
'
'    Sql2 = "insert into cambios (codusu, fechacambio, tabla, numalbar, cadena, valoranterior) values ("
'    Sql2 = Sql2 & DBSet(vSesion.CodUsu, "N") & "," & DBSet(Now, "FH") & "," & DBSet(Tabla, "T") & ","
'    Sql2 = Sql2 & DBSet(Numalbar, "T") & ","
'    Sql2 = Sql2 & DBSet(SQL, "T") & ","
'    If ValorAnterior = ValorNulo Then
'        Sql2 = Sql2 & ValorNulo & ")"
'    Else
'        Sql2 = Sql2 & DBSet(ValorAnterior, "T") & ")"
'    End If
'
'    conn.Execute Sql2
'
'End Sub
    
Public Sub CargarValoresAnteriores(formulario As Form, Optional opcio As Integer, Optional nom_frame As String)
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim cad As String
    Set mTag = New CTag

    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" And Not mTag.EsClave Then
                            If Izda <> "" Then Izda = Izda & " , "
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.columna & " = "
                            'Parte VALUES
                            cad = ValorParaSQL(Control.Text, mTag)
                            Izda = Izda & cad
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Not mTag.EsClave Then
                        If Izda <> "" Then Izda = Izda & " , "
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & " = "
                        If Control.Value = 1 Then
                            cad = "1"
                            Else
                            cad = "0"
                        End If
                        If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                        Izda = Izda & cad
                    End If
                End If
            End If
            
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado And Not mTag.EsClave Then
                        If Izda <> "" Then Izda = Izda & " , "
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & " = "
                        If Control.ListIndex = -1 Then
                            cad = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            cad = Control.ItemData(Control.ListIndex)
                        Else
                            cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        Izda = Izda & cad
                    End If
                End If
            End If
            
        'OPTION BUTTON
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado And Not mTag.EsClave Then
                        If Control.Value Then
                            If Izda <> "" Then Izda = Izda & " , "
                            Izda = Izda & "" & mTag.columna & " = "
                            cad = Control.Index
                            Izda = Izda & cad
                        End If
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is DTPicker Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado And Not mTag.EsClave Then
                        If Izda <> "" Then Izda = Izda & " , "
                        Izda = Izda & "" & mTag.columna & " = "
                        
                        'Parte VALUES
                        If Control.visible Then
                            cad = ValorParaSQL(Control.Value, mTag)
                        Else
                            cad = ValorNulo
                        End If
                        Izda = Izda & cad
                    End If
                End If
            End If
        End If
        
    Next Control

    ValorAnterior = Izda

End Sub


Public Sub CalcularImporteNue(ByRef cantidad As TextBox, ByRef Precio As TextBox, ByRef Importe As TextBox, Tipo As Integer)
'Calcula el Importe de una linea de hcode facturas
Dim vImp As Currency
Dim vCan As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    cantidad = ComprobarCero(cantidad.Text)
    Precio = ComprobarCero(Precio.Text)
    Importe = ComprobarCero(Importe.Text)
    
    Select Case Tipo
        Case 0 ' me han introducido la cantidad
            vImp = CCur(ImporteFormateado(cantidad.Text)) * CCur(ImporteFormateado(Precio.Text))
            vImp = Round2(vImp, 2)
            Importe.Text = Format(vImp, "###,##0.00")
        Case 1 ' me han introducido el precio
            vImp = CCur(ImporteFormateado(cantidad.Text)) * CCur(ImporteFormateado(Precio.Text))
            vImp = Round2(vImp, 2)
            Importe.Text = Format(vImp, "###,##0.00")
        Case 2 ' me han introducido el importe
            vCan = CCur(ImporteFormateado(Importe.Text)) / CCur(ImporteFormateado(Precio.Text))
            vCan = Round2(vCan, 3)
            cantidad.Text = Format(vCan, "##,##0.000")
    End Select
    
End Sub


Public Function ExisteCP(T As TextBox) As Boolean
'comprueba para un campo de texto que sea clave primaria, si ya existe un
'registro con ese valor
Dim vtag As CTag
Dim devuelve As String

    On Error GoTo ErrExiste

    ExisteCP = False
    If T.Text <> "" Then
        If T.Tag <> "" Then
            Set vtag = New CTag
            If vtag.Cargar(T) Then
'                If vtag.EsClave Then
                    devuelve = DevuelveDesdeBDNew(cAgro, vtag.tabla, vtag.columna, vtag.columna, T.Text, vtag.TipoDato)
                    If devuelve <> "" Then
    '                    MsgBox "Ya existe un registro para " & vtag.Nombre & ": " & T.Text, vbExclamation
                        MsgBox "Ya existe el " & vtag.Nombre & ": " & T.Text, vbExclamation
                        ExisteCP = True
                        PonerFoco T
                    End If
'                End If
            End If
            Set vtag = Nothing
        End If
    End If
    Exit Function
    
ErrExiste:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar código.", Err.Description
End Function




Public Function TotalRegistros(vSql As String) As Long
'Devuelve el valor de la SQL
'para obtener COUNT(*) de la tabla
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Set Rs = New ADODB.Recordset
    Rs.Open vSql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalRegistros = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then TotalRegistros = Rs.Fields(0).Value  'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        TotalRegistros = 0
        Err.Clear
    End If
End Function



Public Function Round2(Number As Variant, Optional NumDigitsAfterDecimals As Long) As Variant
Dim Ent As Integer
Dim cad As String

  ' Comprobaciones

  If Not IsNumeric(Number) Then
    Err.Raise 13, "Round2", "Error de tipo. Ha de ser un número."
    Exit Function
  End If

  If NumDigitsAfterDecimals < 0 Then
    Err.Raise 0, "Round2", "NumDigitsAfterDecimals no puede ser negativo."
    Exit Function
  End If

  ' Redondeo.

  cad = "0"
  If NumDigitsAfterDecimals <> 0 Then cad = cad & "." & String(NumDigitsAfterDecimals, "0")
  Round2 = Val(TransformaComasPuntos(Format(Number, cad)))

End Function



Public Function ObtenerLetraSerie(tipMov As String) As String
'Devuelve la letra de serie asociada al tipo de movimiento
Dim LEtra As String

    On Error Resume Next
    
    LEtra = DevuelveDesdeBDNew(cAgro, "stipom", "letraser", "codtipom", tipMov, "T")
    If LEtra = "" Then MsgBox "Las factura de venta no tienen asignada una letra de serie", vbInformation
    ObtenerLetraSerie = LEtra
End Function


Public Function ObtenerLetraSerie2(tipMov As String) As String
'Devuelve la letra de serie asociada al tipo de movimiento
Dim LEtra As String

    On Error Resume Next
    
    LEtra = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "letraser", "codtipom", tipMov, "T")
    If LEtra = "" Then MsgBox "Las factura de venta no tienen asignada una letra de serie", vbInformation
    ObtenerLetraSerie2 = LEtra
End Function



Public Function CalcularPorcentaje(Importe As Currency, Porce As Currency, NumDecimales As Long) As Variant
'devuelve el valor del Porcentaje aplicado al Importe
'Ej el 16% de 120 = 19.2
'Dim vImp As Currency
'Dim vDto As Currency
    
    On Error Resume Next
'
'    Importe = ComprobarCero(Importe)
'    Dto = ComprobarCero(Dto)
'
'    vImp = CCur(Importe)
'    vDto = CCur(Dto)
    
    
    'vImp = Round(vImp, 2)
    
    CalcularPorcentaje = Round2((Importe * Porce) / 100, NumDecimales)
    
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function CalcularDto(Importe As String, Dto As String) As String
'devuelve el Dto% del Importe
'Ej el 16% de 120 = 19.2
Dim vImp As Currency
Dim vDto As Currency
On Error Resume Next

    Importe = ComprobarCero(Importe)
    Dto = ComprobarCero(Dto)
    
    vImp = CCur(Importe)
    vDto = CCur(Dto)
    
    vImp = ((vImp * vDto) / 100)
    'vImp = Round(vImp, 2)
    
    CalcularDto = CStr(vImp)
    If Err.Number <> 0 Then Err.Clear
End Function

'---------------------------------------------------------------------------------
'
'       Para buscar en los checks con las dos opciones de true y false
'
'A partir de un check cualquiera devolvera nombre e indice, si tiene. Si no sera ()
Public Sub CheckBusqueda(ByRef CH As CheckBox)
    NombreCheck = ""
    NombreCheck = CH.Name & "("
    On Error Resume Next
    NombreCheck = NombreCheck & CH.Index
    If Err.Number <> 0 Then Err.Clear
    NombreCheck = NombreCheck & ")"
End Sub

Public Sub CheckCadenaBusqueda(ByRef CH As CheckBox, ByRef CadenaCHECKs As String)
        CheckBusqueda CH
        If InStr(1, CadenaCHECKs, NombreCheck) = 0 Then CadenaCHECKs = CadenaCHECKs & NombreCheck & "|"
End Sub

Public Function PonerAlmacen(codAlm As String) As String
'Comprueba si existe el Almacen y lo pone en el Text
Dim devuelve As String
    
    On Error Resume Next

    If codAlm = "" Then
        MsgBox "Debe introducir el Almacen.", vbInformation
    Else
        devuelve = DevuelveDesdeBDNew(cAgro, "salmpr", "codalmac", "codalmac", codAlm, "N")
        If devuelve = "" Then
            MsgBox "No existe el Almacen: " & Format(codAlm, "000"), vbInformation
            PonerAlmacen = ""
        Else
            PonerAlmacen = Format(codAlm, "000")
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function CalcularImporte(cantidad As String, Precio As String, Dto1 As String, Dto2 As String, TipoDto As Byte, ImpDto As String, Optional Bruto As String) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
'Si DtoProv=sprove.tipodtos, calcular Importe para Proveedores y obtener el tipo de descuento
'del campo sprove.tipodtos, si es para Clientes obtener el tipo de descuento del
'parametro spara1.tipodtos
'Tipo Descuento: 0=aditivo, 1=sobre resto
Dim vImp As Currency
Dim vDto1 As Currency, vDto2 As Currency
Dim vPre As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    cantidad = ComprobarCero(cantidad)
    vPre = ComprobarCero(Precio)
    Dto1 = ComprobarCero(Dto1)
    Dto2 = ComprobarCero(Dto2)
    
    If Bruto <> "" Then
        vImp = CCur(Bruto) - CCur(ImpDto)
    Else
        vImp = (CCur(cantidad) * CCur(vPre)) - CCur(ImpDto)
    End If
        
    If TipoDto = 0 Then 'Dto Aditivo
        vDto1 = (CCur(Dto1) * vImp) / 100
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto1 - vDto2
    ElseIf TipoDto = 1 Then 'Sobre Resto
        vDto1 = (CCur(Dto1) * vImp) / 100
        vImp = vImp - vDto1
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto2
    End If
    vImp = Round(vImp, 2)
    CalcularImporte = CStr(vImp)
End Function

Public Function EsProveedorVarios(codProve As String) As Boolean
Dim devuelve As String

    EsProveedorVarios = False
    devuelve = DevuelveDesdeBD("provario", "proveedor", "codprove", codProve, "N")
    If devuelve <> "" Then EsProveedorVarios = CBool(devuelve)
    'Es proveedor de varios Y podemos recuperar de ????
End Function

Public Function QuitarCero(Valor As String) As String
    On Error Resume Next
    
    If Valor <> "" Then
        If CSng(Valor) = 0 Then
            QuitarCero = ""
        Else
            QuitarCero = Valor
        End If
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function BloqueoManual(cadTabla As String, cadWHERE As String)
Dim Aux As String

On Error GoTo EBLOQ

    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "INSERT INTO zbloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & cadTabla
        Aux = Aux & "',""" & cadWHERE & """)"
        conn.Execute Aux
        BloqueoManual = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If conn.Errors.Count > 0 Then
            If conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
'    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueoManual(cadTabla As String) As Boolean
Dim Sql As String

'Solo me interesa la tabla
On Error Resume Next

        Sql = "DELETE FROM zbloqueos WHERE codusu=" & vUsu.Codigo & " and tabla='" & cadTabla & "'"
        conn.Execute Sql
        If Err.Number <> 0 Then
            Err.Clear
        End If
End Function

'++monica
'funcion de la libreria general de gessocial de Rafa, necesaria para pasar al aridoc
Public Function CApos(Texto As String) As String
    Dim I As Integer
    I = InStr(1, Texto, "'")
    If I = 0 Then
        CApos = Texto
    Else
        CApos = Mid(Texto, 1, I) & "'" & Mid(Texto, I + 1, Len(Texto) - I)
    End If
    '-- Ya que estamos transformamos las Ñ
    Texto = CApos
    I = InStr(1, Texto, "¥")
    If I = 0 Then
        CApos = Texto
    Else
        CApos = Mid(Texto, 1, I - 1) & "Ñ" & Mid(Texto, I + 1, Len(Texto) - I)
    End If
    '-- Y otra más
    Texto = CApos
    I = InStr(1, Texto, "¾")
    If I = 0 Then
        CApos = Texto
    Else
        CApos = Mid(Texto, 1, I - 1) & "Ñ" & Mid(Texto, I + 1, Len(Texto) - I)
    End If
    '-- Seguimos con transformaciones
    Texto = CApos
    I = InStr(1, Texto, "¦")
    If I = 0 Then
        CApos = Texto
    Else
        CApos = Mid(Texto, 1, I - 1) & "ª" & Mid(Texto, I + 1, Len(Texto) - I)
    End If
End Function



Public Function DevuelveValor(vSql As String) As Variant
'Devuelve el valor de la SQL
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Set Rs = New ADODB.Recordset
    Rs.Open vSql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
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

Public Function TotalRegistrosConsulta(cadSQL) As Long
Dim cad As String
Dim Rs As ADODB.Recordset

    On Error GoTo ErrTotReg
    cad = "SELECT count(*) FROM (" & cadSQL & ") x"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not Rs.EOF Then
        TotalRegistrosConsulta = DBLet(Rs.Fields(0).Value, "N")
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
ErrTotReg:
    MuestraError Err.Number, "", Err.Description
End Function


Public Sub BorrarArchivo(nomFich As String)
    
    On Error Resume Next
    
    If Dir(nomFich) <> "" Then Kill nomFich

    If Err.Number <> 0 Then MuestraError Err.Number, "Borrar fichero"

End Sub

'++monica: facturamos segun el campo de forfaits
Public Function TipoFacturarForfaits(Albaran As String, Linea As String) As Byte
' devuelve 0: facturar por unidades
'          1: facturar por kilos
Dim Rs As ADODB.Recordset
Dim Sql As String

    TipoFacturarForfaits = 2
    
    If Trim(Albaran) = "" Or Trim(Linea) = "" Then Exit Function

    Sql = "select forfaits.facturar from albaran_variedad, forfaits "
    Sql = Sql & " where albaran_variedad.numalbar = " & DBSet(Albaran, "N")
    Sql = Sql & " and albaran_variedad.numlinea = " & DBSet(Linea, "N")
    Sql = Sql & " and forfaits.codforfait = albaran_variedad.codforfait "
    Sql = Sql & " order by numlinea "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If Not Rs.EOF Then
        TipoFacturarForfaits = DBLet(Rs.Fields(0).Value, "N")
    End If
    
End Function


Public Function CalidadDestrio(Variedad As String) As String
Dim Rs As ADODB.Recordset
Dim Sql As String

    CalidadDestrio = ""
    
    If Trim(Variedad) = "" Then Exit Function

    Sql = "select codcalid from rcalidad "
    Sql = Sql & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    Sql = Sql & " and rcalidad.tipcalid = 1" ' tipo de calidad de destrio
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        CalidadDestrio = DBLet(Rs.Fields(0).Value, "N")
    End If
    
    Set Rs = Nothing
    
End Function


Public Function CalidadDestrioenClasificacion(Variedad As String, Nota As String, Optional ConKilos As Boolean) As String
'conkilos = true --> miramos que el registro de esa clasificacion tenga kilos <> 0
Dim Rs As ADODB.Recordset
Dim Sql As String

    CalidadDestrioenClasificacion = ""
    
    If Trim(Variedad) = "" Then Exit Function

    Sql = "select rcalidad.codcalid from rclasifica_clasif, rcalidad "
    Sql = Sql & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    Sql = Sql & " and rcalidad.tipcalid = 1" ' tipo de calidad de destrio
    Sql = Sql & " and rclasifica_clasif.numnotac = " & DBSet(Nota, "N")
    Sql = Sql & " and rclasifica_clasif.codvarie = rcalidad.codvarie "
    Sql = Sql & " and rclasifica_clasif.codcalid = rcalidad.codcalid "
    
    If ConKilos Then
        Sql = Sql & " and rclasifica_clasif.kilosnet <> 0"
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        CalidadDestrioenClasificacion = DBLet(Rs.Fields(0).Value, "N")
    End If
    
    Set Rs = Nothing
    
End Function

Public Function CalidadMaximaMuestraenClasificacion(Variedad As String, Nota As String, Optional ConKilos As Boolean) As String
'conkilos = true --> miramos que el registro de esa clasificacion tenga kilos <> 0
Dim Rs As ADODB.Recordset
Dim Sql As String

    CalidadMaximaMuestraenClasificacion = ""
    
    If Trim(Variedad) = "" Then Exit Function

    Sql = "select rclasifica_clasif.codcalid from rclasifica_clasif "
    Sql = Sql & " where rclasifica_clasif.numnotac = " & DBSet(Nota, "N")
    
    If ConKilos Then
        Sql = Sql & " and rclasifica_clasif.kilosnet <> 0"
    End If
    
    Sql = Sql & " and muestra = (select max(rclasifica_clasif.muestra) from rclasifica_clasif "
    Sql = Sql & " where rclasifica_clasif.numnotac = " & DBSet(Nota, "N") & ")"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        CalidadMaximaMuestraenClasificacion = DBLet(Rs.Fields(0).Value, "N")
    End If
    
    Set Rs = Nothing
    
End Function




Public Function HorasDecimal(cantidad As String) As Currency
Dim Entero As Long
Dim vCantidad As String
Dim vDecimal As String
Dim vEntero As String
Dim vHoras As Currency
Dim J As Integer
    HorasDecimal = 0
    
    vCantidad = ImporteSinFormato(cantidad)
    
    J = InStr(1, vCantidad, ",")
    
    If J > 0 Then
        vEntero = Mid(vCantidad, 1, J - 1)
        vDecimal = Mid(vCantidad, J + 1, Len(vCantidad))
        If Len(vDecimal) = 1 Then vDecimal = vDecimal & "0"
    Else
        vEntero = vCantidad
        vDecimal = "0"
    End If
    
    
    vHoras = (CLng(vEntero) * 60) + CLng(vDecimal)

    HorasDecimal = Round2(vHoras / 60, 2)
    
End Function


Public Function DecimalHoras(cantidad As Currency) As Currency
Dim Entero As Long
Dim vCantidad As String
Dim vDecimal As String
Dim vEntero As String
Dim vHoras As Currency
Dim J As Integer
    
    DecimalHoras = 0
    
    vCantidad = ImporteSinFormato(CStr(cantidad))
    
    J = InStr(1, vCantidad, ",")
    
    If J > 0 Then
        vEntero = Mid(vCantidad, 1, J - 1)
        vDecimal = Mid(vCantidad, J + 1, Len(vCantidad))
    Else
        vEntero = vCantidad
        vDecimal = "0"
    End If
    
    
    DecimalHoras = CInt(vEntero) + Round2((CInt(vDecimal) * 60 / 100) / 100, 2)
    
End Function



Public Function Horas(cantidad As String) As Currency
Dim Entero As Long
Dim vCantidad As String
Dim vDecimal As String
Dim vEntero As String
Dim vHoras As Currency
Dim J As Integer

    Horas = 0
    
    vCantidad = ImporteSinFormato(cantidad)
    
    J = InStr(1, vCantidad, ",")
    
    If J > 0 Then
        vEntero = Mid(vCantidad, 1, J - 1)
        vDecimal = Mid(vCantidad, J + 1, Len(vCantidad))
    Else
        vEntero = vCantidad
        vDecimal = 0
    End If
    
    If CLng(vDecimal) >= 60 Then
        vEntero = vEntero + 1
        vDecimal = Format(CInt(vDecimal) - 60, "#0")
    End If
    
    Horas = vEntero & "," & vDecimal
    
End Function




Public Function ComprobacionRangoFechas(Varie As String, Tipo As String, Contador As String, fecha1 As String, fecha2 As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim B As Boolean

    On Error GoTo eComprobacionRangoFechas
    
    ComprobacionRangoFechas = False
    
    Sql = "select rprecios.fechaini, rprecios.fechafin, max(contador) from rprecios "
    Sql = Sql & " where codvarie = " & DBSet(Varie, "N")
    Sql = Sql & " and tipofact = " & DBSet(Tipo, "N")
    
    If Contador <> "" Then
        Sql = Sql & " and contador <> " & DBSet(Contador, "N")
    End If
    Sql = Sql & " group by 1,2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    B = False
    While Not Rs.EOF And Not B
        '[Monica]20/01/2014: añadido el if de que si coinciden no hacer nada
        If fecha1 = DBLet(Rs.Fields(0).Value, "F") And fecha2 = DBLet(Rs.Fields(1).Value, "F") Then
            ComprobacionRangoFechas = True
            Exit Function
        Else
            B = EntreFechas(fecha1, DBLet(Rs.Fields(0).Value, "F"), fecha2)
            If Not B Then B = EntreFechas(fecha1, DBLet(Rs.Fields(1).Value, "F"), fecha2)
            Rs.MoveNext
        End If
    Wend
    
    Rs.Close
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF And Not B
        '[Monica]20/01/2014: añadido el if de que si coinciden no hacer nada
        If fecha1 = DBLet(Rs.Fields(0).Value, "F") And fecha2 = DBLet(Rs.Fields(1).Value, "F") Then
            ComprobacionRangoFechas = True
            Exit Function
        Else
            B = EntreFechas(DBLet(Rs.Fields(0).Value, "F"), fecha1, DBLet(Rs.Fields(1).Value, "F"))
            If Not B Then B = EntreFechas(DBLet(Rs.Fields(0).Value, "F"), fecha2, DBLet(Rs.Fields(1).Value, "F"))
        End If
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    ComprobacionRangoFechas = Not B
    
    Exit Function
    
eComprobacionRangoFechas:
    MuestraError Err.Number, "Comprobacion de rango fechas", Err.Description
End Function

Public Function PartidaCampo(codcampo As String) As String
Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error Resume Next

    PartidaCampo = ""
    
    Sql = "select nomparti from rpartida, rcampos where rcampos.codcampo = " & DBSet(codcampo, "N")
    Sql = Sql & " and rcampos.codparti = rpartida.codparti "
    
    Set Rs = New ADODB.Recordset
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        PartidaCampo = DBLet(Rs.Fields(0).Value, "T")
    End If
    
    Set Rs = Nothing
    
End Function

Public Function HayRegistros(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "Select * FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function

Public Function RellenaAceros(cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim cad As String
    
    cad = Mid("00000000000000000000", 1, longitud)
    If PorLaDerecha Then
        cad = cadena & cad
        RellenaAceros = Left(cad, longitud)
    Else
        cad = cad & cadena
        RellenaAceros = Right(cad, longitud)
    End If
    
End Function

Public Function RellenaABlancos(cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim cad As String
    
    cad = Space(longitud)
    If PorLaDerecha Then
        cad = cadena & cad
        RellenaABlancos = Left(cad, longitud)
    Else
        cad = cad & cadena
        RellenaABlancos = Right(cad, longitud)
    End If
    
End Function


Public Function EstaSocioDeAlta(Socio As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from rsocios where codsocio = " & DBSet(Socio, "N")
    Sql = Sql & " and fechabaja is null"
    
    EstaSocioDeAlta = (TotalRegistros(Sql) > 0)

End Function

Public Function EstaCampoDeAlta(campo As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from rcampos where codcampo = " & DBSet(campo, "N")
    Sql = Sql & " and fecbajas is null"
    
    EstaCampoDeAlta = (TotalRegistros(Sql) > 0)

End Function

Public Function EstaSocioDeAltaSeccion(Socio As String, Secc As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from rsocios_seccion where codsocio = " & DBSet(Socio, "N")
    Sql = Sql & " and codsecci = " & DBSet(Secc, "N")
    Sql = Sql & " and fecbaja is null"
    
    EstaSocioDeAltaSeccion = (TotalRegistros(Sql) > 0)

End Function

Public Function EsSocioDeSeccion(Socio As String, Secc As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from rsocios_seccion where codsocio = " & DBSet(Socio, "N")
    Sql = Sql & " and codsecci = " & DBSet(Secc, "N")
    
    EsSocioDeSeccion = (TotalRegistros(Sql) > 0)

End Function


Public Function CalidadVentaCampo(Variedad As String) As String
Dim Rs As ADODB.Recordset
Dim Sql As String

    CalidadVentaCampo = ""
    
    If Trim(Variedad) = "" Then Exit Function

    Sql = "select codcalid from rcalidad "
    Sql = Sql & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    Sql = Sql & " and rcalidad.tipcalid = 2" ' tipo de calidad de venta campo
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        CalidadVentaCampo = DBLet(Rs.Fields(0).Value, "N")
    End If
    
    Set Rs = Nothing
    
End Function

Public Function CalidadRetirada(Variedad As String) As String
Dim Rs As ADODB.Recordset
Dim Sql As String

    CalidadRetirada = ""
    
    If Trim(Variedad) = "" Then Exit Function

    Sql = "select codcalid from rcalidad "
    Sql = Sql & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    Sql = Sql & " and rcalidad.tipcalid1 = 2" ' tipo de calidad de retirada
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        CalidadRetirada = DBLet(Rs.Fields(0).Value, "N")
    End If
    
    Set Rs = Nothing
    
End Function


Public Function CalidadPrimera(Variedad As String) As String
Dim Rs As ADODB.Recordset
Dim Sql As String

    CalidadPrimera = ""
    
    If Trim(Variedad) = "" Then Exit Function

    Sql = "select min(codcalid) from rcalidad "
    Sql = Sql & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        CalidadPrimera = DBLet(Rs.Fields(0).Value, "N")
    End If
    
    Set Rs = Nothing
    
End Function


Public Function EsCampoSocioVariedad(campo As String, Socio As String, Variedad As String, Optional EsDesdeClasifica As Boolean) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim VarRelacionada As Long

    EsCampoSocioVariedad = True
    
    If campo = "" Or Socio = "" Or Variedad = "" Then Exit Function
    
    Sql = "select count(*) from rcampos where codcampo = " & DBSet(campo, "N")
    Sql = Sql & " and codsocio = " & DBSet(Socio, "N")
    Sql = Sql & " and codvarie = " & DBSet(Variedad, "N")
    
    Sql2 = "select count(*) from rcampos INNER JOIN  rcampos_cooprop  ON rcampos.codcampo = rcampos_cooprop.codcampo and rcampos.codcampo = " & DBSet(campo, "N")
    Sql2 = Sql2 & " and rcampos_cooprop.codsocio = " & DBSet(Socio, "N")
    Sql2 = Sql2 & " and rcampos.codvarie = " & DBSet(Variedad, "N")
    
    
    '[Monica]23/08/2017: miramos si es de una variedad relacionada
    If EsDesdeClasifica Then
        VarRelacionada = DevuelveValor("select codvarie from variedades_rel where codvarie1 = " & DBSet(Variedad, "N"))
        
        Sql3 = "select count(*) from rcampos WHERE rcampos.codcampo = " & DBSet(campo, "N")
        Sql3 = Sql3 & " and rcampos.codsocio = " & DBSet(Socio, "N")
        Sql3 = Sql3 & " and rcampos.codvarie = " & DBSet(VarRelacionada, "N")
        
'        TieneVariedadRelacionada = (TotalRegistros(Sql3) > 0)
    End If
    
    EsCampoSocioVariedad = (TotalRegistros(Sql) > 0) Or (TotalRegistros(Sql2) > 0) Or (TotalRegistros(Sql3) > 0)

End Function

Public Function EsCampoSocio(campo As String, Socio As String) As Boolean
Dim Sql As String
Dim Sql2 As String


    EsCampoSocio = True
    
    If campo = "" Or Socio = "" Then Exit Function
    
    Sql = "select count(*) from rcampos where codcampo = " & DBSet(campo, "N")
    Sql = Sql & " and codsocio = " & DBSet(Socio, "N")
    
    EsCampoSocio = (TotalRegistros(Sql) > 0) Or (TotalRegistros(Sql2) > 0)

End Function


Public Function ContinuarSiAlbaranImpreso(Albaran As String) As Boolean
Dim Sql As String

    ContinuarSiAlbaranImpreso = True
    Sql = "select impreso from rhisfruta where numalbar = " & DBSet(Albaran, "N")
    If DevuelveValor(Sql) = 1 Then
        If MsgBox("Este Albarán ya ha sido impreso. ¿ Desea Continuar ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            ContinuarSiAlbaranImpreso = False
        End If
    End If

End Function


Public Function ExisteNota(Nota As String) As Boolean
Dim Sql As String
Dim Total As Integer

    ExisteNota = False
    
    Sql = "select count(*) from rentradas where numnotac = " & DBSet(Nota, "N")
    Total = TotalRegistros(Sql)
    If Total = 0 Then
        Sql = "select count(*) from rclasifica where numnotac = " & DBSet(Nota, "N")
        Total = TotalRegistros(Sql)
        If Total = 0 Then
            Sql = "select count(*) from rhisfruta_entradas where numnotac = " & DBSet(Nota, "N")
            Total = TotalRegistros(Sql)
            ExisteNota = (Total <> 0)
        Else
            ExisteNota = True
        End If
    Else
        ExisteNota = True
    End If
    
    
End Function


Public Function ExisteAlbaran(Albaran As String, BBDD As String) As Boolean
Dim Sql As String
Dim Total As Integer

    ExisteAlbaran = False
    
    Sql = "select count(*) from " & BBDD & ".rentradas where numalbar = " & DBSet(Albaran, "N")
    Total = TotalRegistros(Sql)
    If Total = 0 Then
        Sql = "select count(*) from " & BBDD & ".rclasifica where numalbar = " & DBSet(Albaran, "N")
        Total = TotalRegistros(Sql)
        If Total = 0 Then
            Sql = "select count(*) from " & BBDD & ".rhisfruta_entradas where numnotac = " & DBSet(Albaran, "N")
            Total = TotalRegistros(Sql)
            ExisteAlbaran = (Total <> 0)
        Else
            ExisteAlbaran = True
        End If
    Else
        ExisteAlbaran = True
    End If
    
    
End Function











Public Function PonerArticulo(ByRef txtCod As TextBox, ByRef txtNom As TextBox, codAlm As String, tipoMov As String, Optional Modo As Byte, Optional AntCodArtic As String, Optional sConLotes As Boolean) As Boolean
'Poner el codigo y nombre correcto de un Articulo
'IN: txtCod: codigo del articulo
'    txtNom: nombre del articulo
'    codAlm: codigo del almacen en el que comprobamos si se esta inventariando (almacen en el que se va a realizar el movimiento)
Dim vArtic As CArticuloADV
Dim Bloquea As Boolean

    PonerArticulo = False
    sConLotes = False
    
    Set vArtic = New CArticuloADV
    
    If vArtic.Existe(txtCod.Text) Then
        If vArtic.LeerDatos(txtCod.Text) Then
            'comprobar que existe el articulo en el almacen del movimiento
            If vArtic.ExisteEnAlmacen(codAlm) Then
            
                'comprobar si el articulo esta inventariandose
                If vArtic.EnInventario(codAlm) Then
                    If Modo = 1 Then 'Insertar lineas
                        txtCod.Text = ""
                        txtNom.Text = ""
                    End If
                    PonerFoco txtCod
                Else
                    'comprobar si el articulo esta bloqueado
'                    vArtic.MostrarStatusArtic Bloquea
                    
                    If Bloquea Then 'El articulo esta bloqueado
                        If Modo = 1 Then
                            txtCod.Text = ""
                            txtNom.Text = ""
                        End If
                        PonerFoco txtCod
                    Else 'Articulo OK
                        PonerArticulo = True
                        
                        'Si es articulo DE VARIOS podemos modificar la descripción del articulo, sino bloqueamos.
'                        If Not EsArticuloVarios(txtCod.Text) Then
                            BloquearTxt txtNom, True
                            'si insertando lineas
                            'If Modo = 1 Then txtNom.Text = vArtic.Nombre
                            txtNom.Text = vArtic.Nombre
'                        Else
                            'si insertando lineas
'                            If Modo = 1 Then
'                                txtNom.Text = vArtic.Nombre
'                            ElseIf Modo = 2 Then
'                                If txtCod.Text <> AntCodArtic Then txtNom.Text = vArtic.Nombre
'                            End If
'                            BloquearTxt txtNom, False
''                            PonerFoco txtNom
'                        End If

                        Select Case tipoMov
                            Case "OFE", "PEV", "ALV", "ALR", "FAV", "PAR": If vArtic.TextoVentas <> "" Then vArtic.MostrarTextoVen
                            Case "PEC", "ALC", "FAC": If vArtic.TextoCompras <> "" Then vArtic.MostrarTextoCom
                        End Select
                        txtCod.Text = UCase(txtCod.Text)
                        
                        'devolver si el articulo lleva control de lotes
'                        sConLotes = vArtic.TieneNumLote
                        
                    End If
                End If
            End If
        Else
            txtNom.Text = vArtic.Nombre
        End If
    End If
    
    Set vArtic = Nothing
End Function


' grupo 6 es del grupo de bodega (vino)
Public Function EsVariedadGrupo6(Variedad As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from variedades inner join productos on variedades.codprodu = productos.codprodu "
    Sql = Sql & " and variedades.codvarie = " & DBSet(Variedad, "N")
    Sql = Sql & " and productos.codgrupo = 6 "
    
    EsVariedadGrupo6 = (TotalRegistros(Sql) > 0)

End Function

' grupo 5 es del grupo de almazara (olivos)
Public Function EsVariedadGrupo5(Variedad As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from variedades inner join productos on variedades.codprodu = productos.codprodu "
    Sql = Sql & " and variedades.codvarie = " & DBSet(Variedad, "N")
    Sql = Sql & " and productos.codgrupo = 5 "
    
    EsVariedadGrupo5 = (TotalRegistros(Sql) > 0)

End Function


'Para ello le decimos el orden  y ya ta
Public Function NumeroSubcadenasInStr(ByRef cadena As String, Separador As String) As Long
Dim I As Integer
Dim J As Integer
Dim cont As Integer
Dim cad As String

    cont = 0
    If Len(cadena) <> 0 Then
        For I = 1 To Len(cadena)
            If Mid(cadena, I, 1) = Separador Then
                cont = cont + 1
            End If
        Next I
    End If
    
    NumeroSubcadenasInStr = cont

End Function


Public Function CalidadMenut(Variedad As String) As String
Dim Rs As ADODB.Recordset
Dim Sql As String

    CalidadMenut = ""
    
    If Trim(Variedad) = "" Then Exit Function

    Sql = "select codcalid from rcalidad "
    Sql = Sql & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    Sql = Sql & " and rcalidad.tipcalid = 4" ' tipo de calidad de menut
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        CalidadMenut = DBLet(Rs.Fields(0).Value, "N")
    End If
    
    Set Rs = Nothing
    
End Function



Public Function DevuelveNomCalidad(cad As String, PosIni As Integer) As String
Dim I As Integer
Dim Encontrado As Boolean

    DevuelveNomCalidad = ""
    
    I = PosIni
    Encontrado = False
    
    While I < Len(cad) And Not Encontrado
        If Mid(cad, I, 1) = """" Then Encontrado = True
        
        I = I + 1
    Wend

    DevuelveNomCalidad = Mid(cad, PosIni, I - PosIni - 1)

End Function

Public Function ComprobacionRangoGrado(Varie As String, Desde As String, Hasta As String, Linea As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim B As Boolean

    On Error GoTo eComprobacionRangoFechas
    
    ComprobacionRangoGrado = False
    
    Sql = "select rbonifica_lineas.desdegrado, rbonifica_lineas.hastagrado from rbonifica_lineas "
    Sql = Sql & " where codvarie = " & DBSet(Varie, "N")
    
    If Linea <> "" Then
        Sql = Sql & " and numlinea <> " & DBSet(Linea, "N")
    End If
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    B = False
    While Not Rs.EOF And Not B
        B = ((CCur(Desde) <= DBLet(Rs.Fields(0).Value, "N")) And (DBLet(Rs.Fields(0).Value, "N") <= CCur(Hasta)))
        If Not B Then B = ((CCur(Desde) <= DBLet(Rs.Fields(1).Value, "N")) And (DBLet(Rs.Fields(1).Value, "N") <= CCur(Hasta)))
        Rs.MoveNext
    Wend
    
    Rs.Close
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF And Not B
        B = ((DBLet(Rs.Fields(0).Value, "N") <= CCur(Desde)) And (CCur(Desde) <= DBLet(Rs.Fields(1).Value, "N")))
        If Not B Then B = ((DBLet(Rs.Fields(0).Value, "N") <= CCur(Hasta)) And (CCur(Hasta) <= DBLet(Rs.Fields(1).Value, "N")))
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    ComprobacionRangoGrado = Not B
    
    Exit Function
    
eComprobacionRangoFechas:
    MuestraError Err.Number, "Comprobacion de rango grado", Err.Description
End Function



Public Function EsGastodeFactura(Codigo As String) As Boolean
Dim Sql As String

    EsGastodeFactura = False
    
    Sql = "select tipogasto from rconcepgasto where codgasto = " & DBSet(Codigo, "N")
        
    EsGastodeFactura = (DevuelveValor(Sql) = 1)
    
End Function



Public Function HayEntradasCampoSocioVariedad(campo As String, Socio As String, Variedad As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Rs As ADODB.Recordset

    HayEntradasCampoSocioVariedad = True

    Sql = "Select count(*) FROM rentradas where codcampo = " & DBSet(campo, "N")
    Sql = Sql & " and codsocio = " & DBSet(Socio, "N")
    Sql = Sql & " and codvarie = " & DBSet(Variedad, "N")
    
    If TotalRegistros(Sql) = 0 Then
        Sql = "select count(*) from rclasifica where codcampo = " & DBSet(campo, "N")
        Sql = Sql & " and codsocio = " & DBSet(Socio, "N")
        Sql = Sql & " and codvarie = " & DBSet(Variedad, "N")
        
        If TotalRegistros(Sql) = 0 Then
            Sql = "select count(*) from rhisfruta where codcampo = " & DBSet(campo, "N")
            Sql = Sql & " and codsocio = " & DBSet(Socio, "N")
            Sql = Sql & " and codvarie = " & DBSet(Variedad, "N")
        
            HayEntradasCampoSocioVariedad = Not (TotalRegistros(Sql) = 0)
        End If
    End If
    
End Function

Public Function HayEntradasSocio(Socio As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Rs As ADODB.Recordset

    HayEntradasSocio = True

    Sql = "Select count(*) FROM rentradas where codsocio = " & DBSet(Socio, "N")
    
    If TotalRegistros(Sql) = 0 Then
        Sql = "select count(*) from rclasifica where codsocio = " & DBSet(Socio, "N")
        
        If TotalRegistros(Sql) = 0 Then
            Sql = "select count(*) from rhisfruta where codsocio = " & DBSet(Socio, "N")
        
            HayEntradasSocio = Not (TotalRegistros(Sql) = 0)
        End If
    End If
    
End Function


Public Function HayAnticiposPdtesCampoSocioVariedad(campo As String, Socio As String, Variedad As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Rs As ADODB.Recordset

    HayAnticiposPdtesCampoSocioVariedad = True

    Sql = "Select count(*) FROM rfactsoc_variedad, rfactsoc, usuarios.stipom stipom where rfactsoc_variedad.codcampo = " & DBSet(campo, "N")
    Sql = Sql & " and rfactsoc.codsocio = " & DBSet(Socio, "N")
    Sql = Sql & " and rfactsoc_variedad.codvarie = " & DBSet(Variedad, "N")
    Sql = Sql & " and rfactsoc_variedad.descontado = 0 "
    Sql = Sql & " and rfactsoc.codtipom = rfactsoc_variedad.codtipom "
    Sql = Sql & " and rfactsoc.numfactu = rfactsoc_variedad.numfactu "
    Sql = Sql & " and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu "
    Sql = Sql & " and rfactsoc.codtipom = stipom.codtipom "
    Sql = Sql & " and stipom.tipodocu = 1 "
    
    HayAnticiposPdtesCampoSocioVariedad = (TotalRegistros(Sql) <> 0)
    
End Function



Public Function ModificarEntradas(campo As String, SocAnt As String, VarAnt As String, SocNue As String, VarNue As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo eModificarEntradas

    ModificarEntradas = False

    Sql = "update rentradas "
    Sql = Sql & " set codsocio = " & DBSet(SocNue, "N")
    Sql = Sql & ", codvarie = " & DBSet(VarNue, "N")
    Sql = Sql & "  where codcampo = " & DBSet(campo, "N")
    Sql = Sql & " and codsocio = " & DBSet(SocAnt, "N")
    Sql = Sql & " and codvarie = " & DBSet(VarAnt, "N")
    
    conn.Execute Sql
    
    Sql = "update rclasifica "
    Sql = Sql & " set codsocio = " & DBSet(SocNue, "N")
    Sql = Sql & ", codvarie = " & DBSet(VarNue, "N")
    Sql = Sql & "  where codcampo = " & DBSet(campo, "N")
    Sql = Sql & " and codsocio = " & DBSet(SocAnt, "N")
    Sql = Sql & " and codvarie = " & DBSet(VarAnt, "N")
    
    conn.Execute Sql
        
    Sql = "update rhisfruta "
    Sql = Sql & " set codsocio = " & DBSet(SocNue, "N")
    Sql = Sql & ", codvarie = " & DBSet(VarNue, "N")
    Sql = Sql & "  where codcampo = " & DBSet(campo, "N")
    Sql = Sql & " and codsocio = " & DBSet(SocAnt, "N")
    Sql = Sql & " and codvarie = " & DBSet(VarAnt, "N")
    
    conn.Execute Sql
    
    ModificarEntradas = True
    Exit Function
    
eModificarEntradas:
    MuestraError Err.Number, "Modificar Entradas", Err.Description
End Function


Public Sub PonerContRegIndicador(ByRef lblIndicador As Label, ByRef vData As Adodc, cadBuscar As String)
'cuando esta en el MODO 2 pone el label de contador de registros añadiendo
'la palabra 'Busqueda' si es el resultado de una busqueda
'devolvera: "1 de 20" o "BUSQUEDA: 1 de 20"
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    cadReg = PonerContRegistros(vData) 'devuelve: "1 de 20"
    
    If cadBuscar = "" Or cadReg = "" Then
        lblIndicador.Caption = cadReg
    Else
        lblIndicador.Caption = "BUSQUEDA: " & cadReg
    End If
End Sub


Public Sub CargaCadenaAyuda(ByRef vCadena As String, Tipo As Integer)
    Select Case Tipo
        Case 0
             
           ' "____________________________________________________________"
           '           1234567890123456789012345678901234567890123456789012345678901234567
             vCadena = "Tipo de Entradas" & vbCrLf & _
                       "=============" & vbCrLf & vbCrLf
            
             vCadena = vCadena & "Entrada de Retirada: " & vbCrLf & _
                       " Cuando se actualiza la entrada de báscula, el porcentaje de destrio de " & vbCrLf & _
                       " la variedad va a la calidad de destrio y el resto a la calidad de retirada." & vbCrLf & _
                       "  " & vbCrLf & vbCrLf & _
                       "Entrada de Venta Comercio:" & vbCrLf & _
                       " Tiene el mismo tratamiento que las entradas normales, únicamente es " & vbCrLf & _
                       " para diferenciarlas.  " & vbCrLf & vbCrLf
    
    End Select
End Sub




Public Function EsCalidadDestrio(Variedad As String, Calidad As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String

    EsCalidadDestrio = False
    
    If Trim(Variedad) = "" Or Trim(Calidad) = "" Then Exit Function

    Sql = "select tipcalid from rcalidad "
    Sql = Sql & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    Sql = Sql & " and rcalidad.codcalid = " & DBSet(Calidad, "N")
    
    EsCalidadDestrio = (DevuelveValor(Sql) = 1)
    
End Function


Public Function EsCalidadDestrioComercial(Variedad As String, Calidad As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String

    EsCalidadDestrioComercial = False
    
    If Trim(Variedad) = "" Or Trim(Calidad) = "" Then Exit Function

    Sql = "select tipcalid from rcalidad "
    Sql = Sql & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    Sql = Sql & " and rcalidad.codcalid = " & DBSet(Calidad, "N")
    
    EsCalidadDestrioComercial = (DevuelveValor(Sql) = 6)
    
End Function




Public Function EsCalidadMerma(Variedad As String, Calidad As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String

    EsCalidadMerma = False
    
    If Trim(Variedad) = "" Or Trim(Calidad) = "" Then Exit Function

    Sql = "select tipcalid from rcalidad "
    Sql = Sql & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    Sql = Sql & " and rcalidad.codcalid = " & DBSet(Calidad, "N")
    
    EsCalidadMerma = (DevuelveValor(Sql) = 3)
    
End Function







Public Function ActualizarTraza(Nota As String, Variedad As String, Socio As String, campo As String, Fecha As String, Hora As String, MenError As String)
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim IdPalet As Currency

    On Error GoTo eActualizarTraza

    ActualizarTraza = True

    If vParamAplic.HayTraza = False Then Exit Function
    
    Sql = "select idpalet from trzpalets where numnotac = " & DBSet(Nota, "N")
    
    
    'Comprobamos si la fecha de abocamiento de alguno de sus palets es inferior a la de la entrada
    'para no permitir modificar la traza
    Sql2 = "select sum(resul) from (select if(fechahora<" & DBSet(Hora, "FH") & ",1,0) as resul "
    Sql2 = Sql2 & " from trzlineas_cargas where idpalet in (" & Sql & ")) aaaaa "
    If CLng(DevuelveValor(Sql2)) > 0 Then
        MenError = MenError & vbCrLf & "No se permite una fecha de entrada superior a la de abocamiento de ninguno de sus palets. Revise."
        ActualizarTraza = False
        Exit Function
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenDynamic, adLockReadOnly, adCmdText
    
    While Not Rs.EOF
        
        Sql1 = "update trzpalets set codvarie = " & DBSet(Variedad, "N")
        Sql1 = Sql1 & ", codsocio = " & DBSet(Socio, "N")
        Sql1 = Sql1 & ", codcampo = " & DBSet(campo, "N")
        Sql1 = Sql1 & ", fecha = " & DBSet(Fecha, "F")
        Sql1 = Sql1 & ", hora = " & DBSet(Hora, "FH")
        Sql1 = Sql1 & " where idpalet = " & DBSet(Rs.Fields(0).Value, "N")
        
        conn.Execute Sql1
        
        Rs.MoveNext
    Wend
        
    Set Rs = Nothing
    
    Exit Function
    
eActualizarTraza:
    ActualizarTraza = False
    MenError = MenError & vbCrLf & Err.Description
End Function



Public Sub ActualizarClasificacionHco(Albaran As String, Kilos As String)
Dim Sql As String
Dim Variedad As String
Dim Calidad As String
Dim KilosNet As Long

    Variedad = DevuelveValor("select codvarie from rhisfruta where numalbar = " & Albaran)

    If DevuelveValor("select count(*) from rcalidad where codvarie = " & Variedad) = 1 Then
        
        Calidad = DevuelveValor("select codcalid from rcalidad where codvarie = " & Variedad)
        
        If DevuelveValor("select count(*) from rhisfruta_clasif where numalbar = " & Albaran) = 0 Then
            Sql = "insert into rhisfruta_clasif (numalbar,codvarie,codcalid,kilosnet) values (" & DBSet(Albaran, "N") & ","
            Sql = Sql & DBSet(Variedad, "N") & "," & DBSet(Calidad, "N") & "," & DBSet(Kilos, "N") & ")"
            
            conn.Execute Sql
        Else
            KilosNet = DevuelveValor("select sum(kilosnet) from rhisfruta_entradas where numalbar = " & DBSet(Albaran, "N"))
            
            Sql = "update rhisfruta_clasif set kilosnet = " & DBSet(KilosNet, "N")
            Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
            
            conn.Execute Sql
        End If
    
    End If

End Sub


'Funcion de David del Ariges
Public Sub ComprobarCobrosSocio(Codsocio As String, FechaDoc As String, Optional DevuelveImporte As String)
'Comprueba en la tabla de Cobros Pendientes (scobro) de la Base de datos de Contabilidad
'si el cliente tiene alguna factura pendiente de cobro que ha vendido
'con fecha de vencimiento anterior a la fecha del documento: Oferta, Pedido, ALbaran,...
Dim Sql As String, vWhere As String
Dim Codmacta As String
Dim Rs As ADODB.Recordset
Dim cadMen As String
Dim ImporteCred As Currency
Dim Importe As Currency
Dim ImpAux As Currency
Dim vSeccion As CSeccion


    Set Rs = New ADODB.Recordset
    ImporteCred = 0
    
    '[Monica]15/05/2013: Solo para socios de la seccion de hortofruticolo
    If CInt(ComprobarCero(vParamAplic.Seccionhorto)) <> 0 Then
        
        'Obtener la cuenta del socio de la tabla rsocios en Ariagrorec
        Sql = "Select nomsocio, codmaccli from rsocios inner join rsocios_seccion on rsocios.codsocio = rsocios_seccion.codsocio "
        Sql = Sql & " where rsocios_seccion.codsecci = " & DBSet(vParamAplic.Seccionhorto, "N") & " and rsocios.codsocio=" & Codsocio
        Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Rs.EOF Then
            Sql = ""
        Else
            Codsocio = Codsocio & " - " & Rs!nomsocio
            Codmacta = DBLet(Rs!codmaccli)
            ImporteCred = 0
            If Codmacta = "" Then Sql = ""
        End If
        Rs.Close
        If Sql = "" Then Exit Sub
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            If vSeccion.AbrirConta Then
        
                'AHORA FEBRERO 2010
                If vParamAplic.ContabilidadNueva Then
                    Sql = "SELECT cobros.* FROM cobros INNER JOIN formapago ON cobros.codforpa=formapago.codforpa "
                    vWhere = " WHERE cobros.codmacta = '" & Codmacta & "'"
    ' lo llamamos desde el mto de socios, campos y contadores
    '                vWhere = vWhere & " AND fecvenci <= ' " & Format(FechaDoc, FormatoFecha) & "' "
                    'Antes mayo 2010
                    'vWhere = vWhere & " AND (sforpa.tipforpa between 0 and 3)"
                    vWhere = vWhere & " AND recedocu=0 "
                    Sql = Sql & vWhere & " ORDER BY fecfactu, numfactu "
                Else
                    Sql = "SELECT scobro.* FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
                    vWhere = " WHERE scobro.codmacta = '" & Codmacta & "'"
    ' lo llamamos desde el mto de socios, campos y contadores
    '                vWhere = vWhere & " AND fecvenci <= ' " & Format(FechaDoc, FormatoFecha) & "' "
                    'Antes mayo 2010
                    'vWhere = vWhere & " AND (sforpa.tipforpa between 0 and 3)"
                    vWhere = vWhere & " AND recedocu=0 "
                    Sql = Sql & vWhere & " ORDER BY fecfaccl, codfaccl "
                End If
                'Lee de la Base de Datos de CONTABILIDAD
                Rs.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
                Importe = 0
                While Not Rs.EOF
                
                    'QUITO LO DE DEVUELTO. MAYO 2010
                    'If Val(RS!Devuelto) = 1 Then
                    '    'SALE SEGURO (si no esta girado otra vez ¿no?
                    '    'Si esta girado otra vez tendra impcobro, con lo cual NO tendra diferencia de importes
                    '    Impaux = RS!ImpVenci + DBLet(RS!gastos, "N") - DBLet(RS!impcobro, "N")
                        
                    'Else
                        'Si esta recibido NO lo saco
                        If Val(Rs!recedocu) = 1 Then
                            ImpAux = 0
                        Else
                            'NO esta recibido. Si tiene diferencia
                            ImpAux = Rs!ImpVenci + DBLet(Rs!Gastos, "N") - DBLet(Rs!impcobro, "N")
                    
                        End If
                '    End If
                    If ImpAux <> 0 Then Importe = Importe + ImpAux
                    Rs.MoveNext
                Wend
                Rs.Close
                Set Rs = Nothing
                '[Monica]30/01/2017: estaba puesto >
                If Importe <> 0 Then
                    If DevuelveImporte <> "" Then
                        'Meto aqui el importer
                        DevuelveImporte = CStr(Importe)
                    Else
                        cadMen = "El Socio tiene facturas vencidas con valor de: " & Format(Importe, FormatoImporte) & " €."
                        If ImporteCred > 0 Then cadMen = cadMen & vbCrLf & "Límite crédito: " & Format(ImporteCred, FormatoImporte) & " €."
                        cadMen = cadMen & vbCrLf & "¿Desea Ver Detalle?"
                        If MsgBox(cadMen, vbYesNo + vbQuestion + vbDefaultButton2, "Cobros Pendientes") = vbYes Then
                            'Mostrar los detalles de los cobros pendientes
                            frmMensajes.cadWHERE = vWhere
                            frmMensajes.vCampos = Codsocio
                            frmMensajes.OpcionMensaje = 1
                            frmMensajes.Show vbModal
                        End If
                    End If
                End If
                
            End If
            vSeccion.CerrarConta
        End If
   End If
            
End Sub









Public Function ObtenerPrecioRecoldeCalidad(Variedad As String, Calidad As String, Tipo As Byte) As Currency
'Tipo = 0, factura transporte socio gastos de recoleccion
'Tipo = 1, gastos de recoleccion del recolector
Dim Sql As String

    Sql = "select "
    If Tipo = 0 Then
        Sql = Sql & "eurrecsoc "
    Else
        Sql = Sql & "eurreccoop "
    End If
    Sql = Sql & " from rcalidad where codvarie = " & DBSet(Variedad, "N") & " and codcalid = " & DBSet(Calidad, "N")

    ObtenerPrecioRecoldeCalidad = DevuelveValor(Sql)

End Function


Public Function CodTipomAnticipos() As String
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Result As String

    CodTipomAnticipos = ""

    Sql = "select codtipom from usuarios.stipom where tipodocu = 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Result = ""
    
    While Not Rs.EOF
        Result = Result & DBSet(Rs!CodTipom, "T") & ","
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If Result <> "" Then
        Result = Mid(Result, 1, Len(Result) - 1)
    End If

    CodTipomAnticipos = Result
    
End Function

Public Function EsCaja(CodEnvase As String) As Boolean
Dim Sql As String

    Sql = "select escaja from confenva where codtipen = " & DBSet(CodEnvase, "N")
    
    EsCaja = DevuelveValor(Sql)

End Function


' FUNCIONES DE POZOS PARA SABER SI ES O NO CONTADO (ESCALONA Y UTXERA)

Public Function EsSocioContadoPOZOS(Socio As String) As Boolean
Dim Sql As String

    Sql = "select cuentaba from rsocios where codsocio = " & DBSet(Socio, "N")
    EsSocioContadoPOZOS = (DevuelveValor(Sql) = "8888888888")

End Function


Public Function EsReciboContadoPOZOS(vWhere As String) As Boolean
Dim Sql As String

    Sql = "select escontado from rrecibpozos where " & vWhere
    EsReciboContadoPOZOS = (DevuelveValor(Sql) = "1")

End Function

Public Function EntregadaFichaCultivo(campo As String) As Boolean
Dim Sql As String

    Sql = "select entregafichaculti from rcampos where codcampo = " & DBSet(campo, "N")

    EntregadaFichaCultivo = (DevuelveValor(Sql) = "0")
    
End Function


Public Function TrabajadorDeBaja(Traba As String, Optional Fecha As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "select fechabaja from straba where codtraba = " & DBSet(Traba, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        TrabajadorDeBaja = (DBLet(Rs!FechaBaja, "F") <> "")
        If Fecha <> "" Then
            TrabajadorDeBaja = (DBLet(Rs!FechaBaja, "F") >= Fecha)
        End If
    Else
        TrabajadorDeBaja = True
    End If
    Set Rs = Nothing
    

End Function


Public Function HayXML() As Boolean
Dim Sql As String

    Sql = "select xml from rparam "
    HayXML = (DevuelveValor(Sql) = 1)


End Function

Public Function EsCalidadConBonificacion(Variedad As String, Calidad As String) As Boolean
Dim Sql As String

    Sql = "select seaplicabonif from rcalidad where codvarie = " & DBSet(Variedad, "N") & " and codcalid = " & DBSet(Calidad, "N")
    
    EsCalidadConBonificacion = (DevuelveValor(Sql) = 1)

End Function

'Si es "" devuelve "" , si no, devuelve el campo formateado
Public Function MiFormat(Valor As String, Formato As String) As String
    If Trim(Valor) = "" Then
       MiFormat = ""
    Else
        If Formato = "" Then
            MiFormat = Valor
        Else
            MiFormat = Format(Valor, Formato)
        End If
    End If
End Function

Public Function SeAplicaPixat(Variedad As String, Fecha As String) As Boolean
Dim Sql As String

    Sql = "select * from variedades where codvarie = " & DBSet(Variedad, "N") & " and fecinipixat <= " & DBSet(Fecha, "F")
    Sql = Sql & " and fecfinpixat >= " & DBSet(Fecha, "F")

    SeAplicaPixat = (TotalRegistrosConsulta(Sql) > 0)

End Function


Public Function EstaEnDocumentoBaja(vCampo As String) As Boolean
Dim Sql As String

    Sql = "select * from rsocios_movim where codcampo = " & DBSet(vCampo, "N")
    EstaEnDocumentoBaja = (TotalRegistrosConsulta(Sql) <> 0)
    
End Function


Public Function EntradaClasificada(Nota As Long) As Boolean
Dim Sql As String

    EntradaClasificada = False
    
    Sql = "select sum(coalesce(kilosnet,0)) from rclasifica_clasif where numnotac = " & DBSet(Nota, "N")
    
    EntradaClasificada = (DevuelveValor(Sql) <> 0)

End Function

Public Function GrupoTrabajo(Trabajador As String) As String
Dim Sql As String

    Sql = DevuelveValor("select codbanpr from straba where codtraba = " & DBSet(Trabajador, "N"))
    If Sql = "0" Then Sql = ""
    
    GrupoTrabajo = Sql

End Function


