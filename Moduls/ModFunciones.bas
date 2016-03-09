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

    CompForm = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox And Control.visible = True Then
            Carga = mTag.Cargar(Control)
            If Carga = True Then
                Correcto = mTag.Comprobar(Control)
                If Not Correcto Then Exit Function
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
                            MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
                            Exit Function
                    End If
                End If
            End If
        End If
    Next Control
    CompForm = True
End Function

'A�ade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function CompForm2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    CompForm2 = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox Then
            If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                Carga = mTag.Cargar(Control)
                If Carga = True Then
                    Correcto = mTag.Comprobar(Control)
                    If Not Correcto Then Exit Function
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
                                    MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
                                    Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next Control
    CompForm2 = True
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
Dim i As Integer
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
Dim Cad As String
    
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
                        Cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & Cad
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
                    Cad = "1"
                    Else
                    Cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                Der = Der & Cad
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
                        Cad = ValorNulo
                    Else
                        Cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & Cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Pr�cticas');
    
    Cad = "INSERT INTO " & mTag.Tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
    
    conn.Execute Cad, , adCmdText
    
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
Dim Cad As String
    
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
                            Cad = ValorParaSQL(Control.Text, mTag)
                            If Der <> "" Then Der = Der & ","
                            Der = Der & Cad
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
                        Cad = "1"
                        Else
                        Cad = "0"
                    End If
                    If Der <> "" Then Der = Der & ","
                    If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                    Der = Der & Cad
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
                            Cad = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            Cad = Control.ItemData(Control.ListIndex)
                        Else
                            Cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        If Der <> "" Then Der = Der & ","
                        Der = Der & Cad
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
                            Cad = Control.Index
                            If Der <> "" Then Der = Der & ","
                            Der = Der & Cad
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
                            Cad = ValorParaSQL(Control.Value, mTag)
                        Else
                            Cad = ValorNulo
                        End If
                        If Der <> "" Then Der = Der & ","
                        Der = Der & Cad
                    End If
                End If
            End If
        End If
        
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Pr�cticas');
    
    Cad = "INSERT INTO " & mTag.Tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
    conn.Execute Cad, , adCmdText
    
     ' ### [Monica] 18/12/2006
    CadenaCambio = Cad
   
    InsertarDesdeForm2 = True
Exit Function

EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function


Public Function CadenaInsertarDesdeForm(ByRef formulario As Form) As String
'Equivale a InsertarDesdeForm, excepto que devuelve la candena SQL y hace el execute fuera de la funci�n.
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim Cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    CadenaInsertarDesdeForm = ""
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
                        Cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & Cad
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
                    Cad = "1"
                    Else
                    Cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                Der = Der & Cad
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
                        Cad = ValorNulo
                    Else
                        Cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & Cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Pr�cticas');
    
    Cad = "INSERT INTO " & mTag.Tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
'    Conn.Execute cad, , adCmdText
    
    CadenaInsertarDesdeForm = Cad
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function


Public Function PonerCamposForma(ByRef formulario As Form, ByRef vData As Adodc) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Cad As String
Dim Valor As Variant
Dim campo As String  'Campo en la base de datos
Dim i As Integer


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
                                Cad = Format(Valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = Cad
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
                    '++MONICA: 15-01-2008 a�adida la condicion de que el valor sea nulo
                    If IsNull(vData.Recordset.Fields(campo)) Then
                        Control.ListIndex = -1
                    Else
                    '++
                        i = 0
                        For i = 0 To Control.ListCount - 1
                            If Control.ItemData(i) = Val(Valor) Then
                                Control.ListIndex = i
                                Exit For
                            End If
                        Next i
                        If i = Control.ListCount Then Control.ListIndex = -1
                    '++MONICA: 15-01-2008 a�adida la condicion de que el valor sea nulo
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

'A�ade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function PonerCamposForma2(ByRef formulario As Form, ByRef vData As Adodc, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Cad As String
Dim Valor As Variant
Dim campo As String  'Campo en la base de datos
Dim i As Integer
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
                                    Cad = Format(Valor, mTag.Formato)
                                    'Antiguo
                                    'Control.Text = TransformaComasPuntos(cad)
                                    'nuevo
                                    Control.Text = Cad
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
                        i = 0
                        For i = 0 To Control.ListCount - 1
                            If Control.ItemData(i) = Val(Valor) Then
                                Control.ListIndex = i
                                Exit For
                            End If
                        Next i
                        If i = Control.ListCount Then Control.ListIndex = -1
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
Dim Cad As String
Dim Valor As Variant
Dim camp As String  'Camp en la BDA
Dim i As Integer

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
                            Cad = Format(Valor, mTag.Formato)
                            Control.Text = Cad
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


Private Function ObtenerMaximoMinimo(vSQL As String, Optional vBD As Byte) As String
Dim RS As Recordset
    ObtenerMaximoMinimo = ""
    Set RS = New ADODB.Recordset
    If vBD = cConta Then
        RS.Open vSQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else
        RS.Open vSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    End If
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            ObtenerMaximoMinimo = CStr(RS.Fields(0))
        End If
    End If
    RS.Close
    Set RS = Nothing
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
'    MuestraError Err.Number, "Obtener b�squeda. "
'End Function

Public Function ObtenerBusqueda(ByRef formulario As Form, Optional CHECK As String, Optional vBD As Byte, Optional cadWHERE As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim Cad As String
    Dim SQL As String
    Dim Tabla As String
    Dim Rc As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda = ""
    SQL = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                If Control.Tag <> "" Then
                    Carga = mTag.Cargar(Control)
                    If Carga Then
                        If Aux = ">>" Then
                            Cad = " MAX("
                        Else
                            Cad = " MIN("
                        End If
                        'monica
                        Select Case mTag.TipoDato
                            Case "FHF"
                                Cad = Cad & "date(" & mTag.columna & "))"
                            Case "FHH"
                                Cad = Cad & "time(" & mTag.columna & "))"
                            Case Else
                                Cad = Cad & mTag.columna & ")"
                        End Select
                        
                        SQL = "Select " & Cad & " from " & mTag.Tabla
                        If cadWHERE <> "" Then SQL = SQL & " WHERE " & cadWHERE
                        SQL = ObtenerMaximoMinimo(SQL, vBD)
                        Select Case mTag.TipoDato
                        Case "N"
                            SQL = mTag.Tabla & "." & mTag.columna & " = " & TransformaComasPuntos(SQL)
                        Case "F"
                            SQL = mTag.Tabla & "." & mTag.columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
                        Case "FHF"
                            SQL = "date(" & mTag.Tabla & "." & mTag.columna & ") = '" & Format(SQL, "yyyy-mm-dd") & "'"
                        Case "FHH"
                            SQL = "time(" & mTag.Tabla & "." & mTag.columna & ") = '" & Format(SQL, "hh:mm:ss") & "'"
                        Case Else
                            '[Monica]04/03/2013: quito las comillas
                            SQL = mTag.Tabla & "." & mTag.columna & " = " & DBSet(SQL, "T") ' & "'"
                        End Select
                        SQL = "(" & SQL & ")"
                    End If
                End If
            End If
        End If
    Next


'++monica: lo he a�adido del anterior obtenerbusqueda
    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then

                    SQL = mTag.Tabla & "." & mTag.columna & " is NULL"
                    SQL = "(" & SQL & ")"
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
                        If mTag.Tabla <> "" Then
                            Tabla = mTag.Tabla & "."
                            Else
                            Tabla = ""
                        End If
                        Rc = SeparaCampoBusqueda(mTag.TipoDato, Tabla & mTag.columna, Aux, Cad)
                        If Rc = 0 Then
                            If SQL <> "" Then SQL = SQL & " AND "
                            SQL = SQL & "(" & Cad & ")"
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
                            Cad = Control.ItemData(Control.ListIndex)
                        Else
                            Cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        Cad = mTag.Tabla & "." & mTag.columna & " = " & Cad
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is CheckBox Then
            '=============== A�ade: Laura, 15/04/05
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    Aux = ""
                    If CHECK <> "" Then
                        Tabla = DBLet(Control.Index, "T")
                        If Tabla <> "" Then Tabla = "(" & Tabla & ")"
                        Tabla = Control.Name & Tabla & "|"
                        If InStr(1, CHECK, Tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    If Aux <> "" Then
'                    If Control.Value = 1 Then
                        Cad = Control.Value
                        Cad = mTag.Tabla & "." & mTag.columna & " = " & Cad
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If
                End If
            End If
            '===================
        End If
    Next Control
    ObtenerBusqueda = SQL
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda = ""
    MuestraError Err.Number, "Obtener b�squeda. " & vbCrLf & Err.Description
End Function

'A�ade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function ObtenerBusqueda2(ByRef formulario As Form, Optional CHECK As String, Optional opcio As Integer, Optional nom_frame As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim Cad As String
    Dim SQL As String
    Dim Tabla As String
    Dim Rc As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda2 = ""
    SQL = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Aux = ">>" Then
                            Cad = " MAX(" & mTag.columna & ")"
                        Else
                            Cad = " MIN(" & mTag.columna & ")"
                        End If
                        SQL = "Select " & Cad & " from " & mTag.Tabla
                        SQL = ObtenerMaximoMinimo(SQL)
                        Select Case mTag.TipoDato
                        Case "N"
                            SQL = mTag.Tabla & "." & mTag.columna & " = " & TransformaComasPuntos(SQL)
                        Case "F"
                            SQL = mTag.Tabla & "." & mTag.columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
                        Case Else
                            '[Monica]04/03/2013: quito las comillas y pongo el dbset
                            SQL = mTag.Tabla & "." & mTag.columna & " = " & DBSet(SQL, "T") ' & "'"
                        End Select
                        SQL = "(" & SQL & ")"
                    End If
                End If
            End If
        End If
    Next

'++monica: lo he a�adido del anterior obtenerbusqueda
    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then

                    SQL = mTag.Tabla & "." & mTag.columna & " is NULL"
                    SQL = "(" & SQL & ")"
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
                        If mTag.Tabla <> "" Then
                            Tabla = mTag.Tabla & "."
                            Else
                            Tabla = ""
                        End If
                        Rc = SeparaCampoBusqueda(mTag.TipoDato, Tabla & mTag.columna, Aux, Cad)
                        If Rc = 0 Then
                            If SQL <> "" Then SQL = SQL & " AND "
                            SQL = SQL & "(" & Cad & ")"
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
            If Control.Tag <> "" Then ' +-+- 12/05/05: canvi de C�sar, no te sentit passar-li un control que no t� TAG +-+-
                mTag.Cargar Control
                If mTag.Cargado Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Control.ListIndex > -1 Then
                            Cad = Control.ItemData(Control.ListIndex)
                            Cad = mTag.Tabla & "." & mTag.columna & " = " & Cad
                            If SQL <> "" Then SQL = SQL & " AND "
                            SQL = SQL & "(" & Cad & ")"
                        End If
                    End If
                End If
            End If
            
         ElseIf TypeOf Control Is CheckBox Then
            '=============== A�ade: Laura, 27/04/05
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    ' a�adido 12022007
                    Aux = ""
                    If CHECK <> "" Then
                        Tabla = DBLet(Control.Index, "T")
                        If Tabla <> "" Then Tabla = "(" & Tabla & ")"
                        Tabla = Control.Name & Tabla & "|"
                        If InStr(1, CHECK, Tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    If Aux <> "" Then
'                    If Control.Value = 1 Then
                        Cad = Control.Value
                        Cad = mTag.Tabla & "." & mTag.columna & " = " & Cad
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If
                End If
            End If
            '===================
        End If
    Next Control
    ObtenerBusqueda2 = SQL
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda2 = ""
    MuestraError Err.Number, "Obtener b�squeda. " & vbCrLf & Err.Description
End Function

'A�ado Optional CHECK As String. Para poder realizar las busquedas con los checks
'monica corresponde al ObtenerBusqueda de laura
Public Function ObtenerBusqueda3(ByRef formulario As Form, paraRPT As Boolean, Optional CHECK As String) As String
Dim Control As Object
Dim Carga As Boolean
Dim mTag As CTag
Dim Aux As String
Dim Cad As String
Dim SQL As String
Dim Tabla As String, columna As String
Dim Rc As Byte

    On Error GoTo EObtenerBusqueda3

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda3 = ""
    SQL = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.visible Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Aux = ">>" Then
                        If Not paraRPT Then
                            Cad = " MAX(" & mTag.columna & ")"
                        Else
                            Cad = " MAX({" & mTag.Tabla & "." & mTag.columna & "})"
                        End If
                    Else
                        If Not paraRPT Then
                            Cad = " MIN(" & mTag.columna & ")"
                        Else
                            Cad = " MIN({" & mTag.Tabla & "." & mTag.columna & "})"
                        End If
                    End If
                    If Not paraRPT Then
                        SQL = "Select " & Cad & " from " & mTag.Tabla
                    Else
                        SQL = "Select " & Cad & " from {" & mTag.Tabla & "}"
                    End If
                    SQL = ObtenerMaximoMinimo(SQL)
                    Select Case mTag.TipoDato
                    Case "N"
                        If SQL <> "" Then
                            If Not paraRPT Then
                                SQL = mTag.Tabla & "." & mTag.columna & " = " & TransformaComasPuntos(SQL)
                            Else
                                SQL = "{" & mTag.Tabla & "." & mTag.columna & "} = " & TransformaComasPuntos(SQL)
                            End If
                        End If
                    Case "F"
                        If SQL = "" Then SQL = "0000-00-00"
                        If Not paraRPT Then
                            SQL = mTag.Tabla & "." & mTag.columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
                        Else
                            SQL = "{" & mTag.Tabla & "." & mTag.columna & "} = '" & Format(SQL, "yyyy-mm-dd") & "'"
                        End If
                    Case Else
                        '[Monica]04/03/2013: quito comillas
                        If Not paraRPT Then
                            SQL = mTag.Tabla & "." & mTag.columna & " = " & DBSet(SQL, "T") '& "'"
                        Else
                            SQL = "{" & mTag.Tabla & "." & mTag.columna & "} = " & DBSet(SQL, "T") ' & "'"
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
                        SQL = mTag.Tabla & "." & mTag.columna & " is NULL"
                    Else
                        SQL = "{" & mTag.Tabla & "." & mTag.columna & "} is NULL"
                    End If
                    SQL = "(" & SQL & ")"
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
                        If mTag.Tabla <> "" Then
                            If Not paraRPT Then
                                Tabla = mTag.Tabla & "."
                            Else
                                Tabla = "{" & mTag.Tabla & "."
                            End If
                        Else
                            Tabla = ""
                        End If
                        If Not paraRPT Then
                            columna = mTag.columna
                        Else
                            columna = mTag.columna & "}"
                        End If
                    Rc = SeparaCampoBusqueda3(mTag.TipoDato, Tabla & columna, Aux, Cad, paraRPT)
                    If Rc = 0 Then
                        If SQL <> "" Then SQL = SQL & " AND "
                        If Not paraRPT Then
                            SQL = SQL & "(" & Cad & ")"
                        Else
                            SQL = SQL & "(" & Cad & ")"
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
                        Cad = Control.ItemData(Control.ListIndex)
                        If Not paraRPT Then
                            Cad = mTag.Tabla & "." & mTag.columna & " = " & Cad
                        Else
                            Cad = "{" & mTag.Tabla & "." & mTag.columna & "} = " & Cad
                        End If
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    Else
                        Cad = Control.List(Control.ListIndex)
                        If Not paraRPT Then
                            Cad = mTag.Tabla & "." & mTag.columna & " = '" & Cad & "'"
                        Else
                            Cad = "{" & mTag.Tabla & "." & mTag.columna & "} = '" & Cad & "'"
                        End If
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
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
                        Tabla = NombreCheck & "|"
                        If InStr(1, CHECK, Tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    
                    If Aux <> "" Then
                        If Not paraRPT Then
                            Cad = mTag.Tabla & "." & mTag.columna
                        Else
                            Cad = "{" & mTag.Tabla & "." & mTag.columna & "} "
                        End If
                        
                        Cad = Cad & " = " & Aux
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If 'cargado
                End If '<>""
            End If
        End If
    
    Next Control
    ObtenerBusqueda3 = SQL
Exit Function
EObtenerBusqueda3:
    ObtenerBusqueda3 = ""
    MuestraError Err.Number, "Obtener b�squeda. "
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
    'WHERE Pa�sDestinatario = 'M�xico';
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
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
    'WHERE Pa�sDestinatario = 'M�xico';
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
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
    'WHERE Pa�sDestinatario = 'M�xico';
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    conn.Execute Aux, , adCmdText

    ModificaDesdeFormulario1 = True
    Exit Function
    
EModificaDesdeFormulario1:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function

Public Sub FormateaCampo(vTex As TextBox)
    Dim mTag As CTag
    Dim Cad As String
    On Error GoTo EFormateaCampo
    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        If vTex.Text <> "" Then
            If mTag.Formato <> "" Then
                Cad = TransformaPuntosComas(vTex.Text)
                Cad = Format(Cad, mTag.Formato)
                vTex.Text = Cad
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
Dim Cad As String
    
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


'A�ade: CESAR
'Para utilizalo en el arreglaGrid
Public Function FormatoCampo2(ByRef objec As Object) As String
'Devuelve el formato del campo en el TAg: "0000"
Dim mTag As CTag
Dim Cad As String

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
Dim Cad As String

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
Dim i As Integer
Dim J As Integer
Dim cont As Integer
Dim Cad As String

    i = 0
    cont = 1
    Cad = ""
    Do
        J = i + 1
        i = InStr(J, cadena, "|")
        If i > 0 Then
            If cont = Orden Then
                Cad = Mid(cadena, J, i - J)
                i = Len(cadena) 'Para salir del bucle
                Else
                    cont = cont + 1
            End If
        End If
    Loop Until i = 0
    RecuperaValor = Cad
End Function


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValorNew(ByRef cadena As String, Separador As String, Orden As Integer) As String
Dim i As Integer
Dim J As Integer
Dim cont As Integer
Dim Cad As String

    i = 0
    cont = 1
    Cad = ""
    Do
        J = i + 1
        i = InStr(J, cadena, Separador)
        If i > 0 Then
            If cont = Orden Then
                Cad = Mid(cadena, J, i - J)
                i = Len(cadena) 'Para salir del bucle
                Else
                    cont = cont + 1
            End If
        End If
    Loop Until i = 0
    RecuperaValorNew = Cad
End Function




'-----------------------------------------------------------------------
'Deshabilitar ciertas opciones del menu
'EN funcion del nivel de usuario
'Esto es a nivel general, cuando el Toolba es el mismo

'Para ello en el tag del button tendremos k poner un numero k nos diara hasta k nivel esta permitido

Public Sub PonerOpcionesMenuGeneral(ByRef formulario As Form)
Dim i As Integer
Dim J As Integer
'Dim bol As Boolean

On Error GoTo EPonerOpcionesMenuGeneral
'bol = vSesion.Nivel < 2

'A�adir, modificar y borrar deshabilitados si no nivel
With formulario
    For i = 1 To .Toolbar1.Buttons.Count
        If .Toolbar1.Buttons(i).Tag <> "" Then
            J = Val(.Toolbar1.Buttons(i).Tag)
            If J < vUsu.Nivel Then
                .Toolbar1.Buttons(i).Enabled = False
            End If
        End If
    Next i
End With

Exit Sub
EPonerOpcionesMenuGeneral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub


Public Sub PonerModoMenuGral(ByRef formulario As Form, activo As Boolean)
Dim i As Integer
'Dim j As Integer

On Error GoTo PonerModoMenuGral

'A�adir, modificar y borrar deshabilitados si no Modo
    With formulario
        For i = 1 To .Toolbar1.Buttons.Count
            Select Case .Toolbar1.Buttons(i).ToolTipText
                Case "Nuevo"
                    .Toolbar1.Buttons(i).visible = Not .DeConsulta
                Case "Modificar", "Eliminar", "Imprimir"
                    .Toolbar1.Buttons(i).visible = Not .DeConsulta
                    .Toolbar1.Buttons(i).Enabled = activo
'                Case "Modificar"
'                Case "Eliminar"
'                Case "Imprimir"
            End Select
        Next i
        
        
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
Dim i As Integer
Dim J As Integer
'Dim bol As Boolean

On Error GoTo EPonerOpcionesMenuGeneralNew
'bol = vSesion.Nivel < 2
'A�adir, modificar y borrar deshabilitados si no nivel
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
Dim i As Integer

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
    Aux = "UPDATE " & mTag.Tabla
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
        If TypeOf Control Is TextBox And Control.visible = True Then
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
    Next Control

    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "select * FROM " & mTag.Tabla
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


'A�ade: CESAR
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
        Aux = "select * FROM " & mTag.Tabla
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
                '�Ya existe el registro, luego esta bloqueada
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
Dim SQL As String

'Solo me interesa la tabla
On Error Resume Next
    Set mTag = New CTag
    mTag.Cargar TextBoxConTag
    If mTag.Cargado Then
'        SQL = "DELETE from zBloqueos where codusu=" & vUsu.Codigo & " and tabla='" & mTag.tabla & "'"
        conn.Execute SQL
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
Dim Cad As String
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
                            Cad = ValorParaSQL(Control.Text, mTag)
                            Izda = Izda & Cad
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
                            Cad = "1"
                            Else
                            Cad = "0"
                        End If
                        If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                        Izda = Izda & Cad
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
                            Cad = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            Cad = Control.ItemData(Control.ListIndex)
                        Else
                            Cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        Izda = Izda & Cad
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
                            Cad = Control.Index
                            Izda = Izda & Cad
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
                            Cad = ValorParaSQL(Control.Value, mTag)
                        Else
                            Cad = ValorNulo
                        End If
                        Izda = Izda & Cad
                    End If
                End If
            End If
        End If
        
    Next Control

    ValorAnterior = Izda

End Sub


Public Sub CalcularImporteNue(ByRef Cantidad As TextBox, ByRef Precio As TextBox, ByRef Importe As TextBox, tipo As Integer)
'Calcula el Importe de una linea de hcode facturas
Dim vImp As Currency
Dim vCan As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    Cantidad = ComprobarCero(Cantidad.Text)
    Precio = ComprobarCero(Precio.Text)
    Importe = ComprobarCero(Importe.Text)
    
    Select Case tipo
        Case 0 ' me han introducido la cantidad
            vImp = CCur(ImporteFormateado(Cantidad.Text)) * CCur(ImporteFormateado(Precio.Text))
            vImp = Round2(vImp, 2)
            Importe.Text = Format(vImp, "###,##0.00")
        Case 1 ' me han introducido el precio
            vImp = CCur(ImporteFormateado(Cantidad.Text)) * CCur(ImporteFormateado(Precio.Text))
            vImp = Round2(vImp, 2)
            Importe.Text = Format(vImp, "###,##0.00")
        Case 2 ' me han introducido el importe
            vCan = CCur(ImporteFormateado(Importe.Text)) / CCur(ImporteFormateado(Precio.Text))
            vCan = Round2(vCan, 3)
            Cantidad.Text = Format(vCan, "##,##0.000")
    End Select
    
End Sub


'Public Function PonerNomEmple(codEmp As String) As String
'Dim nomEmp As String
'Dim cad As String
'
'    'apellidos i nombre del empleado
'    If (codEmp <> "") Then
'        nomEmp = "nomemple"
'        cad = DevuelveDesdeBDNew(cAgro, "empleado", "apeemple", "codemple", codEmp, "N", nomEmp, "codempre", CStr(vSesion.Empresa), "N", "codagenc", CStr(vSesion.Agencia), "N")
'        If cad <> "" Then cad = cad & ", " & nomEmp
'    End If
'    PonerNomEmple = cad
'End Function



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
                    devuelve = DevuelveDesdeBDNew(cAgro, vtag.Tabla, vtag.columna, vtag.columna, T.Text, vtag.TipoDato)
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
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar c�digo.", Err.Description
End Function




Public Function TotalRegistros(vSQL As String) As Long
'Devuelve el valor de la SQL
'para obtener COUNT(*) de la tabla
Dim RS As ADODB.Recordset

    On Error Resume Next

    Set RS = New ADODB.Recordset
    RS.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalRegistros = 0
    If Not RS.EOF Then
        If RS.Fields(0).Value > 0 Then TotalRegistros = RS.Fields(0).Value  'Solo es para saber que hay registros que mostrar
    End If
    RS.Close
    Set RS = Nothing

    If Err.Number <> 0 Then
        TotalRegistros = 0
        Err.Clear
    End If
End Function



Public Function Round2(Number As Variant, Optional NumDigitsAfterDecimals As Long) As Variant
Dim Ent As Integer
Dim Cad As String

  ' Comprobaciones

  If Not IsNumeric(Number) Then
    Err.Raise 13, "Round2", "Error de tipo. Ha de ser un n�mero."
    Exit Function
  End If

  If NumDigitsAfterDecimals < 0 Then
    Err.Raise 0, "Round2", "NumDigitsAfterDecimals no puede ser negativo."
    Exit Function
  End If

  ' Redondeo.

  Cad = "0"
  If NumDigitsAfterDecimals <> 0 Then Cad = Cad & "." & String(NumDigitsAfterDecimals, "0")
  Round2 = Val(TransformaComasPuntos(Format(Number, Cad)))

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


Public Function CalcularImporte(Cantidad As String, Precio As String, Dto1 As String, Dto2 As String, TipoDto As Byte, ImpDto As String, Optional Bruto As String) As String
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
    Cantidad = ComprobarCero(Cantidad)
    vPre = ComprobarCero(Precio)
    Dto1 = ComprobarCero(Dto1)
    Dto2 = ComprobarCero(Dto2)
    
    If Bruto <> "" Then
        vImp = CCur(Bruto) - CCur(ImpDto)
    Else
        vImp = (CCur(Cantidad) * CCur(vPre)) - CCur(ImpDto)
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
                '�Ya existe el registro, luego esta bloqueada
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
Dim SQL As String

'Solo me interesa la tabla
On Error Resume Next

        SQL = "DELETE FROM zbloqueos WHERE codusu=" & vUsu.Codigo & " and tabla='" & cadTabla & "'"
        conn.Execute SQL
        If Err.Number <> 0 Then
            Err.Clear
        End If
End Function

'++monica
'funcion de la libreria general de gessocial de Rafa, necesaria para pasar al aridoc
Public Function CApos(Texto As String) As String
    Dim i As Integer
    i = InStr(1, Texto, "'")
    If i = 0 Then
        CApos = Texto
    Else
        CApos = Mid(Texto, 1, i) & "'" & Mid(Texto, i + 1, Len(Texto) - i)
    End If
    '-- Ya que estamos transformamos las �
    Texto = CApos
    i = InStr(1, Texto, "�")
    If i = 0 Then
        CApos = Texto
    Else
        CApos = Mid(Texto, 1, i - 1) & "�" & Mid(Texto, i + 1, Len(Texto) - i)
    End If
    '-- Y otra m�s
    Texto = CApos
    i = InStr(1, Texto, "�")
    If i = 0 Then
        CApos = Texto
    Else
        CApos = Mid(Texto, 1, i - 1) & "�" & Mid(Texto, i + 1, Len(Texto) - i)
    End If
    '-- Seguimos con transformaciones
    Texto = CApos
    i = InStr(1, Texto, "�")
    If i = 0 Then
        CApos = Texto
    Else
        CApos = Mid(Texto, 1, i - 1) & "�" & Mid(Texto, i + 1, Len(Texto) - i)
    End If
End Function



Public Function DevuelveValor(vSQL As String) As Variant
'Devuelve el valor de la SQL
Dim RS As ADODB.Recordset

    On Error Resume Next

    Set RS = New ADODB.Recordset
    RS.Open vSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    DevuelveValor = 0
    If Not RS.EOF Then
        ' antes RS.Fields(0).Value > 0
        If Not IsNull(RS.Fields(0).Value) Then DevuelveValor = RS.Fields(0).Value   'Solo es para saber que hay registros que mostrar
    End If
    RS.Close
    Set RS = Nothing

    If Err.Number <> 0 Then
        DevuelveValor = 0
        Err.Clear
    End If
End Function

Public Function TotalRegistrosConsulta(cadSQL) As Long
Dim Cad As String
Dim RS As ADODB.Recordset

    On Error GoTo ErrTotReg
    Cad = "SELECT count(*) FROM (" & cadSQL & ") x"
    Set RS = New ADODB.Recordset
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not RS.EOF Then
        TotalRegistrosConsulta = DBLet(RS.Fields(0).Value, "N")
    End If
    RS.Close
    Set RS = Nothing
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
Public Function TipoFacturarForfaits(Albaran As String, linea As String) As Byte
' devuelve 0: facturar por unidades
'          1: facturar por kilos
Dim RS As ADODB.Recordset
Dim SQL As String

    TipoFacturarForfaits = 2
    
    If Trim(Albaran) = "" Or Trim(linea) = "" Then Exit Function

    SQL = "select forfaits.facturar from albaran_variedad, forfaits "
    SQL = SQL & " where albaran_variedad.numalbar = " & DBSet(Albaran, "N")
    SQL = SQL & " and albaran_variedad.numlinea = " & DBSet(linea, "N")
    SQL = SQL & " and forfaits.codforfait = albaran_variedad.codforfait "
    SQL = SQL & " order by numlinea "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If Not RS.EOF Then
        TipoFacturarForfaits = DBLet(RS.Fields(0).Value, "N")
    End If
    
End Function


Public Function CalidadDestrio(Variedad As String) As String
Dim RS As ADODB.Recordset
Dim SQL As String

    CalidadDestrio = ""
    
    If Trim(Variedad) = "" Then Exit Function

    SQL = "select codcalid from rcalidad "
    SQL = SQL & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    SQL = SQL & " and rcalidad.tipcalid = 1" ' tipo de calidad de destrio
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        CalidadDestrio = DBLet(RS.Fields(0).Value, "N")
    End If
    
    Set RS = Nothing
    
End Function


Public Function CalidadDestrioenClasificacion(Variedad As String, Nota As String, Optional ConKilos As Boolean) As String
'conkilos = true --> miramos que el registro de esa clasificacion tenga kilos <> 0
Dim RS As ADODB.Recordset
Dim SQL As String

    CalidadDestrioenClasificacion = ""
    
    If Trim(Variedad) = "" Then Exit Function

    SQL = "select rcalidad.codcalid from rclasifica_clasif, rcalidad "
    SQL = SQL & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    SQL = SQL & " and rcalidad.tipcalid = 1" ' tipo de calidad de destrio
    SQL = SQL & " and rclasifica_clasif.numnotac = " & DBSet(Nota, "N")
    SQL = SQL & " and rclasifica_clasif.codvarie = rcalidad.codvarie "
    SQL = SQL & " and rclasifica_clasif.codcalid = rcalidad.codcalid "
    
    If ConKilos Then
        SQL = SQL & " and rclasifica_clasif.kilosnet <> 0"
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        CalidadDestrioenClasificacion = DBLet(RS.Fields(0).Value, "N")
    End If
    
    Set RS = Nothing
    
End Function

Public Function CalidadMaximaMuestraenClasificacion(Variedad As String, Nota As String, Optional ConKilos As Boolean) As String
'conkilos = true --> miramos que el registro de esa clasificacion tenga kilos <> 0
Dim RS As ADODB.Recordset
Dim SQL As String

    CalidadMaximaMuestraenClasificacion = ""
    
    If Trim(Variedad) = "" Then Exit Function

    SQL = "select rclasifica_clasif.codcalid from rclasifica_clasif "
    SQL = SQL & " where rclasifica_clasif.numnotac = " & DBSet(Nota, "N")
    
    If ConKilos Then
        SQL = SQL & " and rclasifica_clasif.kilosnet <> 0"
    End If
    
    SQL = SQL & " and muestra = (select max(rclasifica_clasif.muestra) from rclasifica_clasif "
    SQL = SQL & " where rclasifica_clasif.numnotac = " & DBSet(Nota, "N") & ")"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        CalidadMaximaMuestraenClasificacion = DBLet(RS.Fields(0).Value, "N")
    End If
    
    Set RS = Nothing
    
End Function




Public Function HorasDecimal(Cantidad As String) As Currency
Dim Entero As Long
Dim vCantidad As String
Dim vDecimal As String
Dim vEntero As String
Dim vHoras As Currency
Dim J As Integer
    HorasDecimal = 0
    
    vCantidad = ImporteSinFormato(Cantidad)
    
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


Public Function DecimalHoras(Cantidad As Currency) As Currency
Dim Entero As Long
Dim vCantidad As String
Dim vDecimal As String
Dim vEntero As String
Dim vHoras As Currency
Dim J As Integer
    
    DecimalHoras = 0
    
    vCantidad = ImporteSinFormato(CStr(Cantidad))
    
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



Public Function Horas(Cantidad As String) As Currency
Dim Entero As Long
Dim vCantidad As String
Dim vDecimal As String
Dim vEntero As String
Dim vHoras As Currency
Dim J As Integer

    Horas = 0
    
    vCantidad = ImporteSinFormato(Cantidad)
    
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




Public Function ComprobacionRangoFechas(Varie As String, tipo As String, Contador As String, fecha1 As String, fecha2 As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim b As Boolean

    On Error GoTo eComprobacionRangoFechas
    
    ComprobacionRangoFechas = False
    
    SQL = "select rprecios.fechaini, rprecios.fechafin, max(contador) from rprecios "
    SQL = SQL & " where codvarie = " & DBSet(Varie, "N")
    SQL = SQL & " and tipofact = " & DBSet(tipo, "N")
    
    If Contador <> "" Then
        SQL = SQL & " and contador <> " & DBSet(Contador, "N")
    End If
    SQL = SQL & " group by 1,2 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    b = False
    While Not RS.EOF And Not b
        '[Monica]20/01/2014: a�adido el if de que si coinciden no hacer nada
        If fecha1 = DBLet(RS.Fields(0).Value, "F") And fecha2 = DBLet(RS.Fields(1).Value, "F") Then
            ComprobacionRangoFechas = True
            Exit Function
        Else
            b = EntreFechas(fecha1, DBLet(RS.Fields(0).Value, "F"), fecha2)
            If Not b Then b = EntreFechas(fecha1, DBLet(RS.Fields(1).Value, "F"), fecha2)
            RS.MoveNext
        End If
    Wend
    
    RS.Close
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not RS.EOF And Not b
        '[Monica]20/01/2014: a�adido el if de que si coinciden no hacer nada
        If fecha1 = DBLet(RS.Fields(0).Value, "F") And fecha2 = DBLet(RS.Fields(1).Value, "F") Then
            ComprobacionRangoFechas = True
            Exit Function
        Else
            b = EntreFechas(DBLet(RS.Fields(0).Value, "F"), fecha1, DBLet(RS.Fields(1).Value, "F"))
            If Not b Then b = EntreFechas(DBLet(RS.Fields(0).Value, "F"), fecha2, DBLet(RS.Fields(1).Value, "F"))
        End If
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    ComprobacionRangoFechas = Not b
    
    Exit Function
    
eComprobacionRangoFechas:
    MuestraError Err.Number, "Comprobacion de rango fechas", Err.Description
End Function

Public Function PartidaCampo(codcampo As String) As String
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error Resume Next

    PartidaCampo = ""
    
    SQL = "select nomparti from rpartida, rcampos where rcampos.codcampo = " & DBSet(codcampo, "N")
    SQL = SQL & " and rcampos.codparti = rpartida.codparti "
    
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        PartidaCampo = DBLet(RS.Fields(0).Value, "T")
    End If
    
    Set RS = Nothing
    
End Function

Public Function HayRegistros(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim RS As ADODB.Recordset

    SQL = "Select * FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RS.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function

Public Function RellenaAceros(cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim Cad As String
    
    Cad = Mid("00000000000000000000", 1, longitud)
    If PorLaDerecha Then
        Cad = cadena & Cad
        RellenaAceros = Left(Cad, longitud)
    Else
        Cad = Cad & cadena
        RellenaAceros = Right(Cad, longitud)
    End If
    
End Function

Public Function RellenaABlancos(cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim Cad As String
    
    Cad = Space(longitud)
    If PorLaDerecha Then
        Cad = cadena & Cad
        RellenaABlancos = Left(Cad, longitud)
    Else
        Cad = Cad & cadena
        RellenaABlancos = Right(Cad, longitud)
    End If
    
End Function


Public Function EstaSocioDeAlta(Socio As String) As Boolean
Dim SQL As String

    SQL = "select count(*) from rsocios where codsocio = " & DBSet(Socio, "N")
    SQL = SQL & " and fechabaja is null"
    
    EstaSocioDeAlta = (TotalRegistros(SQL) > 0)

End Function

Public Function EstaCampoDeAlta(campo As String) As Boolean
Dim SQL As String

    SQL = "select count(*) from rcampos where codcampo = " & DBSet(campo, "N")
    SQL = SQL & " and fecbajas is null"
    
    EstaCampoDeAlta = (TotalRegistros(SQL) > 0)

End Function

Public Function EstaSocioDeAltaSeccion(Socio As String, Secc As String) As Boolean
Dim SQL As String

    SQL = "select count(*) from rsocios_seccion where codsocio = " & DBSet(Socio, "N")
    SQL = SQL & " and codsecci = " & DBSet(Secc, "N")
    SQL = SQL & " and fecbaja is null"
    
    EstaSocioDeAltaSeccion = (TotalRegistros(SQL) > 0)

End Function

Public Function EsSocioDeSeccion(Socio As String, Secc As String) As Boolean
Dim SQL As String

    SQL = "select count(*) from rsocios_seccion where codsocio = " & DBSet(Socio, "N")
    SQL = SQL & " and codsecci = " & DBSet(Secc, "N")
    
    EsSocioDeSeccion = (TotalRegistros(SQL) > 0)

End Function


Public Function CalidadVentaCampo(Variedad As String) As String
Dim RS As ADODB.Recordset
Dim SQL As String

    CalidadVentaCampo = ""
    
    If Trim(Variedad) = "" Then Exit Function

    SQL = "select codcalid from rcalidad "
    SQL = SQL & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    SQL = SQL & " and rcalidad.tipcalid = 2" ' tipo de calidad de venta campo
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        CalidadVentaCampo = DBLet(RS.Fields(0).Value, "N")
    End If
    
    Set RS = Nothing
    
End Function

Public Function CalidadRetirada(Variedad As String) As String
Dim RS As ADODB.Recordset
Dim SQL As String

    CalidadRetirada = ""
    
    If Trim(Variedad) = "" Then Exit Function

    SQL = "select codcalid from rcalidad "
    SQL = SQL & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    SQL = SQL & " and rcalidad.tipcalid1 = 2" ' tipo de calidad de retirada
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        CalidadRetirada = DBLet(RS.Fields(0).Value, "N")
    End If
    
    Set RS = Nothing
    
End Function





Public Function CalidadPrimera(Variedad As String) As String
Dim RS As ADODB.Recordset
Dim SQL As String

    CalidadPrimera = ""
    
    If Trim(Variedad) = "" Then Exit Function

    SQL = "select min(codcalid) from rcalidad "
    SQL = SQL & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        CalidadPrimera = DBLet(RS.Fields(0).Value, "N")
    End If
    
    Set RS = Nothing
    
End Function




Public Function EsCampoSocioVariedad(campo As String, Socio As String, Variedad As String) As Boolean
Dim SQL As String
Dim Sql2 As String


    EsCampoSocioVariedad = True
    
    If campo = "" Or Socio = "" Or Variedad = "" Then Exit Function
    
    SQL = "select count(*) from rcampos where codcampo = " & DBSet(campo, "N")
    SQL = SQL & " and codsocio = " & DBSet(Socio, "N")
    SQL = SQL & " and codvarie = " & DBSet(Variedad, "N")
    
    Sql2 = "select count(*) from rcampos INNER JOIN  rcampos_cooprop  ON rcampos.codcampo = rcampos_cooprop.codcampo and rcampos.codcampo = " & DBSet(campo, "N")
    Sql2 = Sql2 & " and rcampos_cooprop.codsocio = " & DBSet(Socio, "N")
    Sql2 = Sql2 & " and rcampos.codvarie = " & DBSet(Variedad, "N")
    
    
    EsCampoSocioVariedad = (TotalRegistros(SQL) > 0) Or (TotalRegistros(Sql2) > 0)

End Function

Public Function EsCampoSocio(campo As String, Socio As String) As Boolean
Dim SQL As String
Dim Sql2 As String


    EsCampoSocio = True
    
    If campo = "" Or Socio = "" Then Exit Function
    
    SQL = "select count(*) from rcampos where codcampo = " & DBSet(campo, "N")
    SQL = SQL & " and codsocio = " & DBSet(Socio, "N")
    
    EsCampoSocio = (TotalRegistros(SQL) > 0) Or (TotalRegistros(Sql2) > 0)

End Function


Public Function ContinuarSiAlbaranImpreso(Albaran As String) As Boolean
Dim SQL As String

    ContinuarSiAlbaranImpreso = True
    SQL = "select impreso from rhisfruta where numalbar = " & DBSet(Albaran, "N")
    If DevuelveValor(SQL) = 1 Then
        If MsgBox("Este Albar�n ya ha sido impreso. � Desea Continuar ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            ContinuarSiAlbaranImpreso = False
        End If
    End If

End Function


Public Function ExisteNota(Nota As String) As Boolean
Dim SQL As String
Dim Total As Integer

    ExisteNota = False
    
    SQL = "select count(*) from rentradas where numnotac = " & DBSet(Nota, "N")
    Total = TotalRegistros(SQL)
    If Total = 0 Then
        SQL = "select count(*) from rclasifica where numnotac = " & DBSet(Nota, "N")
        Total = TotalRegistros(SQL)
        If Total = 0 Then
            SQL = "select count(*) from rhisfruta_entradas where numnotac = " & DBSet(Nota, "N")
            Total = TotalRegistros(SQL)
            ExisteNota = (Total <> 0)
        Else
            ExisteNota = True
        End If
    Else
        ExisteNota = True
    End If
    
    
End Function

Public Sub AyudaFamiliasCom(frmCom As frmComercial, Optional CodActual As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|600|;S|txtAux(1)|T|Descripci�n|4000|;"
    frmCom.CadenaConsulta = "SELECT sfamia.codfamia, sfamia.nomfamia "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sfamia "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    frmCom.Tag1 = "C�digo|N|N|0|999|sfamia|codfamia|000|S|"
    frmCom.Tag2 = "Descripci�n|T|N|||sfamia|nomfamia|||"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 25
    frmCom.Caption = "Familias de Comercial"
    frmCom.Orden = " ORDER BY sfamia.codfamia"
    frmCom.CadLimpia = "sfamia.codfamia = -1"
    frmCom.CadenaSituar = "codfamia="
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    frmCom.Show vbModal
End Sub


Public Sub AyudaFamiliasADV(frmCom As frmComercial, Optional CodActual As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|600|;S|txtAux(1)|T|Descripci�n|4000|;"
    frmCom.CadenaConsulta = "SELECT advfamia.codfamia, advfamia.nomfamia "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM advfamia "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    frmCom.Tag1 = "C�digo|N|N|0|9999|advfamia|codfamia|0000|S|"
    frmCom.Tag2 = "Descripci�n|T|N|||advfamia|nomfamia|||"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 25
    frmCom.Caption = "Familias de Art�culos ADV"
    frmCom.Orden = " ORDER BY advfamia.codfamia"
    frmCom.CadLimpia = "advfamia.codfamia = -1"
    frmCom.CadenaSituar = "codfamia="
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    frmCom.Show vbModal
End Sub



Public Sub AyudaTUnidadesCom(frmCom As frmComercial, Optional CodActual As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|400|;S|txtAux(1)|T|Descripci�n|4200|;"
    frmCom.CadenaConsulta = "SELECT sunida.codunida, sunida.nomunida "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sunida "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    frmCom.Tag1 = "C�digo|N|N|0|99|sunida|codunida|00|S|"
    frmCom.Tag2 = "Descripci�n|T|N|||sunida|nomunida|||"
    frmCom.Maxlen1 = 2
    frmCom.Maxlen2 = 10
    frmCom.Caption = "Tipos de Unidad de Comercial"
    frmCom.Orden = " ORDER BY sunida.codunida"
    frmCom.CadLimpia = "sunida.codunida = -1"
    frmCom.CadenaSituar = "codunida="
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    frmCom.Show vbModal
End Sub

Public Sub AyudaProveedoresCom(frmCom As frmComercial, Optional CodActual As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|700|;S|txtAux(1)|T|Nombre|3900|;"
    frmCom.CadenaConsulta = "SELECT proveedor.codprove, proveedor.nomprove "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM proveedor "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    frmCom.Tag1 = "C�digo|N|N|0|999999|proveedor|codprove|000000|S|"
    frmCom.Tag2 = "Nombre|T|N|||proveedor|nomprove|||"
    frmCom.Maxlen1 = 6
    frmCom.Maxlen2 = 40
    frmCom.Caption = "Proveedores de Comercial"
    frmCom.Orden = " ORDER BY proveedor.codprove"
    frmCom.CadLimpia = "proveedor.codprove = -1"
    frmCom.CadenaSituar = "codprove="
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    frmCom.Show vbModal
End Sub

Public Sub AyudaProductosCom(frmCom As frmComercial, Optional CodActual As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|700|;S|txtAux(1)|T|Nombre|3900|;"
    frmCom.CadenaConsulta = "SELECT productos.codprodu, productos.nomprodu "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM productos "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    frmCom.Tag1 = "C�digo|N|N|0|999|productos|codprodu|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||productos|nomprodu|||"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Caption = "Productos de Comercial"
    frmCom.Orden = " ORDER BY productos.codprodu"
    frmCom.CadLimpia = "productos.codprodu = -1"
    frmCom.CadenaSituar = "codprodu="
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaForfaitsCom(frmCom As frmComercial, Optional CodActual As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|1500|;S|txtAux(1)|T|Nombre|2900|;"
    frmCom.CadenaConsulta = "SELECT forfaits.codforfait, forfaits.nomconfe "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM forfaits "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    frmCom.Tag1 = "C�digo|T|N|||forfaits|codforfait||S|"
    frmCom.Tag2 = "Nombre|T|N|||forfaits|nomconfe|||"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Caption = "Forfaits de Comercial"
    frmCom.Orden = " ORDER BY forfaits.codforfait"
    frmCom.CadLimpia = "forfaits.codforfait = ''"
    frmCom.CadenaSituar = "codforfait="
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaFPagoCom(frmCom As frmComercial, Optional CodActual As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|700|;S|txtAux(1)|T|Nombre|3900|;"
    frmCom.CadenaConsulta = "SELECT forpago.codforpa, forpago.nomforpa "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM forpago "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    frmCom.Tag1 = "C�digo|N|N|0|999|forpago|codforpa|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||forpago|nomforpa|||"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Caption = "Formas Pago de Comercial"
    frmCom.Orden = " ORDER BY forpago.codforpa"
    frmCom.CadLimpia = "forpago.codforpa = -1"
    frmCom.CadenaSituar = "codforpa="
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    frmCom.Show vbModal
End Sub



Public Sub AyudaAlmacenCom(frmCom As frmComercial, Optional CodActual As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|700|;S|txtAux(1)|T|Nombre|3900|;"
    frmCom.CadenaConsulta = "SELECT salmpr.codalmac, salmpr.nomalmac "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM salmpr "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    frmCom.Tag1 = "C�digo|N|N|0|999|salmpr|codalmac|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||salmpr|nomalmac|||"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Caption = "Almacenes Propios de Comercial"
    frmCom.Orden = " ORDER BY salmpr.codalmac"
    frmCom.CadLimpia = "salmpr.codalmac = -1"
    frmCom.CadenaSituar = "codalmac="
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaBancosCom(frmCom As frmComercial, Optional CodActual As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|700|;S|txtAux(1)|T|Nombre|3900|;"
    frmCom.CadenaConsulta = "SELECT banpropi.codbanpr, banpropi.nombanpr "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM banpropi "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    frmCom.Tag1 = "C�digo|N|N|0|999|banpropi|codbanpr|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||banpropi|nombanpr|||"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Caption = "Bancos Propios de Comercial"
    frmCom.Orden = " ORDER BY banpropi.codbanpr"
    frmCom.CadLimpia = "banpropi.codbanpr = -1"
    frmCom.CadenaSituar = "codbanpr="
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    frmCom.Show vbModal
End Sub




Public Sub AyudaClasesCom(frmCom As frmComercial, Optional CodActual As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|700|;S|txtAux(1)|T|Nombre|3900|;"
    frmCom.CadenaConsulta = "SELECT clases.codclase, clases.nomclase "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM clases "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    frmCom.Tag1 = "C�digo|N|N|0|999|clases|codclase|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||clases|nomclase|||"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Caption = "Clases de Comercial"
    frmCom.Orden = " ORDER BY clases.codclase"
    frmCom.CadLimpia = "clases.codclase = -1"
    frmCom.CadenaSituar = "codclase="
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    frmCom.Show vbModal
End Sub

Public Sub AyudaGrupoCom(frmCom As frmComercial, Optional CodActual As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|700|;S|txtAux(1)|T|Nombre|3900|;"
    frmCom.CadenaConsulta = "SELECT grupopro.codgrupo, grupopro.nomgrupo "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM grupopro "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    frmCom.Tag1 = "C�digo|N|N|0|999|grupopro|codgrupo|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||grupopro|nomgrupo|||"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Caption = "Grupos de Producto de Comercial"
    frmCom.Orden = " ORDER BY grupopro.codgrupo"
    frmCom.CadLimpia = "grupopro.codgrupo = -1"
    frmCom.CadenaSituar = "codgrupo="
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaHorarioCom(frmCom As frmComercial, Optional CodActual As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|C�digo|700|;S|txtAux(1)|T|Nombre|3900|;"
    frmCom.CadenaConsulta = "SELECT cchorario.codhorario, cchorario.descripc "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM cchorario "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    frmCom.Tag1 = "C�digo|N|N|0|999|cchorario|codhorario|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||cchorario|descripc|||"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Caption = "Horarios Costes de Comercial"
    frmCom.Orden = " ORDER BY cchorario.codhorario"
    frmCom.CadLimpia = "cchorario.codhorario = -1"
    frmCom.CadenaSituar = "cchorario="
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    frmCom.Show vbModal
End Sub




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
                        
                        'Si es articulo DE VARIOS podemos modificar la descripci�n del articulo, sino bloqueamos.
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
Dim SQL As String

    SQL = "select count(*) from variedades inner join productos on variedades.codprodu = productos.codprodu "
    SQL = SQL & " and variedades.codvarie = " & DBSet(Variedad, "N")
    SQL = SQL & " and productos.codgrupo = 6 "
    
    EsVariedadGrupo6 = (TotalRegistros(SQL) > 0)

End Function

' grupo 5 es del grupo de almazara (olivos)
Public Function EsVariedadGrupo5(Variedad As String) As Boolean
Dim SQL As String

    SQL = "select count(*) from variedades inner join productos on variedades.codprodu = productos.codprodu "
    SQL = SQL & " and variedades.codvarie = " & DBSet(Variedad, "N")
    SQL = SQL & " and productos.codgrupo = 5 "
    
    EsVariedadGrupo5 = (TotalRegistros(SQL) > 0)

End Function


'Para ello le decimos el orden  y ya ta
Public Function NumeroSubcadenasInStr(ByRef cadena As String, Separador As String) As Long
Dim i As Integer
Dim J As Integer
Dim cont As Integer
Dim Cad As String

    cont = 0
    If Len(cadena) <> 0 Then
        For i = 1 To Len(cadena)
            If Mid(cadena, i, 1) = Separador Then
                cont = cont + 1
            End If
        Next i
    End If
    
    NumeroSubcadenasInStr = cont

End Function


Public Function CalidadMenut(Variedad As String) As String
Dim RS As ADODB.Recordset
Dim SQL As String

    CalidadMenut = ""
    
    If Trim(Variedad) = "" Then Exit Function

    SQL = "select codcalid from rcalidad "
    SQL = SQL & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    SQL = SQL & " and rcalidad.tipcalid = 4" ' tipo de calidad de menut
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        CalidadMenut = DBLet(RS.Fields(0).Value, "N")
    End If
    
    Set RS = Nothing
    
End Function



Public Function DevuelveNomCalidad(Cad As String, PosIni As Integer) As String
Dim i As Integer
Dim Encontrado As Boolean

    DevuelveNomCalidad = ""
    
    i = PosIni
    Encontrado = False
    
    While i < Len(Cad) And Not Encontrado
        If Mid(Cad, i, 1) = """" Then Encontrado = True
        
        i = i + 1
    Wend

    DevuelveNomCalidad = Mid(Cad, PosIni, i - PosIni - 1)

End Function

Public Function ComprobacionRangoGrado(Varie As String, Desde As String, Hasta As String, linea As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim b As Boolean

    On Error GoTo eComprobacionRangoFechas
    
    ComprobacionRangoGrado = False
    
    SQL = "select rbonifica_lineas.desdegrado, rbonifica_lineas.hastagrado from rbonifica_lineas "
    SQL = SQL & " where codvarie = " & DBSet(Varie, "N")
    
    If linea <> "" Then
        SQL = SQL & " and numlinea <> " & DBSet(linea, "N")
    End If
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    b = False
    While Not RS.EOF And Not b
        b = ((CCur(Desde) <= DBLet(RS.Fields(0).Value, "N")) And (DBLet(RS.Fields(0).Value, "N") <= CCur(Hasta)))
        If Not b Then b = ((CCur(Desde) <= DBLet(RS.Fields(1).Value, "N")) And (DBLet(RS.Fields(1).Value, "N") <= CCur(Hasta)))
        RS.MoveNext
    Wend
    
    RS.Close
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not RS.EOF And Not b
        b = ((DBLet(RS.Fields(0).Value, "N") <= CCur(Desde)) And (CCur(Desde) <= DBLet(RS.Fields(1).Value, "N")))
        If Not b Then b = ((DBLet(RS.Fields(0).Value, "N") <= CCur(Hasta)) And (CCur(Hasta) <= DBLet(RS.Fields(1).Value, "N")))
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    ComprobacionRangoGrado = Not b
    
    Exit Function
    
eComprobacionRangoFechas:
    MuestraError Err.Number, "Comprobacion de rango grado", Err.Description
End Function



Public Function EsGastodeFactura(Codigo As String) As Boolean
Dim SQL As String

    EsGastodeFactura = False
    
    SQL = "select tipogasto from rconcepgasto where codgasto = " & DBSet(Codigo, "N")
        
    EsGastodeFactura = (DevuelveValor(SQL) = 1)
    
End Function



Public Function HayEntradasCampoSocioVariedad(campo As String, Socio As String, Variedad As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim RS As ADODB.Recordset

    HayEntradasCampoSocioVariedad = True

    SQL = "Select count(*) FROM rentradas where codcampo = " & DBSet(campo, "N")
    SQL = SQL & " and codsocio = " & DBSet(Socio, "N")
    SQL = SQL & " and codvarie = " & DBSet(Variedad, "N")
    
    If TotalRegistros(SQL) = 0 Then
        SQL = "select count(*) from rclasifica where codcampo = " & DBSet(campo, "N")
        SQL = SQL & " and codsocio = " & DBSet(Socio, "N")
        SQL = SQL & " and codvarie = " & DBSet(Variedad, "N")
        
        If TotalRegistros(SQL) = 0 Then
            SQL = "select count(*) from rhisfruta where codcampo = " & DBSet(campo, "N")
            SQL = SQL & " and codsocio = " & DBSet(Socio, "N")
            SQL = SQL & " and codvarie = " & DBSet(Variedad, "N")
        
            HayEntradasCampoSocioVariedad = Not (TotalRegistros(SQL) = 0)
        End If
    End If
    
End Function

Public Function HayEntradasSocio(Socio As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim RS As ADODB.Recordset

    HayEntradasSocio = True

    SQL = "Select count(*) FROM rentradas where codsocio = " & DBSet(Socio, "N")
    
    If TotalRegistros(SQL) = 0 Then
        SQL = "select count(*) from rclasifica where codsocio = " & DBSet(Socio, "N")
        
        If TotalRegistros(SQL) = 0 Then
            SQL = "select count(*) from rhisfruta where codsocio = " & DBSet(Socio, "N")
        
            HayEntradasSocio = Not (TotalRegistros(SQL) = 0)
        End If
    End If
    
End Function


Public Function HayAnticiposPdtesCampoSocioVariedad(campo As String, Socio As String, Variedad As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim RS As ADODB.Recordset

    HayAnticiposPdtesCampoSocioVariedad = True

    SQL = "Select count(*) FROM rfactsoc_variedad, rfactsoc, usuarios.stipom stipom where rfactsoc_variedad.codcampo = " & DBSet(campo, "N")
    SQL = SQL & " and rfactsoc.codsocio = " & DBSet(Socio, "N")
    SQL = SQL & " and rfactsoc_variedad.codvarie = " & DBSet(Variedad, "N")
    SQL = SQL & " and rfactsoc_variedad.descontado = 0 "
    SQL = SQL & " and rfactsoc.codtipom = rfactsoc_variedad.codtipom "
    SQL = SQL & " and rfactsoc.numfactu = rfactsoc_variedad.numfactu "
    SQL = SQL & " and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu "
    SQL = SQL & " and rfactsoc.codtipom = stipom.codtipom "
    SQL = SQL & " and stipom.tipodocu = 1 "
    
    HayAnticiposPdtesCampoSocioVariedad = (TotalRegistros(SQL) <> 0)
    
End Function



Public Function ModificarEntradas(campo As String, SocAnt As String, VarAnt As String, SocNue As String, VarNue As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim RS As ADODB.Recordset

    On Error GoTo eModificarEntradas

    ModificarEntradas = False

    SQL = "update rentradas "
    SQL = SQL & " set codsocio = " & DBSet(SocNue, "N")
    SQL = SQL & ", codvarie = " & DBSet(VarNue, "N")
    SQL = SQL & "  where codcampo = " & DBSet(campo, "N")
    SQL = SQL & " and codsocio = " & DBSet(SocAnt, "N")
    SQL = SQL & " and codvarie = " & DBSet(VarAnt, "N")
    
    conn.Execute SQL
    
    SQL = "update rclasifica "
    SQL = SQL & " set codsocio = " & DBSet(SocNue, "N")
    SQL = SQL & ", codvarie = " & DBSet(VarNue, "N")
    SQL = SQL & "  where codcampo = " & DBSet(campo, "N")
    SQL = SQL & " and codsocio = " & DBSet(SocAnt, "N")
    SQL = SQL & " and codvarie = " & DBSet(VarAnt, "N")
    
    conn.Execute SQL
        
    SQL = "update rhisfruta "
    SQL = SQL & " set codsocio = " & DBSet(SocNue, "N")
    SQL = SQL & ", codvarie = " & DBSet(VarNue, "N")
    SQL = SQL & "  where codcampo = " & DBSet(campo, "N")
    SQL = SQL & " and codsocio = " & DBSet(SocAnt, "N")
    SQL = SQL & " and codvarie = " & DBSet(VarAnt, "N")
    
    conn.Execute SQL
    
    ModificarEntradas = True
    Exit Function
    
eModificarEntradas:
    MuestraError Err.Number, "Modificar Entradas", Err.Description
End Function


Public Sub PonerContRegIndicador(ByRef lblIndicador As Label, ByRef vData As Adodc, cadBuscar As String)
'cuando esta en el MODO 2 pone el label de contador de registros a�adiendo
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


Public Sub CargaCadenaAyuda(ByRef vCadena As String, tipo As Integer)
    Select Case tipo
        Case 0
             
           ' "____________________________________________________________"
           '           1234567890123456789012345678901234567890123456789012345678901234567
             vCadena = "Tipo de Entradas" & vbCrLf & _
                       "=============" & vbCrLf & vbCrLf
            
             vCadena = vCadena & "Entrada de Retirada: " & vbCrLf & _
                       " Cuando se actualiza la entrada de b�scula, el porcentaje de destrio de " & vbCrLf & _
                       " la variedad va a la calidad de destrio y el resto a la calidad de retirada." & vbCrLf & _
                       "  " & vbCrLf & vbCrLf & _
                       "Entrada de Venta Comercio:" & vbCrLf & _
                       " Tiene el mismo tratamiento que las entradas normales, �nicamente es " & vbCrLf & _
                       " para diferenciarlas.  " & vbCrLf & vbCrLf
    
    End Select
End Sub




Public Function EsCalidadDestrio(Variedad As String, Calidad As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String

    EsCalidadDestrio = False
    
    If Trim(Variedad) = "" Or Trim(Calidad) = "" Then Exit Function

    SQL = "select tipcalid from rcalidad "
    SQL = SQL & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    SQL = SQL & " and rcalidad.codcalid = " & DBSet(Calidad, "N")
    
    EsCalidadDestrio = (DevuelveValor(SQL) = 1)
    
End Function


Public Function EsCalidadMerma(Variedad As String, Calidad As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String

    EsCalidadMerma = False
    
    If Trim(Variedad) = "" Or Trim(Calidad) = "" Then Exit Function

    SQL = "select tipcalid from rcalidad "
    SQL = SQL & " where rcalidad.codvarie = " & DBSet(Variedad, "N")
    SQL = SQL & " and rcalidad.codcalid = " & DBSet(Calidad, "N")
    
    EsCalidadMerma = (DevuelveValor(SQL) = 3)
    
End Function


Public Sub AyudaGlobalGap(frmBas As frmBasico, Optional CodActual As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|800|;S|txtAux(1)|T|Descripci�n|3930|;"
    frmBas.CadenaConsulta = "SELECT rglobalgap.codigo, rglobalgap.descripcion "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM rglobalgap "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "C�digo|T|N|||rglobalgap|codigo||S|"
    frmBas.Tag2 = "Descripci�n|T|N|||rglobalgap|descripcion|||"
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 40
    frmBas.Tabla = "rglobalgap"
    frmBas.CampoCP = "codigo"
    frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "GlobalGap"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaClienteAriges(frmBas As frmBasico, Optional CodActual As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|800|;S|txtAux(1)|T|Descripci�n|3930|;"
    frmBas.CadenaConsulta = "SELECT sclien.codclien, sclien.nomclien "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM " & vParamAplic.BDAriges & ".sclien "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "C�digo|N|N|||sclien|codclien|000000|S|"
    frmBas.Tag2 = "Nombre|T|N|||sclien|nomclien|||"
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 40
    frmBas.Tabla = vParamAplic.BDAriges & ".sclien"
    frmBas.CampoCP = "codclien"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Clientes de Suministros"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaClienteCom(frmBas As frmBasico, Optional CodActual As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|800|;S|txtAux(1)|T|Descripci�n|3930|;"
    frmBas.CadenaConsulta = "SELECT clientes.codclien, clientes.nomclien "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM clientes "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "C�digo|N|N|||clientes|codclien|000000|S|"
    frmBas.Tag2 = "Nombre|T|N|||clientes|nomclien|||"
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 40
    frmBas.Tabla = "clientes"
    frmBas.CampoCP = "codclien"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Clientes de Comercial"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub


Public Function ActualizarTraza(Nota As String, Variedad As String, Socio As String, campo As String, Fecha As String, Hora As String, MenError As String)
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim IdPalet As Currency

    On Error GoTo eActualizarTraza

    ActualizarTraza = True

    If vParamAplic.HayTraza = False Then Exit Function
    
    SQL = "select idpalet from trzpalets where numnotac = " & DBSet(Nota, "N")
    
    
    'Comprobamos si la fecha de abocamiento de alguno de sus palets es inferior a la de la entrada
    'para no permitir modificar la traza
    Sql2 = "select sum(resul) from (select if(fechahora<" & DBSet(Hora, "FH") & ",1,0) as resul "
    Sql2 = Sql2 & " from trzlineas_cargas where idpalet in (" & SQL & ")) aaaaa "
    If CLng(DevuelveValor(Sql2)) > 0 Then
        MenError = MenError & vbCrLf & "No se permite una fecha de entrada superior a la de abocamiento de ninguno de sus palets. Revise."
        ActualizarTraza = False
        Exit Function
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenDynamic, adLockReadOnly, adCmdText
    
    While Not RS.EOF
        
        Sql1 = "update trzpalets set codvarie = " & DBSet(Variedad, "N")
        Sql1 = Sql1 & ", codsocio = " & DBSet(Socio, "N")
        Sql1 = Sql1 & ", codcampo = " & DBSet(campo, "N")
        Sql1 = Sql1 & ", fecha = " & DBSet(Fecha, "F")
        Sql1 = Sql1 & ", hora = " & DBSet(Hora, "FH")
        Sql1 = Sql1 & " where idpalet = " & DBSet(RS.Fields(0).Value, "N")
        
        conn.Execute Sql1
        
        RS.MoveNext
    Wend
        
    Set RS = Nothing
    
    Exit Function
    
eActualizarTraza:
    ActualizarTraza = False
    MenError = MenError & vbCrLf & Err.Description
End Function



Public Sub ActualizarClasificacionHco(Albaran As String, Kilos As String)
Dim SQL As String
Dim Variedad As String
Dim Calidad As String
Dim KilosNet As Long

    Variedad = DevuelveValor("select codvarie from rhisfruta where numalbar = " & Albaran)

    If DevuelveValor("select count(*) from rcalidad where codvarie = " & Variedad) = 1 Then
        
        Calidad = DevuelveValor("select codcalid from rcalidad where codvarie = " & Variedad)
        
        If DevuelveValor("select count(*) from rhisfruta_clasif where numalbar = " & Albaran) = 0 Then
            SQL = "insert into rhisfruta_clasif (numalbar,codvarie,codcalid,kilosnet) values (" & DBSet(Albaran, "N") & ","
            SQL = SQL & DBSet(Variedad, "N") & "," & DBSet(Calidad, "N") & "," & DBSet(Kilos, "N") & ")"
            
            conn.Execute SQL
        Else
            KilosNet = DevuelveValor("select sum(kilosnet) from rhisfruta_entradas where numalbar = " & DBSet(Albaran, "N"))
            
            SQL = "update rhisfruta_clasif set kilosnet = " & DBSet(KilosNet, "N")
            SQL = SQL & " where numalbar = " & DBSet(Albaran, "N")
            
            conn.Execute SQL
        End If
    
    End If

End Sub


'Funcion de David del Ariges
Public Sub ComprobarCobrosSocio(Codsocio As String, FechaDoc As String, Optional DevuelveImporte As String)
'Comprueba en la tabla de Cobros Pendientes (scobro) de la Base de datos de Contabilidad
'si el cliente tiene alguna factura pendiente de cobro que ha vendido
'con fecha de vencimiento anterior a la fecha del documento: Oferta, Pedido, ALbaran,...
Dim SQL As String, vWhere As String
Dim Codmacta As String
Dim RS As ADODB.Recordset
Dim cadMen As String
Dim ImporteCred As Currency
Dim Importe As Currency
Dim ImpAux As Currency
Dim vSeccion As CSeccion


    Set RS = New ADODB.Recordset
    ImporteCred = 0
    
    '[Monica]15/05/2013: Solo para socios de la seccion de hortofruticolo
    If CInt(ComprobarCero(vParamAplic.Seccionhorto)) <> 0 Then
        
        'Obtener la cuenta del socio de la tabla rsocios en Ariagrorec
        SQL = "Select nomsocio, codmaccli from rsocios inner join rsocios_seccion on rsocios.codsocio = rsocios_seccion.codsocio "
        SQL = SQL & " where rsocios_seccion.codsecci = " & DBSet(vParamAplic.Seccionhorto, "N") & " and rsocios.codsocio=" & Codsocio
        RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If RS.EOF Then
            SQL = ""
        Else
            Codsocio = Codsocio & " - " & RS!nomsocio
            Codmacta = DBLet(RS!codmaccli)
            ImporteCred = 0
            If Codmacta = "" Then SQL = ""
        End If
        RS.Close
        If SQL = "" Then Exit Sub
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            If vSeccion.AbrirConta Then
        
                'AHORA FEBRERO 2010
                SQL = "SELECT scobro.* FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
                vWhere = " WHERE scobro.codmacta = '" & Codmacta & "'"
' lo llamamos desde el mto de socios, campos y contadores
'                vWhere = vWhere & " AND fecvenci <= ' " & Format(FechaDoc, FormatoFecha) & "' "
                'Antes mayo 2010
                'vWhere = vWhere & " AND (sforpa.tipforpa between 0 and 3)"
                vWhere = vWhere & " AND recedocu=0 "
                SQL = SQL & vWhere & " ORDER BY fecfaccl, codfaccl "
                
                'Lee de la Base de Datos de CONTABILIDAD
                RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
                Importe = 0
                While Not RS.EOF
                
                    'QUITO LO DE DEVUELTO. MAYO 2010
                    'If Val(RS!Devuelto) = 1 Then
                    '    'SALE SEGURO (si no esta girado otra vez �no?
                    '    'Si esta girado otra vez tendra impcobro, con lo cual NO tendra diferencia de importes
                    '    Impaux = RS!ImpVenci + DBLet(RS!gastos, "N") - DBLet(RS!impcobro, "N")
                        
                    'Else
                        'Si esta recibido NO lo saco
                        If Val(RS!recedocu) = 1 Then
                            ImpAux = 0
                        Else
                            'NO esta recibido. Si tiene diferencia
                            ImpAux = RS!ImpVenci + DBLet(RS!Gastos, "N") - DBLet(RS!impcobro, "N")
                    
                        End If
                '    End If
                    If ImpAux <> 0 Then Importe = Importe + ImpAux
                    RS.MoveNext
                Wend
                RS.Close
                Set RS = Nothing
                
                If Importe > 0 Then
                    If DevuelveImporte <> "" Then
                        'Meto aqui el importer
                        DevuelveImporte = CStr(Importe)
                    Else
                        cadMen = "El Socio tiene facturas vencidas con valor de: " & Format(Importe, FormatoImporte) & " �."
                        If ImporteCred > 0 Then cadMen = cadMen & vbCrLf & "L�mite cr�dito: " & Format(ImporteCred, FormatoImporte) & " �."
                        cadMen = cadMen & vbCrLf & "�Desea Ver Detalle?"
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




Public Sub AyudaIncidenciasOrdenesRecogida(frmBas As frmBasico, Optional CodActual As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|C�digo|800|;S|txtAux(1)|T|Descripci�n|3930|;"
    frmBas.CadenaConsulta = "SELECT rplagasaux.idplaga, rplagasaux.nomplaga "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM rplagasaux "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "C�digo|N|N|||rplagasaux|idplaga|000000|S|"
    frmBas.Tag2 = "Nombre|T|N|||rplagasaux|nomplaga|||"
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 40
    frmBas.Tabla = "rplagasaux"
    frmBas.CampoCP = "idplaga"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Incidencias Ordenes Recogida"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub



Public Function ObtenerPrecioRecoldeCalidad(Variedad As String, Calidad As String, tipo As Byte) As Currency
'Tipo = 0, factura transporte socio gastos de recoleccion
'Tipo = 1, gastos de recoleccion del recolector
Dim SQL As String

    SQL = "select "
    If tipo = 0 Then
        SQL = SQL & "eurrecsoc "
    Else
        SQL = SQL & "eurreccoop "
    End If
    SQL = SQL & " from rcalidad where codvarie = " & DBSet(Variedad, "N") & " and codcalid = " & DBSet(Calidad, "N")

    ObtenerPrecioRecoldeCalidad = DevuelveValor(SQL)

End Function


Public Function CodTipomAnticipos() As String
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Result As String

    CodTipomAnticipos = ""

    SQL = "select codtipom from usuarios.stipom where tipodocu = 1"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Result = ""
    
    While Not RS.EOF
        Result = Result & DBSet(RS!CodTipom, "T") & ","
        
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    
    If Result <> "" Then
        Result = Mid(Result, 1, Len(Result) - 1)
    End If

    CodTipomAnticipos = Result
    
End Function

Public Function EsCaja(CodEnvase As String) As Boolean
Dim SQL As String

    SQL = "select escaja from confenva where codtipen = " & DBSet(CodEnvase, "N")
    
    EsCaja = DevuelveValor(SQL)

End Function


' FUNCIONES DE POZOS PARA SABER SI ES O NO CONTADO (ESCALONA Y UTXERA)

Public Function EsSocioContadoPOZOS(Socio As String) As Boolean
Dim SQL As String

    SQL = "select cuentaba from rsocios where codsocio = " & DBSet(Socio, "N")
    EsSocioContadoPOZOS = (DevuelveValor(SQL) = "8888888888")

End Function


Public Function EsReciboContadoPOZOS(vWhere As String) As Boolean
Dim SQL As String

    SQL = "select escontado from rrecibpozos where " & vWhere
    EsReciboContadoPOZOS = (DevuelveValor(SQL) = "1")

End Function

Public Function EntregadaFichaCultivo(campo As String) As Boolean
Dim SQL As String

    SQL = "select entregafichaculti from rcampos where codcampo = " & DBSet(campo, "N")

    EntregadaFichaCultivo = (DevuelveValor(SQL) = "0")
    
End Function


Public Function TrabajadorDeBaja(Traba As String, Optional Fecha As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset

    SQL = "select fechabaja from straba where codtraba = " & DBSet(Traba, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        TrabajadorDeBaja = (DBLet(RS!FechaBaja, "F") <> "")
        If Fecha <> "" Then
            TrabajadorDeBaja = (DBLet(RS!FechaBaja, "F") >= Fecha)
        End If
    Else
        TrabajadorDeBaja = True
    End If
    Set RS = Nothing
    

End Function


Public Function HayXML() As Boolean
Dim SQL As String

    SQL = "select xml from rparam "
    HayXML = (DevuelveValor(SQL) = 1)


End Function

Public Function EsCalidadConBonificacion(Variedad As String, Calidad As String) As Boolean
Dim SQL As String

    SQL = "select seaplicabonif from rcalidad where codvarie = " & DBSet(Variedad, "N") & " and codcalid = " & DBSet(Calidad, "N")
    
    EsCalidadConBonificacion = (DevuelveValor(SQL) = 1)

End Function
