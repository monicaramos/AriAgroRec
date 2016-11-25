Attribute VB_Name = "ModFormularios"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'   FUNCIONES GENERALES
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------


'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Public Function SugerirCodigoSiguienteStr(NomTabla As String, NomCodigo As String, Optional CondLineas As String) As String
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo ESugerirCodigo

    'SQL = "Select Max(codtipar) from stipar"
    Sql = "Select Max(" & NomCodigo & ") from " & NomTabla
    If CondLineas <> "" Then
        Sql = Sql & " WHERE " & CondLineas
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, , , adCmdText
    Sql = "1"
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            If IsNumeric(Rs.Fields(0)) Then
                Sql = CStr(Rs.Fields(0) + 1)
            Else
                If Asc(Left(Rs.Fields(0), 1)) <> 122 Then 'Z
                Sql = Left(Rs.Fields(0), 1) & CStr(Asc(Right(Rs.Fields(0), 1)) + 1)
                End If
            End If
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    SugerirCodigoSiguienteStr = Sql
ESugerirCodigo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Public Sub BloquearFrameAux(ByRef formulario As Form, nom_frame As String, Modo As Byte, Optional NumTabMto As Integer)
Dim i As Byte
Dim b As Boolean
Dim Control As Object

    On Error GoTo EBloquear

    'b = (Modo = 3 Or Modo = 4 Or Modo = 5)
    b = (Modo = 5) 'And (NumTabMto = 3)
    
    For Each Control In formulario.Controls
        'If (Control.Tag <> "") And (Control.Visible = True) And (Control.Container.Name = nom_frame) Then
        If (Control.Tag <> "") Then
           If (Control.Container.Name = nom_frame) Then
                If (TypeOf Control Is TextBox) And (Control.Name = "txtAux") Then
                    Control.Locked = Not b
                    If b Then
                        Control.BackColor = vbWhite
                    Else
                        Control.BackColor = &H80000018 'Amarillo Claro
                    End If
                    If Modo = 3 Then Control.Text = "" 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                    
                ElseIf (TypeOf Control Is ComboBox) And (Control.Name = "cmbAux") Then
                    'Control.Locked = Not b
                    Control.Enabled = b
                    If b Then
                        Control.BackColor = vbWhite
                    Else
                        Control.BackColor = &H80000018 'Amarillo Claro
                    End If
                    If Modo = 3 Then Control.ListIndex = 0 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                End If
            End If
        End If
    
    Next Control

EBloquear:
    If Err.Number <> 0 Then Err.Clear
End Sub





Public Sub BloquearFrameAux2(ByRef formulario As Form, nom_frame As String, Bloquea As Boolean)
Dim b As Boolean
Dim Control As Object

    On Error GoTo EBloquear

    'b = (Modo = 3 Or Modo = 4 Or Modo = 5)
'    b = (Modo = 5) And (NumTabMto = 3)
    b = Bloquea
    
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) Then 'TEXT
            If (Control.Name = "txtAux") And (Control.Container.Name = nom_frame) Then
                If (Control.Tag <> "") Then
                    Control.Locked = b
                    If Not b Then
                        Control.BackColor = vbWhite
                    Else
                        Control.BackColor = &H80000018 'Amarillo Claro
                    End If
'                    If Modo = 3 Then Control.Text = "" 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                End If
            End If
            
        ElseIf (TypeOf Control Is ComboBox) Then 'COMBO
            If (Control.Name = "cmbAux") And (Control.Container.Name = nom_frame) Then
                Control.Enabled = Not b
                If Not b Then
                    Control.BackColor = vbWhite
                Else
                    Control.BackColor = &H80000018 'Amarillo Claro
                End If
'                If Modo = 3 Then Control.ListIndex = 0 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
            End If
        End If
    Next Control

EBloquear:
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub BloquearText1(ByRef formulario As Form, Modo As Byte)
'Bloquea controles q se llamen TEXT1 si no estamos en Modo: 3.-Insertar, 4.-Modificar
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualización
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar...)
Dim i As Byte
Dim b As Boolean
Dim vtag As CTag
On Error Resume Next

    With formulario
        'b = (Modo = 3 Or Modo = 4 Or Modo = 1 Or Modo = 5) 'And ModoLineas = 1))
        b = (Modo = 3 Or Modo = 4 Or Modo = 1) '06/09/2005, lleve el modo 5 per a que no es puga modificar la capçalera mentre treballe en les llínies
        
        For i = 0 To .Text1.Count - 1 'En principio todos los TExt1 tiene TAG
            Set vtag = New CTag
            vtag.Cargar .Text1(i)
            If vtag.Cargado Then
                If vtag.EsClave And (Modo = 4 Or Modo = 5) Then
                    .Text1(i).Locked = True
                    .Text1(i).BackColor = &H80000018 'groc
                Else
                     .Text1(i).Locked = Not b  '((Not b) And (Modo <> 1))
                    If b Then
                        .Text1(i).BackColor = vbWhite
                    Else
                        .Text1(i).BackColor = &H80000018 'groc
                    End If
                    If Modo = 3 Then .Text1(i).Text = "" 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                End If
'            Else
'                .text1(i).Locked = Not b  '((Not b) And (Modo <> 1))
'                If b Then
'                    .text1(i).BackColor = vbWhite
'                Else
'                    .text1(i).BackColor = &H80000018 'groc
'                End If
'                If Modo = 3 Then .text1(i).Text = "" 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
            End If
            Set vtag = Nothing
        Next i
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearTxt(ByRef Text As TextBox, b As Boolean, Optional EsContador As Boolean)
'Bloquea un control de tipo TextBox
'Si lo bloquea lo poner de color amarillo claro sino lo pone en color blanco (sino es contador)
'pero si es contador lo pone color azul claro
On Error Resume Next

    Text.Locked = b
    If Not b And Text.Enabled = False Then Text.Enabled = True
    If b Then
        If EsContador Then
            'Si Es un campo que se obtiene de un contador poner color azul
'            Text.BackColor = &H80000013 'Azul Claro
            Text.BackColor = &HFFFFC0   'Azul claro con vista
        Else
            Text.BackColor = &H80000018 'Amarillo Claro
        End If
    Else
        Text.BackColor = vbWhite
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearCmb(ByRef Cmb As ComboBox, b As Boolean, Optional EsContador As Boolean)
'Bloqueja un control de tipo ComboBox
'Si el bloqueja el posa de color gris claro, sino el posa de color blanc (sino es contador)
'pero si es contador el posa color blau clar
On Error Resume Next

    'Cmb.Locked = b
    Cmb.Enabled = Not b
    'If Not b And Cmb.Enabled = False Then Cmb.Enabled = True
    If b Then
        If EsContador Then
            'Si Es un campo que se obtiene de un contador poner color azul
            Cmb.BackColor = &H80000013 'Azul Claro
        Else
            Cmb.BackColor = &H80000018 'Amarillo Claro
        End If
    Else
        Cmb.BackColor = vbWhite
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearCheck1(ByRef formulario As Form, Modo As Byte)
'Bloquea controles q sean CheckBox si no estamos en Modo: 3.-Insertar, 4.-Modificar
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualización
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar...)
    Dim b As Boolean
'    Dim Control As Control

    On Error Resume Next

    b = (Modo = 3 Or Modo = 4 Or Modo = 1)
    With formulario
        For i = 0 To .Check1.Count - 1
            .Check1(i).Enabled = b
            If Modo = 3 Then .Check1(i).Value = 0 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
        Next i
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Sub




Public Sub BloquearChk(ByRef chk As CheckBox, b As Boolean)
'Bloquea un control de tipo CheckBox
'(IN) b : sera true o false segun si bloquea o no
    On Error Resume Next

    chk.Enabled = Not b
   
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearChecks(ByRef formulario As Form, Modo As Byte)
'Bloquea controles q sean CheckBox si no estamos en Modo: 3.-Insertar, 4.-Modificar
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualización
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar...)
Dim b As Boolean
Dim Control As Control
    
    On Error Resume Next

    b = (Modo = 3 Or Modo = 4 Or Modo = 1)
    
    With formulario
        For Each Control In formulario.Controls
            If TypeOf Control Is CheckBox Then
                If InStr(1, Control.Name, "Aux") Then
                
                Else
                    Control.Enabled = b
                    If Modo = 3 Then Control.Value = 0 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                End If
            End If
        Next Control
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub BloquearCombo(ByRef formulario As Form, Modo As Byte)
'Bloquea controles q sean ComboBox si no estamos en Modo: 3.-Insertar, 4.-Modificar
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualización
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar,...)
Dim b As Boolean
    
    On Error Resume Next

    'b = (Modo = 3 Or Modo = 4 Or Modo = 1 Or Modo = 5)
    b = (Modo = 3 Or Modo = 4 Or Modo = 1) '06/09/2005, lleve el modo 5 per a que no es puga modificar la capçalera mentre treballe en les llínies
    
    With formulario
        For i = 0 To .Combo1.Count - 1
            Set vtag = New CTag
            vtag.Cargar .Combo1(i)
            If vtag.Cargado Then
                If vtag.EsClave And (Modo = 4 Or Modo = 5) Then
                    .Combo1(i).Enabled = False
                    .Combo1(i).BackColor = &H80000018 'groc
                Else
                    .Combo1(i).Enabled = b
                    If b Then
                        .Combo1(i).BackColor = vbWhite
                    Else
                        .Combo1(i).BackColor = &H80000018 'Amarillo Claro
                    End If
                    If Modo = 3 Then .Combo1(i).ListIndex = 0 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                End If
            End If
        Next i
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub BloquearComboANTIC(ByRef formulario As Form, Modo As Byte)
''Bloquea controles q sean ComboBox si no estamos en Modo: 3.-Insertar, 4.-Modificar
''IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualización
''       Modo: modo del mantenimiento (Insertar, Modificar,Buscar,...)
'Dim b As Boolean
'On Error Resume Next
'
'    b = (Modo = 3 Or Modo = 4 Or Modo = 1)
'    With formulario
'        For i = 0 To .Combo1.Count - 1
'            .Combo1(i).Enabled = b
'            If b Then
'                .Combo1(i).BackColor = vbWhite
'            Else
'                .Combo1(i).BackColor = &H80000018 'Amarillo Claro
'            End If
'            If Modo = 3 Then .Combo1(i).ListIndex = 0 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
'        Next i
'    End With
'    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub BloquearImgBuscar(ByRef formulario As Form, Modo As Byte, Optional ModoLineas As Byte)
'Bloquea controles q sean ComboBox si no estamos en Modo: 3.-Insertar, 4.-Modificar
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualización
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar, Insertar/Modificar Lineas...)
Dim b As Boolean
On Error Resume Next

'    b = (Modo = 3 Or Modo = 4 Or Modo = 1 Or (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2)))
    b = (Modo = 3 Or Modo = 4 Or Modo = 1)
    
    With formulario
        For i = 0 To .imgBuscar.Count - 1
            .imgBuscar(i).Enabled = b
            .imgBuscar(i).visible = b
        Next i
    End With
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearImgBuscar2(ByRef formulario As Form, Modo As Byte, Optional ModoLineas As Byte)
'Bloquea controles q sean ComboBox si no estamos en Modo: 3.-Insertar, 4.-Modificar
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualización
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar, Insertar/Modificar Lineas...)
'En el TAG del ImgBuscar pongo un 1 si la imagen pertenece a  al tabla principal
'y un 0 si pertenece a los frame txtAux
Dim b As Boolean
'Dim bAux As Boolean
    
    On Error Resume Next

    b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Or (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2)))
'    bAux = (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2))
    
    With formulario
        For i = 0 To .imgBuscar.Count - 1
            If .imgBuscar(i).Tag = 1 Then 'esta en la cabecera
                .imgBuscar(i).Enabled = b
                .imgBuscar(i).visible = b
            Else 'esta en las lineas
                .imgBuscar(i).Enabled = False
                .imgBuscar(i).visible = False
            End If
        Next i
    End With
    If Err.Number <> 0 Then Err.Clear
End Sub






Public Sub BloquearImgZoom(ByRef formulario As Form, Modo As Byte, Optional ModoLineas As Byte)
'Bloquea los controles q sean Image zoom si no estamos en Modo: 3.-Insertar, 4.-Modificar
'(IN) -> formulario: formulario en el que se van a poner los controles Image zoom en modo visualización
'(IN) -> Modo: modo del mantenimiento (Insertar, Modificar,Buscar, Insertar/Modificar Lineas...)

    Dim b As Boolean

    On Error Resume Next

    b = (Modo = 3 Or Modo = 4 Or Modo = 2 Or (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2)))
    With formulario
        For i = 0 To .imgZoom.Count - 1
            .imgZoom(i).Enabled = b
            .imgZoom(i).visible = b
        Next i
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Sub




Public Sub BloquearImgFec(ByRef formulario As Form, Index As Integer, Modo As Byte, Optional ModoLineas As Byte)
Dim b As Boolean
    On Error Resume Next

    b = (Modo = 3 Or Modo = 4 Or Modo = 1 Or (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2)))
    formulario.imgFec(Index).Enabled = b
    formulario.imgFec(Index).visible = b
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearImage(ByRef img As Image, b As Boolean)

    On Error Resume Next
    
    img.Enabled = Not b
    img.visible = Not b
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub BloquearBtn(ByRef btn As CommandButton, b As Boolean)

    On Error Resume Next
    
    btn.Enabled = Not b
  '  btn.visible = Not b
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub BloquearList(ByRef List As ListBox, b As Boolean)
On Error Resume Next

    'List.Locked = b
    List.Enabled = Not b
    'If Not b And List.Enabled = False Then List.Enabled = True
    If b Then
        List.BackColor = &H80000018 'Amarillo Claro
    Else
        List.BackColor = vbWhite
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BloquearOption(ByRef Opt As OptionButton, b As Boolean)
On Error Resume Next

    'Opt.Locked = b
    Opt.Enabled = Not b
    'If Not b And Opt.Enabled = False Then Opt.Enabled = True
    If b Then
        Opt.BackColor = &H80000018 'Amarillo Claro
    Else
        Opt.BackColor = vbWhite
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub PonerLongCamposGnral(ByRef formulario As Form, Modo As Byte, Opcion As Byte)
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'ya que en busqueda se permite introducir criterios más largos del tamaño del campo
'en busqueda permitimos escribir: "0001:0004"
'en cambio al insertar o modificar la longitud solo debe permitir ser: "0001"
'(IN) formulario y Modo en que se encuentra el formulario
'(IN) Opcion : 1 para los TEXT1, 3 para los txtAux

    Dim i As Integer
    
    On Error Resume Next

    With formulario
        If Modo = 1 Then 'BUSQUEDA
            Select Case Opcion
                Case 1 'Para los TEXT1
                    For i = 0 To .Text1.Count - 1
                        With .Text1(i)
                            If .MaxLength <> 0 Then
                               .HelpContextID = .MaxLength 'guardamos es maxlenth para reestablecerlo despues
                                .MaxLength = 0 'tamaño infinito
                            End If
                        End With
                    Next i
                
                Case 3 'para los TXTAUX
                    For i = 0 To .txtAux.Count - 1
                        With .txtAux(i)
                            If .MaxLength <> 0 Then
                               .HelpContextID = .MaxLength 'guardamos es maxlenth para reestablecerlo despues
                                .MaxLength = 0 'tamaño infinito
                            End If
                        End With
                    Next i
            End Select
            
        Else 'resto de modos
            Select Case Opcion
                Case 1 'par los Text1
                    For i = 0 To .Text1.Count - 1
                        With .Text1(i)
                            If .HelpContextID <> 0 Then
                                .MaxLength = .HelpContextID 'volvemos a poner el valor real del maxlenth
                                .HelpContextID = 0
                            End If
                        End With
                    Next i
                Case 3 'para los txtAux
                    For i = 0 To .txtAux.Count - 1
                        With .txtAux(i)
                            If .HelpContextID <> 0 Then
                                .MaxLength = .HelpContextID 'volvemos a poner el valor real del maxlenth
                                .HelpContextID = 0
                            End If
                        End With
                    Next i
            End Select
        End If
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub DesplazamientoData(ByRef vData As Adodc, Index As Integer)
'Para desplazarse por los registros de control Data
    If vData.Recordset.EOF Then Exit Sub
    Select Case Index
        Case 0 'Primer Registro
            If Not vData.Recordset.BOF Then vData.Recordset.MoveFirst
        Case 1 'Anterior
            vData.Recordset.MovePrevious
            If vData.Recordset.BOF Then vData.Recordset.MoveFirst
        Case 2 'Siguiente
            vData.Recordset.MoveNext
            If vData.Recordset.EOF Then vData.Recordset.MoveLast
        Case 3 'Ultimo
            vData.Recordset.MoveLast
    End Select
End Sub


'===========================
Public Function SituarData(ByRef vData As Adodc, vWhere As String, ByRef Indicador As String, Optional NoRefresca As Boolean) As Boolean
'Situa un DataControl en el registo que cumple vwhere
    On Error GoTo ESituarData

        'Actualizamos el recordset
        If Not NoRefresca Then vData.Refresh
        
        'El sql para que se situe en el registro en especial es el siguiente
        vData.Recordset.Find vWhere
        If vData.Recordset.EOF Then
            If vData.Recordset.RecordCount > 0 Then vData.Recordset.MoveFirst
            GoTo ESituarData
        End If
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarData = True
        Exit Function

ESituarData:
        If Err.Number <> 0 Then Err.Clear
        SituarData = False
End Function

'===========================
Public Function SituarDataMULTI(ByRef vData As Adodc, vWhere As String, ByRef Indicador As String, Optional NoRefresca As Boolean) As Boolean
'Situa un DataControl en el registo que cumple vwhere
On Error GoTo ESituarData
        'Actualizamos el recordset
        If Not NoRefresca Then vData.Refresh
        'El sql para que se situe en el registro en especial es el siguiente
        Multi_Find vData.Recordset, vWhere
        'vData.Recordset.Find vWhere
        If vData.Recordset.EOF Then GoTo ESituarData
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarDataMULTI = True
        Exit Function
ESituarData:
        If Err.Number <> 0 Then Err.Clear
        SituarDataMULTI = False
End Function


Public Sub Multi_Find(ByRef oRs As ADODB.Recordset, sCriteria As String)

    Dim clone_rs As ADODB.Recordset
    Set clone_rs = oRs.Clone
    
    clone_rs.Filter = sCriteria
    
    If clone_rs.EOF Or clone_rs.BOF Then
     oRs.MoveLast
     oRs.MoveNext
    Else
     oRs.Bookmark = clone_rs.Bookmark
    End If
    
    clone_rs.Close
    Set clone_rs = Nothing

End Sub

' ### [Monica] 02/10/2006 añadido de laura
Public Sub Multi_Find2(ByRef oRs As ADODB.Recordset, sCriteria As String)
'para el situarDataMULTI
On Error Resume Next

    oRs.Filter = ""
    oRs.MoveFirst
    oRs.Filter = sCriteria
    
    If oRs.EOF Or oRs.BOF Then
     oRs.MoveLast
     oRs.MoveNext
    Else
'        x = oRs.AbsolutePosition
'     oRs.Bookmark = clone_rs.Bookmark
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub

'===========================
'## SUSTITUIR POR SituarDataMulti
Public Function SituarDataGen(ByRef vData As Adodc, ByRef T1 As TextBox, ByRef T2 As TextBox, Indicador As String) As Boolean
'Situa un DataControl en el registo que cumple vwhere
Dim mTag1 As CTag, mTag2 As CTag
Dim Valor1 As Variant, Valor2 As Variant
Dim Dato1, Dato2
Dim Encontrado As Boolean
On Error GoTo ESituarData

    SituarDataGen = False
    
    If T1.Tag <> "" And T2.Tag <> "" Then
        'Cargamos el Tag del TEXT1
        Set mTag1 = New CTag
        mTag1.Cargar T1
        If mTag1.Cargado Then
            Select Case mTag1.TipoDato
                Case "T": Valor1 = T1.Text
                Case "N": Valor1 = Val(T1.Text)
            End Select
        Else
            Exit Function
        End If
        
        'Cargamos el Tag del TEXT2
        Set mTag2 = New CTag
        mTag2.Cargar T2
        If mTag2.Cargado Then
            Select Case mTag2.TipoDato
                Case "T": Valor2 = T2.Text
                Case "N": Valor2 = Val(T2.Text)
            End Select
        Else
            Exit Function
        End If
        
        'Actualizamos el recordset
        vData.Refresh
        If vData.Recordset.EOF Then GoTo ESituarData
        
        Encontrado = False
        While Not Encontrado And Not vData.Recordset.EOF
            'valor del dato de la columna asociada al Text1
            Select Case mTag1.TipoDato
                Case "T": Dato1 = vData.Recordset.Fields(mTag1.columna).Value
                Case "N": Dato1 = Val(vData.Recordset.Fields(mTag1.columna).Value)
            End Select

            'valor del dato de la columna asociada al Text2
            Select Case mTag2.TipoDato
                Case "T": Dato2 = vData.Recordset.Fields(mTag2.columna).Value
                Case "N": Dato2 = Val(vData.Recordset.Fields(mTag2.columna).Value)
            End Select

            If Dato1 = Valor1 And Dato2 = Valor2 Then
'                If cod3 = "" Then
                    Encontrado = True
'                Else
'                    Select Case T3
'                        Case "T": Dato3 = vData.Recordset.Fields(2).Value
'                        Case "N": Dato3 = Val(vData.Recordset.Fields(2).Value)
'                    End Select
'                    If Dato3 = valor3 Then
'                        encontrado = True
'                    Else
'                        vData.Recordset.MoveNext
'                    End If
'                End If
            Else
                vData.Recordset.MoveNext
            End If
        Wend
        Set mTag1 = Nothing
        Set mTag2 = Nothing
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarDataGen = True
        Exit Function
    End If
ESituarData:
        If Err.Number <> 0 Then Err.Clear
        SituarDataGen = False
End Function




'===========================
Public Function SituarDataPosicion(ByRef vData As Adodc, NumPos As Long, ByRef Indicador As String) As Boolean
'Situa un DataControl en el registro que ocupa la posicion NumPos
Dim TotalReg As Long

    On Error GoTo ESituarDataPosicion
    
'        'Actualizamos el recordset
'        If Not NoRefresca Then vdata.Refresh

        TotalReg = vData.Recordset.RecordCount
        
        If vData.Recordset.EOF Then GoTo ESituarDataPosicion
        
        If NumPos <= TotalReg Then
            vData.Recordset.Move NumPos - 1
        Else
            vData.Recordset.Move NumPos
        End If
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarDataPosicion = True
        Exit Function
        
ESituarDataPosicion:
        If Err.Number <> 0 Then Err.Clear
        SituarDataPosicion = False
End Function






Public Function SituarDataTrasEliminar(ByRef vData As Adodc, NumReg, Optional no_refre As Boolean) As Boolean
    On Error GoTo ESituarDataElim

    If Not no_refre Then vData.Refresh 'quan siga False o no es passe a la funció, es refrescarà. Hi ha que passar-lo com a True quan el manteniment siga Grid per a que no refresque
    
    If Not vData.Recordset.EOF Then    'Solo habia un registro
        If NumReg > vData.Recordset.RecordCount Then
            vData.Recordset.MoveLast
        Else
            vData.Recordset.MoveFirst
            vData.Recordset.Move NumReg - 1
        End If
        SituarDataTrasEliminar = True
    Else
        SituarDataTrasEliminar = False
    End If
        
ESituarDataElim:
    If Err.Number <> 0 Then
        Err.Clear
        SituarDataTrasEliminar = False
    End If
End Function


Public Sub PonerFoco(ByRef Text As TextBox)
On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoBtn(ByRef btn As CommandButton)
On Error Resume Next
    If btn.visible Then btn.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoCmb(ByRef combo As ComboBox)
On Error Resume Next
    combo.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub PonerFocoChk(ByRef chk As CheckBox)
On Error Resume Next
    chk.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub PonerFocoGrid(ByRef DGrid As DataGrid)
    On Error Resume Next
    DGrid.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub PonerFocoListView(ByRef LView As ListView)
    On Error Resume Next
    LView.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub




Public Sub ConseguirFoco(ByRef Text As TextBox, Modo As Byte)
'Acciones que se realizan en el evento:GotFocus de los TextBox:Text1
'en los formularios de Mantenimiento
On Error Resume Next

    If (Modo <> 0 And Modo <> 2) Then
        If Modo = 1 Then 'Modo 1: Busqueda
            Text.BackColor = vbYellow
        End If
        Text.SelStart = 0
        Text.SelLength = Len(Text.Text)
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub ConseguirFocoLin(ByRef Text As TextBox)
'Acciones que se realizan en el evento:GotFocus de los TextBox:TxtAux para LINEAS
'en los formularios de Mantenimiento
On Error Resume Next

'    If (Modo <> 0 And Modo <> 2) Then
'        If Modo = 1 Then 'Modo 1: Busqueda
'            Text.BackColor = vbYellow
'        End If
        With Text
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
'    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function PerderFocoGnral(ByRef Text As TextBox, Modo As Byte) As Boolean
Dim Comprobar As Boolean
'Dim mTag As CTag

    On Error Resume Next

    If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then
        PerderFocoGnral = False
        Exit Function
    End If

    With Text
        'Quitamos blancos por los lados
        .Text = Trim(.Text)
        
        
         If .BackColor = vbYellow Then
            If .Locked Then
                .BackColor = &H80000018
            Else
                .BackColor = vbWhite
            End If
        End If
        
        
        'Si no estamos en modo: 3=Insertar o 4=Modificar o 1=Busqueda, no hacer ninguna comprobacion
        If (Modo <> 3 And Modo <> 4 And Modo <> 1 And Modo <> 5) Then
            PerderFocoGnral = False
            Exit Function
        End If
        
        If Modo = 1 Then
            'Si estamos en modo busqueda y contiene un caracter especial no realizar
            'las comprobaciones
            Comprobar = ContieneCaracterBusqueda(.Text)
            If Comprobar Then
                PerderFocoGnral = False
                Exit Function
            End If
        End If
        PerderFocoGnral = True
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function PerderFocoGnralLineas(ByRef Txt As TextBox, ModoLineas As Byte) As Boolean
'Para el LostFocus de los txtAux de Mto de lineas


    On Error Resume Next

    If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then
        PerderFocoGnralLineas = False
        Exit Function
    End If
    
    With Txt
        'Quitamos blancos por los lados
        .Text = Trim(.Text)

        If .BackColor = vbYellow Then
'    '        Text1(Index).BackColor = &H80000018
            .BackColor = vbWhite
        End If

        'Si no estamos en modo: 1=Insertar o 4=Modificar o 1=Busqueda, no hacer ninguna comprobacion
        If (ModoLineas <> 1 And ModoLineas <> 2) Then
            PerderFocoGnralLineas = False
            Exit Function
        End If
    End With

    PerderFocoGnralLineas = True
    If Err.Number <> 0 Then Err.Clear



'Dim Comprobar As Boolean
'On Error Resume Next
'    With Txt
'
'        'Quitamos blancos por los lados
'        .Text = Trim(.Text)
'
'        If .BackColor = vbYellow Then
'    '        Text1(Index).BackColor = &H80000018
'            .BackColor = vbWhite
'        End If
'
'        'Si no estamos en modo: 1=Insertar o 4=Modificar o 1=Busqueda, no hacer ninguna comprobacion
'        If (ModoLineas <> 1 And ModoLineas <> 2 And ModoLineas <> 1) Then
'            PerderFocoGnralLineas = False
'            Exit Function
'        End If
'
'        If ModoLineas = 1 Then
'            'Si estamos en modo busqueda y contiene un caracter especial no realizar
'            'las comprobaciones
'            Comprobar = ContieneCaracterBusqueda(.Text)
'            If Comprobar Then
'                PerderFocoGnralLineas = False
'                Exit Function
'            End If
'        End If
'        PerderFocoGnralLineas = True
'    End With
'    If Err.Number <> 0 Then Err.Clear
End Function


Public Sub limpiar(ByRef formulario As Form)
    Dim Control As Object

    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        End If
    Next Control
End Sub



Public Sub LimpiarText1(ByRef formulario As Form)
'Dim i As Integer
'
'    With formulario
'        For i = 0 To .Text1.Count - 1
'            .Text1(i).Text = ""
'        Next i
'    End With
End Sub


Public Sub LimpiarTxtAux(ByRef formulario As Form)
'Dim i As Integer
'
'    With formulario
'        For i = 0 To .txtAux.Count - 1
'            .txtAux(i).Text = ""
'        Next i
'    End With
End Sub


Public Sub LimpiarLin(ByRef formulario As Form, nomframe As String)
'Limpiar los controles Text que esten dentro del frame nomFrame
    Dim Control As Object

    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            If Control.Container.Name = nomframe Then
                Control.Text = ""
            End If
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Container.Name = nomframe Then
                Control.ListIndex = -1
            End If
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Container.Name = nomframe Then
                Control.Value = 0
            End If
        End If
    Next Control
End Sub



Public Function EsVacio(ByRef campo As TextBox) As Boolean
'    If (campo.Text = "" Or campo.Text = "0") Then
'        EsVacio = True
'    Else
'        EsVacio = False
'    End If
End Function




Public Sub DesplazamientoVisible(ByRef toolb As Toolbar, iniBoton As Byte, bol As Boolean, NReg As Byte)
'Oculta o Muestra las botones de desplazamiento de la toolbar
Dim i As Byte

    Select Case NReg
        Case 0, 1 '0 o 1 registro no mostrar los botones despl.
            For i = iniBoton To iniBoton + 3
                toolb.Buttons(i).visible = False
            Next i
        Case Else '>1 reg, mostrar si bol
            For i = iniBoton To iniBoton + 3
                toolb.Buttons(i).visible = bol
            Next i
    End Select
End Sub



Public Function EsNumerico(Texto As String) As Boolean
Dim i As Integer
Dim c As Integer
Dim L As Integer
Dim Cad As String
Dim b As Boolean
    
    EsNumerico = False
    b = True
    Cad = ""
    If Not IsNumeric(Texto) Then
        Cad = "El campo debe ser numérico"
        b = False
        '======= Añade Laura
        'formato: (.25)
        i = InStr(1, Texto, ".")
        If i = 1 Then
            If IsNumeric(Mid(Texto, 2, Len(Texto))) Then b = True
        'añado el caso -.25 [Monica]04/06/2013
        Else
            If i = 2 And Mid(Texto, 1, 1) = "-" Then
                If IsNumeric(Mid(Texto, 3, Len(Texto))) Then b = True
            End If
        End If
        '======================
    Else
        'Vemos si ha puesto mas de un punto
        c = 0
        L = 1
        Do
            i = InStr(L, Texto, ",")
            If i > 0 Then
                L = i + 1
                c = c + 1
            End If
        Loop Until i = 0
        If c > 1 Then
            Cad = "Numero de comas incorrecto"
            b = False
        End If
        
        'Si no ha puesto ninguna coma y tiene más de un punto
        If c = 0 Then
            L = 1
            Do
                i = InStr(L, Texto, ".")
                If i > 0 Then
                    L = i + 1
                    c = c + 1
                End If
            Loop Until i = 0
            If c > 1 Then
                Cad = "Numero incorrecto"
                b = False
            End If
        End If
    End If
    If Not b Then
        MsgBox Cad, vbExclamation
    Else
        EsNumerico = b
    End If
End Function


Public Function PonerFormatoEntero(ByRef T As TextBox) As Boolean
'Comprueba que el valor del textbox es un entero y le pone el formato
Dim mTag As CTag
Dim Cad As String
Dim Formato As String
On Error GoTo EPonerFormato

    
    If T.Text = "" Then Exit Function
    PonerFormatoEntero = True
    
    Set mTag = New CTag
    mTag.Cargar T
    If mTag.Cargado Then
       Cad = mTag.Nombre 'descripcion del campo
       Formato = mTag.Formato
    End If
    Set mTag = Nothing

    If Not EsEntero(T.Text) Then
        PonerFormatoEntero = False
        MsgBox "El campo " & Cad & " tiene que ser numérico.", vbExclamation
        PonerFoco T
    Else
         'T.Text = Format(T.Text, Formato)
         ' **** 21-11-2005 Canvi de Cèsar. Per a que formatetge be si es posa un
         ' número negatiu, li lleve un 0 a la màscara per a que el número
         ' càpiga dins del textbox en el maxlength asignat.
         ' Si es crida a esta funció la màscara es del tipo 0000
         If T.Text < 0 Then _
            Formato = Replace(Formato, "0", "", 1, 1)
        ' *************************************************************************
         
         T.Text = Format(T.Text, Formato)
    End If
    
EPonerFormato:
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function PosarFormatTelefon(ByRef T As TextBox) As Boolean
'Comprova que el Telèfon/Fax/Mòbil no te espais en blanc i només té números
Dim mTag As CTag
Dim Cad As String

On Error GoTo EPosarFormatTelefon

    If T.Text = "" Then Exit Function
    PosarFormatTelefon = True
    
    T.Text = Replace(T.Text, " ", "")
       
    Set mTag = New CTag
    mTag.Cargar T
    If mTag.Cargado Then
       Cad = mTag.Nombre 'descripció del camp
    End If
    Set mTag = Nothing

    If (InStr(1, T.Text, ",") > 0) Or (InStr(1, T.Text, ".") > 0) Or (InStr(1, T.Text, "+") > 0) Or (InStr(1, T.Text, "-") > 0) Or (Not IsNumeric(T.Text)) Then
        PosarFormatTelefon = False
        MsgBox "El campo " & Cad & " tiene que ser numérico.", vbExclamation
        PonerFoco T
    End If
    
EPosarFormatTelefon:
    If Err.Number <> 0 Then Err.Clear
End Function


'=================================
Public Function PonerFormatoDecimal(ByRef T As TextBox, tipoF As Single) As Boolean
'tipoF: tipo de Formato a aplicar
'  1 -> Decimal(12,2)
'  2 -> Decimal(8,3)  antes 10,4  ### [Monica] 25/09/2006
'  3 -> Decimal(10,2)
'  4 -> Decimal(5,2)
'  5 -> Decimal(10,3) antes 8,4    ### [Monica] 25/09/2006
'  6 -> Decimal(5,4)  nuevo para impuesto   ### [Monica] 25/09/2006
'  7 -> Decimal(8,4)  nuevo
'  8 -> Decimal(6,4)  nuevo
'  9 -> Decimal(8,6)  nuevo
' 10 -> decimal(6,2)  nuevo
' 11 -> decimal(10,4)  nuevo
' 12 -> decimal(6,3)

Dim Valor As Double
Dim PEntera As Currency
Dim NoOK As Boolean
Dim i As Byte
Dim cadEnt As String
'Dim mTas As CTag

    If T.Text = "" Then Exit Function
    PonerFormatoDecimal = False
    NoOK = False
    With T
'        If Not EsEntero(.Text) Then
        If Not EsNumerico(CStr(.Text)) Then
'             MsgBox "El campo debe ser numérico.", vbExclamation
'            .Text = ""
            PonerFoco T
            Exit Function
        End If


        If InStr(1, .Text, ",") > 0 Then
            Valor = ImporteFormateado(.Text)
        Else
            cadEnt = .Text
            i = InStr(1, cadEnt, ".")
            If i > 0 Then cadEnt = Mid(cadEnt, 1, i - 1)
            If tipoF = 1 And Len(cadEnt) > 10 Then
                MsgBox "El valor no puede ser mayor de 9999999999,99", vbExclamation
                NoOK = True
            End If
            If NoOK Then
'                    .Text = ""
                T.SetFocus
                Exit Function
            End If
            Valor = CDbl(TransformaPuntosComas(.Text))
        End If
            
        'Comprobar la longitud de la Parte Entera
        PEntera = Int(Valor)
        Select Case tipoF 'Comprobar longitud
            Case 1 'Decimal(12,2)
                If Len(CStr(PEntera)) > 10 Then
                    MsgBox "El valor no puede ser mayor de 9999999999,99", vbExclamation
                    NoOK = True
                End If
            Case 2 'Decimal(8,3)
                If Len(CStr(PEntera)) > 5 Then
                    MsgBox "El valor no puede ser mayor de 99999,999", vbExclamation
                    NoOK = True
                End If
            Case 3 'Decimal(10,2)
                If Len(CStr(PEntera)) > 8 Then
                    MsgBox "El valor no puede ser mayor de 99999999,99", vbExclamation
                    NoOK = True
                End If
            Case 4 'Decimal(5,2)
                If Len(CStr(PEntera)) > 3 Then
                    MsgBox "El valor no puede ser mayor de 999,99", vbExclamation
                    NoOK = True
                End If
            Case 5 'Decimal(10,3)
                If Len(CStr(PEntera)) > 7 Then
                    MsgBox "El valor no puede ser mayor de 9999999,999", vbExclamation
                    NoOK = True
                End If
            Case 6 'decimal(5,4)
                If Len(CStr(PEntera)) > 1 Then
                    MsgBox "El valor no puede ser mayor de 9,9999", vbExclamation
                    NoOK = True
                End If
            Case 7 'decimal(8,4)
                If Len(CStr(PEntera)) > 4 Then
                    MsgBox "El valor no puede ser mayor de 9999,9999", vbExclamation
                    NoOK = True
                End If
            Case 8 'decimal(6,4)
                If Len(CStr(PEntera)) > 2 Then
                    MsgBox "El valor no puede ser mayor de 99,9999", vbExclamation
                    NoOK = True
                End If
            Case 9 'decimal(8,6)
                If Len(CStr(PEntera)) > 2 Then
                    MsgBox "El valor no puede ser mayor de 99,999999", vbExclamation
                    NoOK = True
                End If
            Case 10 'decimal(6,2)
                If Len(CStr(PEntera)) > 4 Then
                    MsgBox "El valor no puede ser mayor de 9999,99", vbExclamation
                    NoOK = True
                End If
            Case 11 'decimal(10,4)
                If Len(CStr(PEntera)) > 6 Then
                    MsgBox "El valor no puede ser mayor de 999999,9999", vbExclamation
                    NoOK = True
                End If
            Case 12 'decimal(6,3)
                If Len(CStr(PEntera)) > 3 Then
                    MsgBox "El valor no puede ser mayor de 999,999", vbExclamation
                    NoOK = True
                End If
                       
            
        End Select




'       valor = CCur(TransformaPuntosComas(.Text))
'        If Not EsNumerico(CStr(valor)) Then
'             MsgBox "El campo debe ser numérico.", vbExclamation
''            .Text = ""
'            PonerFoco T
'        Else
'            Set mTag = New CTag
'            If mTag.Cargar(T) Then
'                NoOK = mTag.Comprobar(T)
'                If NoOK = False Then Exit Function
'            End If
'            Set mTag = Nothing
            
           
            
            If NoOK Then
                PonerFormatoDecimal = False
'                .Text = ""
                T.SetFocus
                Exit Function
            End If
            
            'Poner el Formato
            Select Case tipoF
                Case 1 'Formato Decimal(12,2)
                    .Text = Format(Valor, FormatoImporte)
                Case 2 'Formato Decimal(8,3)
                    .Text = Format(Valor, FormatoPrecio)
                Case 3 'Formato Decimal(10,2)
                    .Text = Format(Valor, FormatoDec10d2)
                Case 4 'Formato Decimal(5,2)
                    .Text = Format(Valor, FormatoPorcen)
                Case 5 'Formato Decimal(10,3)
                    .Text = Format(Valor, FormatoDec10d3)
                Case 6 'Formato Decimal(5,4)
                    .Text = Format(Valor, FormatoDec5d4)
                Case 7 'Formato Decimal(8,4)
                    .Text = Format(Valor, FormatoDec8d4)
                Case 8 'Formato Decimal(6,4)
                    .Text = Format(Valor, FormatoDec6d4)
                Case 9 'Formato Decimal(8,6)
                    .Text = Format(Valor, FormatoDec8d6)
                Case 10 'Formato Decimal(6,2)
                    .Text = Format(Valor, FormatoDec6d2)
                Case 11 'Formato Decimal(10,4)
                    .Text = Format(Valor, FormatoDec10d4)
                Case 12 'Formato Decimal(6,3)
                    .Text = Format(Valor, FormatoDec6d3)
            
            
            End Select
            PonerFormatoDecimal = True
'        End If
    End With
End Function


Public Function PonerNombreDeCod(ByRef Txt As TextBox, Tabla As String, campo As String, Optional Codigo As String, Optional Tipo As String, Optional cBD As Byte, Optional codigo2 As String, Optional Valor2 As String, Optional tipo2 As String) As String
'Devuelve el nombre/Descripción asociado al Código correspondiente
'Además pone formato al campo txt del código a partir del Tag
Dim Sql As String
Dim devuelve As String
Dim vtag As CTag
Dim ValorCodigo As String

    On Error GoTo EPonerNombresDeCod

    ValorCodigo = Txt.Text
    If ValorCodigo <> "" Then
        Set vtag = New CTag
        If vtag.Cargar(Txt) Then
            If Codigo = "" Then Codigo = vtag.columna
            If Tipo = "" Then Tipo = vtag.TipoDato
            
            If cBD = 0 Then cBD = cAgro
            Sql = DevuelveDesdeBDNew(cBD, Tabla, campo, Codigo, ValorCodigo, Tipo, , codigo2, Valor2, tipo2)
            If vtag.TipoDato = "N" Then ValorCodigo = Format(ValorCodigo, vtag.Formato)
            Txt.Text = ValorCodigo 'Valor codigo formateado
            If Sql = "" Then
'                If vtag.Nombre <> "" Then
'                    devuelve = "No existe el " & vtag.Nombre & ": " & ValorCodigo
'                Else
'                    devuelve = "No existe el " & Texto & ": " & ValorCodigo
'                End If
'                MsgBox devuelve, vbExclamation
'                Txt.Text = ""
'                PonerFoco Txt
            Else
                PonerNombreDeCod = Sql 'Descripcion del codigo
            End If
        End If
        Set vtag = Nothing
    Else
        PonerNombreDeCod = ""
    End If
'    Exit Function
EPonerNombresDeCod:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Nombre asociado a código: " & Codigo, Err.Description
End Function





Public Sub PonerIndicador(ByRef lblIndicador As Label, Modo As Byte, Optional ModoLineas As Byte)
'Pone el titulo del label lblIndicador
    Select Case Modo
        Case 0    'Modo Inicial
            lblIndicador.Caption = ""
        Case 1 'Modo Buscar
            lblIndicador.Caption = "BUSQUEDA"
        Case 2    'Preparamos para que pueda Modificar
            lblIndicador.Caption = ""

        Case 3 'Modo Insertar
            lblIndicador.Caption = "INSERTAR"
        Case 4 'MODIFICAR
            lblIndicador.Caption = "MODIFICAR"
            
        Case 5 'Modo Lineas
            If ModoLineas = 1 Then
                lblIndicador.Caption = "INSERTAR LINEA"
            ElseIf ModoLineas = 2 Then
                lblIndicador.Caption = "MODIFICAR LINEA"
            End If
        Case Else
            lblIndicador.Caption = ""
    End Select
End Sub

Public Function PonerContRegistros(ByRef vData As Adodc) As String
'indicador del registro donde nos encontramos: "1 de 20"
    On Error GoTo EPonerReg
    
    If Not vData.Recordset.EOF Then
        PonerContRegistros = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
    Else
        PonerContRegistros = ""
    End If
    
EPonerReg:
    If Err.Number <> 0 Then
        Err.Clear
        PonerContRegistros = ""
    End If
End Function


Public Sub KEYdown(KeyCode As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
On Error Resume Next
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
            'SendKeys "+{tab}"
            CreateObject("WScript.Shell").SendKeys "+{tab}"
            
        Case 40 'Desplazamiento Flecha Hacia Abajo
            'SendKeys "{tab}"
            CreateObject("WScript.Shell").SendKeys "{tab}"
            
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub

' ### [Monica] 06/09/2006
' añadido este procedimiento del ariges de laura
Public Sub KEYpressGnral(KeyAscii As Integer, Modo As Byte, cerrar As Boolean)
'IN: codigo keyascii tecleado, y modo en que esta el formulario
'OUT: si se tiene que cerrar el formulario o no
    cerrar = False
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
'        SendKeys "{tab}"
        CreateObject("WScript.Shell").SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then cerrar = True
    End If
End Sub

Public Sub AnyadirLinea(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
On Error Resume Next

    vDataGrid.AllowAddNew = True
    If vData.Recordset.RecordCount > 0 Then
        vDataGrid.HoldFields
        vData.Recordset.MoveLast
        vDataGrid.Row = vDataGrid.Row + 1
    End If
    vDataGrid.Enabled = False
    
    If Err.Number <> 0 Then Err.Clear
End Sub





Public Function LanzaHomeGnral(nomWeb As String) As Boolean
On Error GoTo ELanzaHome
Dim ruta As String
'Dim wb As WebClassLibrary
'Dim wb1 As WebClass

    LanzaHomeGnral = False
    
    If nomWeb = "" Then
        MsgBox "No hay una dirección Web para mostrar.", vbInformation
        Exit Function
    End If

'    If Opcion = "webversion" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "?version=" & App.Major & "." & App.Minor & "." & App.Revision

'    Set wb = New WebClassLibrary
'    Set wb1 = New WebClass
'
'    wb1.SERVER = Ruta & " " & nomWeb
'    If wb1.Error Then
'        MsgBox "error en el sigpac", vbExclamation
'    End If

    'Lanzamos
'   ruta = "C:\Archivos de programa\Internet Explorer\IEXPLORE.EXE"
    If vConfig.Explorador <> "" Then
       Shell vConfig.Explorador & " " & nomWeb, vbMaximizedFocus
'        Shell ruta & " " & nomWeb, vbMaximizedFocus
        LanzaHomeGnral = True
    End If
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, nomWeb & vbCrLf & Err.Description
End Function



Public Function LanzaMailGnral(dirMail As String) As Boolean
'LLama al Programa de Correo (Outlook,...)
On Error GoTo ELanzaHome

    LanzaMailGnral = False
    If dirMail = "" Then
        MsgBox "No hay dirección e-mail a la que enviar.", vbExclamation
        Exit Function
    End If

    Call ShellExecute(hWnd, "Open", "mailto: " & dirMail, "", "", vbNormalFocus)
    LanzaMailGnral = True
    
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, vbCrLf & Err.Description
'    CadenaDesdeOtroForm = ""
End Function


Public Sub SubirItemList(ByRef LView As ListView)
'Subir el item seleccionado del listview una posicion
Dim i As Byte, item As Byte
Dim Aux As String
On Error Resume Next
   
    For i = 2 To LView.ListItems.Count
        If LView.ListItems(i).Selected Then
            item = i
            Aux = LView.ListItems(i).Text
            LView.ListItems(i).Text = LView.ListItems(i - 1).Text
            LView.ListItems(i - 1).Text = Aux
        End If
    Next i
    If item <> 0 Then
        LView.ListItems(item).Selected = False
        LView.ListItems(item - 1).Selected = True
    End If
    LView.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub BajarItemList(ByRef LView As ListView)
'Bajar el item seleccionado del listview una posicion
Dim i As Byte, item As Byte
Dim Aux As String
On Error Resume Next

    For i = 1 To LView.ListItems.Count - 1
        If LView.ListItems(i).Selected Then
            item = i
            Aux = LView.ListItems(i).Text
            LView.ListItems(i).Text = LView.ListItems(i + 1).Text
            LView.ListItems(i + 1).Text = Aux
        End If
    Next i
    If item <> 0 Then
        LView.ListItems(item).Selected = False
        LView.ListItems(item + 1).Selected = True
    End If
    LView.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function EsCodigoCero(cod As String, Formato As String) As Boolean
    EsCodigoCero = False
    If cod <> "" Then
        If IsNumeric(cod) Then
            If Val(cod) = Val(0) Then
                EsCodigoCero = True
                MsgBox "El código " & Formato & " no se puede modificar ni eliminar.", vbExclamation
                Screen.MousePointer = vbDefault
            End If
        End If
    End If
End Function




Public Sub CargaGridGnral(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, Sql As String, PrimeraVez As Boolean)
    On Error GoTo ECargaGRid

    vDataGrid.Enabled = True
    '    vdata.Recordset.Cancel
    vData.ConnectionString = conn
    vData.RecordSource = Sql
    vData.CursorType = adOpenDynamic
    vData.LockType = adLockPessimistic
    vDataGrid.ScrollBars = dbgNone
    vData.Refresh
    
    Set vDataGrid.DataSource = vData
    vDataGrid.AllowRowSizing = False
    vDataGrid.RowHeight = 290
    
    If PrimeraVez Then
        vDataGrid.ClearFields
        vDataGrid.ReBind
        vDataGrid.Refresh
    End If
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "CargaGrid", Err.Description
End Sub


Public Sub DeseleccionaGrid(ByRef vDataGrid As DataGrid)
    On Error GoTo EDeseleccionaGrid

    While vDataGrid.SelBookmarks.Count > 0
        vDataGrid.SelBookmarks.Remove 0
    Wend
    vDataGrid.SelStartCol = -1
    vDataGrid.SelEndCol = -1
    
    Exit Sub
        
EDeseleccionaGrid:
    Err.Clear
End Sub



Public Sub PosicionarCombo(ByRef Combo1 As ComboBox, Valor As Integer)
'Situa el combo en la posicion de un valor concreto
Dim J As Integer

    On Error GoTo EPosCombo
    
    For J = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(J) = Valor Then
            Combo1.ListIndex = J
            Exit For
        End If
    Next J

EPosCombo:
    If Err.Number <> 0 Then Err.Clear
End Sub


' posicionamos el combo cogiendo sólo las tres primeras posiciones

Public Sub PosicionarCombo2(ByRef Combo1 As ComboBox, Valor As String)
'Situa el combo en la posicion de un valor concreto
Dim J As Integer

    On Error GoTo EPosCombo
    
    For J = 0 To Combo1.ListCount - 1
        If Mid(Combo1.List(J), 1, 3) = Trim(Valor) Then
            Combo1.ListIndex = J
            Exit For
        End If
    Next J

EPosCombo:
    If Err.Number <> 0 Then Err.Clear
End Sub





'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'   FUNCIONES Para PLANNER TOURS
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------

Public Sub DatosPoblacion(CodPobla As String, desPobla As String, CPostal As String, Provi As String, PAIS As String, Optional Prefix As String)
'IN --> codPobla
'OUT -> desPobla (Descripcion de la poblacion)
'        CPostal, Provi, Pais
Dim Sql As String
Dim Rs As ADODB.Recordset

    If CodPobla <> "" Then
        If EsEntero(CodPobla) Then
            Sql = "SELECT poblacio.despobla,poblacio.codposta, provinci.desprovi, naciones.desnacio, provinci.preprovi"
            Sql = Sql & " FROM poblacio, provinci, naciones WHERE codpobla= " & CodPobla
            Sql = Sql & " AND provinci.codprovi = poblacio.codprovi AND naciones.codnacio = provinci.codnacio"

            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, , , adCmdText
            If Not Rs.EOF Then
                CodPobla = Format(CodPobla, "000000")
                desPobla = Rs.Fields!desPobla
                CPostal = DBLet(Rs.Fields!codposta, "T")
                Provi = Rs.Fields!desProvi
                PAIS = Rs.Fields!desnacio
                If Not IsNull(Rs.Fields!preprovi) Then _
                    Prefix = CStr(Rs.Fields!preprovi)
            Else
'                MsgBox "No existe el código de Población: " & codPobla, vbInformation
                CodPobla = "NoExiste"
                desPobla = ""
                CPostal = ""
                Provi = ""
                PAIS = ""
                Prefix = ""
            End If
            Rs.Close
            Set Rs = Nothing
        Else
             MsgBox "El Código de Población debe ser numérico.", vbInformation
             CodPobla = ""
        End If
    Else
        CodPobla = ""
        desPobla = ""
        CPostal = ""
        Provi = ""
        PAIS = ""
    End If
End Sub


Public Sub PonerDatosPoblacion(ByRef Tcpob As TextBox, ByRef Tdpob As TextBox, Optional Tcp As TextBox, Optional Tdprov As TextBox, Optional Tdpai As TextBox, Optional Nuevo As Boolean, Optional Telefon As TextBox)
Dim CodPobla As String, desPobla As String
Dim CPostal As String
Dim desProvi As String, desPais As String
Dim Prefix As String
Dim cadMen As String

    CodPobla = Tcpob.Text
    DatosPoblacion CodPobla, desPobla, CPostal, desProvi, desPais, Prefix
    Tdpob.Text = desPobla
    'Tcp.Text = CPostal
    If Not Tcp Is Nothing Then Tcp.Text = CPostal
    If Not Tdprov Is Nothing Then Tdprov.Text = desProvi
    If Not Tdpai Is Nothing Then Tdpai.Text = desPais
    If (Not Telefon Is Nothing) Then _
        If (Telefon.Text = "") Then Telefon.Text = Prefix
'    If Not Tdprov = Nothing Then
'        Tdprov.Text = desProvi
'    End If
'    Tdpai.Text = desPais
    If CodPobla = "NoExiste" Then
        'cadMen = "No existe el código de Población: " & Format(Tcpob.Text, "000000")
        cadMen = "No existe la Población: " & Format(Tcpob.Text, "000000")
        cadMen = cadMen & vbCrLf & "¿Desea Crearla?"
        If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
            Nuevo = True
'                    Indice = Index
'            Set frmPob = New frmPoblacio
'            frmPob.DatosADevolverBusqueda = "0|1|2|3|4|"
'            frmPob.NuevoCodigo = Tcpob.Text
'            Tcpob.Text = ""
'            TerminaBloquear
'            frmPob.Show vbModal
'            Set frmPob = Nothing
'            If Modo = 4 Then Bloquea = True
        Else
            CodPobla = ""
            Tcpob.Text = CodPobla
        End If
            PonerFoco Tcpob
    Else
        Tcpob.Text = CodPobla 'Devuelve el campo formateado
    End If
End Sub




Public Function PonerNomCliente(ByRef T As TextBox) As String
'Obtiene la cadena "apellido, nombre" o "nom.comercial" del cliente del codigo en T
'segun sea una persona o empresa.
Dim Cad As String, cadNom As String
Dim tipCli As String 'tipo de cliente (persona/empresa)
On Error Resume Next

    If T.Text = "" Then
        PonerNomCliente = ""
        Exit Function
    End If
    
    If PonerFormatoEntero(T) Then
'    If Not EsEntero(T.Text) Then
'        '***************+ canviar el mensage ***********************
'        MsgBox "El Código de Cliente tiene que ser numérico", vbExclamation
'        '**********************************************************++
'        T.Text = ""
'        PonerNomCliente = ""
'        PonerFoco T
'        Exit Function
'    Else
        Cad = "nom_come" 'nombre persona/nom comercial empresa
        tipCli = DevuelveDesdeBDNew(cAgro, "clientes", "tipclien", "codclien", T.Text, "N", Cad)
        If tipCli = "" Then
            MsgBox "No existe el cliente: " & T.Text, vbExclamation
            T.Text = ""
            PonerFoco T
        ElseIf tipCli = 1 Then 'persona
            T.Text = Format(T.Text, "000000")
            'obtenemos el Apellido
            cadNom = DevuelveDesdeBDNew(cAgro, "clientes", "ape_raso", "codclien", T.Text, "N")
            If cadNom <> "" Then
                cadNom = cadNom & ", " & Cad 'apellido, nombre
                PonerNomCliente = cadNom
            End If
        ElseIf tipCli = 2 Then 'empresa
            T.Text = Format(T.Text, "000000")
            PonerNomCliente = Cad
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Function









Public Function PonerNomProveedor(ByRef T As TextBox) As String
'Obtiene la cadena "nombre" del proveedor del codigo en T
Dim cadNom As String
Dim cadMen As String

    On Error Resume Next

'    Nuevo = False
    
    If T.Text = "" Then
        PonerNomProveedor = ""
        Exit Function
    End If
    
    If Not EsEntero(T.Text) Then
        '***************+ canviar el mensage ***********************
        MsgBox "El Código de Proveedor tiene que ser numérico", vbExclamation
        '**********************************************************++
        T.Text = ""
        PonerNomProveedor = ""
        PonerFoco T
        Exit Function
    Else
        cadNom = PonerNombreDeCod(T, "proveedo", "nomcomer", "codprove", "N")
'        cadNom = DevuelveDesdeBDnew(cAgro, "proveedor", "nomcomer", "codprove", T.Text, "N")
        If cadNom = "" Then
            cadMen = "No existe el proveedor: " & T.Text & vbCrLf
            MsgBox cadMen, vbExclamation
'            cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
'            If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                Nuevo = True
'            Else
                 T.Text = ""
'            End If
            PonerNomProveedor = ""
            PonerFoco T
        Else 'empresa
'            T.Text = Format(T.Text, "000000")
            PonerNomProveedor = cadNom
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
End Function




Public Function PonerBancoPropio(codempre As String, codBanpr As String, nomBanpr As String) As String
'devuelve la cuenta: ES-2077-0014-11-01010225252
'en nomBanco devuelve el nombre del banco
Dim Sql As String
Dim nomempre As String
Dim Rs As ADODB.Recordset

     'Poner banco Propio
    If codBanpr <> "" Then
        'comprobamos que existe el banco propio en la BD
        Sql = DevuelveDesdeBDNew(cAgro, "bancctas", "codbanpr", "codempre", codempre, "N", , "codbanpr", codBanpr, "N")
        If Sql = "" Then 'No existe el cod. banpr
            nomempre = DevuelveDesdeBDNew(cAgro, "empresas", "nomempre", "codempre", codempre, "N")
            Sql = "No existe el código de Banco Propio: " & codBanpr
            Sql = Sql & vbCrLf & "para la empresa: " & Format(codempre, "000") & " - " & nomempre
            MsgBox Sql, vbExclamation
            PonerBancoPropio = ""
            nomBanpr = "Error"
        Else
            Sql = "SELECT DISTINCT naciones.ibanpais, bancctas.codbanco, bancctas.codsucur, bancctas.digcontr, bancctas.ctabanco, bancsofi.nombanco "
            Sql = Sql & " FROM bancctas, naciones, bancsofi WHERE codempre = " & codempre & " AND codbanpr= " & codBanpr
            Sql = Sql & " AND bancctas.codnacio = naciones.codnacio "
            Sql = Sql & " AND (bancctas.codnacio = bancsofi.codnacio AND bancctas.codbanco = bancsofi.codbanco) "
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            PonerBancoPropio = Rs.Fields(0).Value & "-" & Format(Rs.Fields(1).Value, "0000") & "-" & Format(Rs.Fields(2).Value, "0000") & "-" & Format(Rs.Fields(3).Value, "00") & "-" & Format(Rs.Fields(4).Value, "0000000000")
            nomBanpr = Rs.Fields!NomBanco
            Rs.Close
            Set Rs = Nothing
        End If
    Else
        PonerBancoPropio = ""
        nomBanpr = ""
    End If
End Function



Public Function ValidarCuentaBancaria(ByRef txtB As TextBox, ByRef txtS As TextBox, ByRef txtDC As TextBox, ByRef txtC As TextBox) As Boolean
''IN: Controles textbox a Validar
'
'    ValidarCuentaBancaria = False
'
'    'Banco
'    If txtB.Text <> "" And Len(txtB.Text) < 4 Then
'            MsgBox "El campo Banco" & " debe tener 4 dígitos", vbExclamation
'            PonerFoco txtB
'            Exit Function
'    End If
'
'    'Sucursal
'    If txtS.Text <> "" And Len(txtS.Text) < 4 Then
'            MsgBox "El campo Sucursal" & " debe tener 4 dígitos", vbExclamation
'            PonerFoco txtS
'            Exit Function
'    End If
'
'    'Digito de Control
'    If txtDC.Text <> "" And Len(txtDC.Text) < 2 Then
'        MsgBox "El campo digito de control debe tener 2 dígitos", vbExclamation
'        PonerFoco txtDC
'        Exit Function
'    End If
'
'    'Cuenta Bancaria
'    If txtC.Text <> "" And Len(txtC.Text) < 10 Then
'        MsgBox "El campo Cuenta Bancaria debe tener 10 dígitos", vbExclamation
'        PonerFoco txtC
'        Exit Function
'    End If
'    ValidarCuentaBancaria = True
End Function



' ### [Monica] 02/10/2006 de laura
Public Function SituarRSetMULTI(ByRef vData As ADODB.Recordset, vWhere As String) As Boolean
'Situa un ADODB.Recordset en el registo que cumple vwhere
On Error GoTo ESituarData
    
        'El sql para que se situe en el registro en especial es el siguiente
        Multi_Find2 vData, vWhere
        If vData.EOF Or vData.BOF Then GoTo ESituarData
        
        SituarRSetMULTI = True
        Exit Function
ESituarData:
        If Err.Number <> 0 Then Err.Clear
        SituarRSetMULTI = False
End Function


' ### [Monica] 10/10/2006 de laura
Public Sub CargarProgres(ByRef PBar As ProgressBar, Valor As Integer)
On Error Resume Next
    PBar.Max = 100
    PBar.Value = 0
    PBar.Tag = Valor
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub IncrementarProgres(ByRef PBar As ProgressBar, Veces As Integer)
On Error Resume Next
    PBar.Value = PBar.Value + ((Veces * PBar.Max) / CInt(PBar.Tag))
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub CargarProgresNew(ByRef PBar As ProgressBar, Valor As Integer)
On Error Resume Next
    PBar.Max = Valor
    PBar.Value = 0
'    PBar.Tag = Valor
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub IncrementarProgresNew(ByRef PBar As ProgressBar, Veces As Integer)
On Error Resume Next
'    PBar.Value = PBar.Value + ((Veces * PBar.Max) / CInt(PBar.Tag))
    PBar.Value = PBar.Value + Veces
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function EsNumerico2(Texto As String) As Boolean
Dim L As Long
    
    On Error GoTo eEsnumerico2
    
    EsNumerico2 = True
    
    L = CLng(Texto)
    
eEsnumerico2:
    If Err.Number <> 0 Then
        EsNumerico2 = False
    End If
End Function

' viene del ariges
Public Function ObtenerAlto(ByRef vDataGrid As DataGrid, Optional alto As Integer) As Single
Dim anc As Single
    anc = vDataGrid.Top + alto
    If vDataGrid.Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + vDataGrid.RowTop(vDataGrid.Row)
    End If
    ObtenerAlto = anc
End Function

Public Sub CancelaADODC(ByRef vData As Adodc)
On Error Resume Next
    vData.Recordset.Cancel
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub ActualizarToolbarGnral(ByRef Toolbar1 As Toolbar, Modo As Byte, Kmodo As Byte, posic As Byte)
'Modo: Modo antiguo
'Kmodo: Modo que se va a poner
Dim b As Boolean
    
    b = (Modo = 5 Or Modo = 6 Or Modo = 7)
    
    If (b) And (Kmodo <> 5 And Kmodo <> 6 And Kmodo <> 7) Then 'Cabecera
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(posic).Image = 3
        Toolbar1.Buttons(posic).ToolTipText = "Nuevo"
        '-- Modificar
        Toolbar1.Buttons(posic + 1).Image = 4
        Toolbar1.Buttons(posic + 1).ToolTipText = "Modificar"
        '-- eliminar
        Toolbar1.Buttons(posic + 2).Image = 5
        Toolbar1.Buttons(posic + 2).ToolTipText = "Eliminar"
    End If
    If (Kmodo = 5 Or Kmodo = 6 Or Kmodo = 7) Then 'Lineas
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(posic).Image = 3 '12
        Toolbar1.Buttons(posic).ToolTipText = "Nueva linea"
        '-- Modificar
        Toolbar1.Buttons(posic + 1).Image = 4 '13
        Toolbar1.Buttons(posic + 1).ToolTipText = "Modificar linea"
        '-- eliminar
        Toolbar1.Buttons(posic + 2).Image = 5 '14
        Toolbar1.Buttons(posic + 2).ToolTipText = "Eliminar linea"
    End If
End Sub

Public Function ObtenerCadKey(actCampo As Integer, sigCampo As Integer) As Integer
    Dim cadkey As Integer

    On Error Resume Next
    
    If actCampo > sigCampo Then
        cadkey = 38 'flecha superior
    Else
        cadkey = 40 'flecha inferior
    End If
    If sigCampo = 0 Then cadkey = 0
    
    ObtenerCadKey = cadkey
    
    If Err.Number <> 0 Then Err.Clear
End Function



Public Sub BloquearbtnBuscar(ByRef formulario As Form, Modo As Byte, Optional ModoLineas As Byte, Optional nomframe As String)
'Bloquea controles q sean ComboBox si no estamos en Modo: 3.-Insertar, 4.-Modificar
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualización
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar, Insertar/Modificar Lineas...)
Dim b As Boolean
On Error Resume Next

'    b = (Modo = 3 Or Modo = 4 Or Modo = 1 Or (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2)))
    b = (Modo = 3 Or Modo = 4 Or Modo = 1) Or (Modo = 5 And ModoLineas = 2)
    
    With formulario
        For i = 0 To .btnBuscar.Count - 1
            If formulario.btnBuscar(i).Container.Name = nomframe Then
                .btnBuscar(i).Enabled = Not b
                .btnBuscar(i).visible = Not b
            Else
                .btnBuscar(i).Enabled = False
                .btnBuscar(i).visible = False
            End If
        Next i
    End With
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function ValorCombo(ByRef Cbo As ComboBox) As Integer
'obtiene el valor del combo de la posicion en la q se encuentra

    On Error GoTo EValCombo
    
    If Cbo.ListIndex < 0 Then
        ValorCombo = -1
    Else
        ValorCombo = Cbo.ItemData(Cbo.ListIndex)
    End If
    Exit Function

EValCombo:
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function TextoCombo(ByRef Cbo As ComboBox) As String
'obtiene la descripcion del combo de la posicion en la q se encuentra

    On Error GoTo ErrTexCombo
    
    If Cbo.ListIndex < 0 Then
        TextoCombo = ""
    Else
        TextoCombo = Cbo.List(Cbo.ListIndex)
    End If
    Exit Function

ErrTexCombo:
    If Err.Number <> 0 Then Err.Clear
End Function

Public Sub ConseguirfocoChk(Modo As Byte)
     If Modo = 0 Or Modo = 2 Then
        KEYpressGnral 13, Modo, False
    End If
End Sub

