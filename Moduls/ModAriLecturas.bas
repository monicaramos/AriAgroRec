Attribute VB_Name = "ModAriLecturas"
Public Function ComprobarCero(Valor As String) As String
    If Valor = "" Then
        ComprobarCero = "0"
    Else
        ComprobarCero = Valor
    End If
End Function


Public Function DevuelveValor(vSQL As String) As Variant
'Devuelve el valor de la SQL
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
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

