Attribute VB_Name = "ModContadores"
Option Explicit




'Public Function ObtenerContadorExpGrp(cadFec As String, cadCont As String) As Boolean
''Obtiene el contador para el expediente de grupos
''(IN) cadFec: fecha del expediente
''(OUT) cadCont: contador correspondiente
'
'
''    Dim cadFec As String
'    Dim vCont As CContador
'    Dim b As Boolean
'
''    cadFec = Trim(Text1(1).Text)
'    cadFec = Trim(cadFec)
'
'    If cadFec = "" Then
'        MsgBox "El campo fecha debe tener valor para obtener un contador.", vbExclamation
'        b = False
'    Else
'        Set vCont = New CContador
'        If vCont.ConseguirContador(cadFec, "expgrp_a", "expgrp_b", True) Then
'            If vCont.AnyoActual Then
''                txtAux(46).Text = vCont.Contador1
'                cadCont = vCont.Contador1
'            Else
''                txtAux(46).Text = vCont.Contador2
'                cadCont = vCont.Contador2
'            End If
'            b = True
'        End If
'        Set vCont = Nothing
'    End If
'
'    ObtenerContadorExpGrp = b
'End Function
'





'Public Function ObtenerContadorExpInd(cadFec As String, cadCont As String) As Boolean
''Obtiene el contador para el expediente de individuales
''(IN) cadFec: fecha del expediente
''(OUT) cadCont: contador correspondiente
'
'    Dim vCont As CContador
'    Dim b As Boolean
'
'    cadFec = Trim(cadFec)
'
'    If cadFec = "" Then
'        MsgBox "El campo fecha debe tener valor para obtener un contador.", vbExclamation
'        b = False
'    Else
'        Set vCont = New CContador
'        If vCont.ConseguirContador(cadFec, "expind_a", "expind_b", True) Then
'            If vCont.AnyoActual Then
'                cadCont = vCont.Contador1
'            Else
'                cadCont = vCont.Contador2
'            End If
'            b = True
'        End If
'        Set vCont = Nothing
'    End If
'
'    ObtenerContadorExpInd = b
'End Function



'Public Function ObtenerContadorVenta(cadFec As String, cadCont As String) As Boolean
''Obtiene el contador para una venta
''(IN) cadFec: fecha de la venta
''(OUT) cadCont: contador correspondiente
'
'    Dim vCont As CContador
'    Dim b As Boolean
'
'    cadFec = Trim(cadFec)
'
'    If cadFec = "" Then
'        MsgBox "El campo fecha de la venta debe tener valor para obtener un contador.", vbExclamation
'        b = False
'    Else
'        Set vCont = New CContador
'        If vCont.ConseguirContador(cadFec, "pventa_a", "pventa_b", True) Then
'            If vCont.AnyoActual Then
'                cadCont = vCont.Contador1
'            Else
'                cadCont = vCont.Contador2
'            End If
'            b = True
'        End If
'        Set vCont = Nothing
'    End If
'
'    ObtenerContadorVenta = b
'End Function




'Private Function ObtenerContador() As Boolean
''Obtiene el contador para la venta
'Dim cadFec As String
'Dim vCont As CContador
'Dim b As Boolean
'
'        cadFec = Trim(Text1(1).Text)
'        If cadFec = "" Then
'            MsgBox "El campo fecha de la venta debe tener valor para obtener un contador.", vbExclamation
'            b = False
'        Else
'            Set vCont = New CContador
'            If vCont.ConseguirContador(cadFec, "pventa_a", "pventa_b", True) Then
'                If vCont.AnyoActual Then
'                    Text1(0).Text = vCont.Contador1
'                Else
'                    Text1(0).Text = vCont.Contador2
'                End If
'                FormateaCampo Text1(0)
'                b = True
'            End If
'            Set vCont = Nothing
'        End If
'
'        ObtenerContador = b
'End Function
