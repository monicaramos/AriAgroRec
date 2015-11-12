Attribute VB_Name = "ModFechas"
Option Explicit


'=== DAVID (estaban en Modulo:bus) (NO LA USO!!!)
Public Function DiasMes(mes As Byte, Anyo As Integer) As Integer
    Select Case mes
    Case 2
        If (Anyo Mod 4) = 0 Then
            DiasMes = 29
        Else
            DiasMes = 28
        End If
    Case 1, 3, 5, 7, 8, 10, 12
        DiasMes = 31
    Case Else
        DiasMes = 30
    End Select
End Function


'=== DAVID (estaban en Modulo:bus)
'Public Function EsFechaOK(ByRef T As TextBox) As Boolean
''Dim cad As String
''
''    cad = T.Text
''    If InStr(1, cad, "/") = 0 Then
''        If Len(T.Text) = 8 Then
''            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
''        Else
''            If Len(T.Text) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
''        End If
''    End If
''
''    If IsDate(cad) Then
''        EsFechaOK = True
''        T.Text = Format(cad, "dd/mm/yyyy")
''    Else
''        EsFechaOK = False
''    End If
''EsFechaOK = EsFechaOKString
'End Function

'=== DAVID (estaban en Modulo:bus, antes era ESFechaOKString)
Public Function EsFechaOK(T As String) As Boolean
Dim cad As String
Dim mes As String, dia As String
    
    cad = T
    If InStr(1, cad, "/") = 0 Then
        
      '==== Anade: Laura 04/02/2005 =============
        If Len(cad) < 6 Then
            EsFechaOK = False
            Exit Function
        End If
        
        'Comprobar que el dia es correcto, valores entre 1-31
        dia = Mid(cad, 1, 2)
        If IsNumeric(dia) Then
            If dia < 1 Or dia > 31 Then
                EsFechaOK = False
                Exit Function
            End If
        Else
            EsFechaOK = False
            Exit Function
        End If
        
        'Comprobar que el mes es correcto, valores entre 1-12
        mes = Mid(cad, 3, 2)
        If IsNumeric(mes) Then
            If mes < 1 Or mes > 12 Then
                EsFechaOK = False
                Exit Function
            End If
        Else
            EsFechaOK = False
            Exit Function
        End If
      '============================================
        
        If Len(T) = 8 Then
            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        Else
            If Len(T) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        End If
    Else
        dia = Mid(cad, 1, 2)
        mes = Mid(cad, 4, 2)
    End If
    
    If IsDate(cad) Then
        EsFechaOK = True
        T = Format(cad, "dd/MM/yyyy")
      '==== Añade: Laura 08/02/2005
        If Month(T) <> Val(mes) Then EsFechaOK = False
        If Day(T) <> Val(dia) Then EsFechaOK = False
      '====
    Else
        EsFechaOK = False
    End If
End Function



'=== DAVID (estaba en Modulo:bus)
Public Function EsHoraOK(T As String) As Boolean
Dim cad As String
    
    cad = T
    If InStr(1, cad, ":") = 0 Then
        Select Case Len(T)
            Case 8
                cad = Mid(cad, 1, 2) & ":" & Mid(cad, 3, 2) & ":" & Mid(cad, 5)
            Case 6
                cad = Mid(cad, 1, 2) & ":" & Mid(cad, 3, 2) & ":" & Mid(cad, 5)
            Case 4
                cad = Mid(cad, 1, 2) & ":" & Mid(cad, 3, 2) & ":00"
        End Select
    End If
    
    If IsDate(cad) Then
        EsHoraOK = True
        T = Format(cad, "hh:mm:ss")
    Else
        EsHoraOK = False
    End If
End Function


'==== LAURA
Public Function PonerFormatoFecha(ByRef T As TextBox, Optional ComproFecCamp As Boolean) As Boolean
Dim cad As String

    cad = T.Text
    If cad <> "" Then
        If Not EsFechaOK(cad) Then
            MsgBox "Fecha incorrecta. (dd/MM/yyyy)", vbExclamation
            cad = "mal"
        End If
        If cad <> "" And cad <> "mal" Then
            T.Text = cad
            '++monica: comprobamos que todas las fechas se encuentran dentro de campaña
            '[Monica]28/08/2013: antes not NoComproFecCamp
            If ComproFecCamp Then FechaDentroDeCampanya T.Text
            '++
            PonerFormatoFecha = True
        Else
'                T.Text = ""
            PonerFoco T
        End If
    End If
End Function


'==== LAURA
Public Function PonerFormatoHora(ByRef T As TextBox) As Boolean
Dim cad As String

    cad = T.Text
    If cad <> "" Then
        If Not EsHoraOK(cad) Then
            MsgBox "Hora incorrecta. (hh:mm:ss)", vbExclamation
            cad = "mal"
        End If
        If cad <> "" And cad <> "mal" Then
            T.Text = cad
            PonerFormatoHora = True
        Else
            T.Text = ""
            PonerFoco T
        End If
    End If
End Function


'==== LAURA
Public Function EsFechaPosterior(FIni As String, FFin As String, MError As Boolean, Optional Men As String) As Boolean
'Comprueba que la Fecha Fin es posterior a la Fecha de Inicio
'Si se pasa un cadena Men, se muestra esta como Mensaje de Error
On Error Resume Next

    EsFechaPosterior = True
    If Trim(FIni) <> "" And Trim(FFin) <> "" Then
        If CDate(FIni) >= CDate(FFin) Then
            EsFechaPosterior = False
            If MError Then
                If Men <> "" Then
                    MsgBox Men, vbInformation
                Else
                    MsgBox "La Fecha Fin debe ser posterior a la Fecha Inicio", vbInformation
                End If
            End If
        Else
            EsFechaPosterior = True
        End If
    End If
End Function


'==== LAURA
Public Function EsFechaIgualPosterior(FIni As String, FFin As String, MError As Boolean, Optional Men As String) As Boolean
'Comprueba que la Fecha Fin es igual o posterior a la Fecha de Inicio
'Si se pasa un cadena Men, se muestra esta como Mensaje de Error
'OUT -> true: Ffin >= Fini
On Error Resume Next

'    EsFechaIgualPosterior = True
    If Trim(FIni) <> "" And Trim(FFin) <> "" Then
        If CDate(FIni) > CDate(FFin) Then
            EsFechaIgualPosterior = False
            If MError Then
                If Men <> "" Then
                    MsgBox Men, vbInformation
                Else
                    MsgBox "La Fecha Fin debe ser igual o posterior a la Fecha Inicio", vbInformation
                End If
            End If
        Else
            EsFechaIgualPosterior = True
        End If
    Else
        EsFechaIgualPosterior = True
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


'==== LAURA
Public Function EntreFechas(FIni As String, FechaComp As String, FFin As String) As Boolean
Dim b As Boolean
    b = False
    If FIni <> "" And FFin <> "" Then
        If EsFechaIgualPosterior(FIni, FechaComp, False) And EsFechaIgualPosterior(FechaComp, FFin, False) Then
            b = True
        End If
    ElseIf FIni = "" And FFin <> "" Then
        If EsFechaIgualPosterior(FechaComp, FFin, False) Then
            b = True
        End If
    ElseIf FIni <> "" And FFin = "" Then
        If EsFechaIgualPosterior(FIni, FechaComp, False) Then
            b = True
        End If
    End If
    EntreFechas = b
End Function

'==== LAURA
Public Function CalculaSemana(Fecha As Date) As Integer
    CalculaSemana = DatePart("ww", Fecha)
End Function


Public Function UltimoDiaMes(Fecha As String, udia As Byte) As String
Dim F As Date
Dim ultDia As String

    F = CDate(Fecha)
    If Day(F) = 31 Then 'Ya es el ultimo dia
        UltimoDiaMes = Fecha
        Exit Function
    ElseIf udia > 27 Then
        ultDia = udia & "/" & Month(F) & "/" & Year(F)
        If EsFechaOK(ultDia) Then
            UltimoDiaMes = ultDia
            Exit Function
        Else
            UltimoDiaMes = UltimoDiaMes(Fecha, udia - 1)
'            ultDia = "30" & "/" & Month(F) & "/" & Year(F)
'            If EsFechaOK(ultDia) Then
'                UltimoDiaMes = ultDia
'                Exit Function
'            Else
'
'            End If
        End If
    End If

End Function

'++monica
Public Function FechaDentroDeCampanya(fec As String) As Boolean
Dim b As Boolean
Dim cad As String
Dim Mens As String
    On Error GoTo Err3
    
    FechaDentroDeCampanya = True
    b = ((CDate(fec) >= CDate(vParam.FecIniCam)) And (CDate(fec) <= CDate(vParam.FecFinCam)))
    If Not b Then
        cad = "La Fecha introducida no se encuentra dentro de Campaña. Revise." '& vbCrLf & vbCrLf & "    Fecha inicio campaña: "
'        cad = cad & Format(vParam.FecIniCam, "dd/mm/yyyy") & vbCrLf
'        cad = cad & "    Fecha fin de campaña: " & Format(vParam.FecFinCam, "dd/mm/yyyy") & vbCrLf
        
        MsgBox cad, vbExclamation
    End If
    Exit Function
    
Err3:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    FechaDentroDeCampanya = False

End Function

Public Function EsCampanyaActual(bd As String) As Boolean
Dim SQL As String

    SQL = "select usuarios.empresasariagro.usuario from usuarios.empresasariagro "
    SQL = SQL & " where ariagro = " & DBSet(bd, "T")
    
    EsCampanyaActual = (DevuelveValor(SQL) = "*")

End Function

