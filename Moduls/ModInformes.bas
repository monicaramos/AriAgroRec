Attribute VB_Name = "ModInformes"
Option Explicit


'==============================================================
'====== FUNCIONES GENERALES  PARA INFORMES ====================

'Esta funcion lo que hace es genera el valor del campo
'El campo lo coge del recordset, luego sera field(i), y el tipo es para añadirle
'las coimllas, o quitarlas comas
'  Si es numero viene un 1 si no nada
'## NO LA USO, UTILIZO DBSET
'Public Function ParaBD(ByRef campo As ADODB.Field, Optional EsNumerico As Byte) As String
'
'    If IsNull(campo) Then
'        ParaBD = "NULL"
'    Else
'        Select Case EsNumerico
'        Case 1
'            ParaBD = TransformaComasPuntos(CStr(campo))
'        Case 2
'            'Fechas
'            ParaBD = "'" & Format(CStr(campo), "dd/MM/yyyy") & "'"
'        Case Else
'            ParaBD = "'" & campo & "'"
'        End Select
'    End If
'    ParaBD = "," & ParaBD
'End Function

Public Sub AbrirListadoPOZ(numero As Byte)
    Screen.MousePointer = vbHourglass
    frmPOZListado.OpcionListado = numero
    frmPOZListado.Show vbModal
    Screen.MousePointer = vbDefault
End Sub


Public Sub AbrirListadoBodEntradas(numero As Byte)
    Screen.MousePointer = vbHourglass
    frmBodListEntradas.OpcionListado = numero
    frmBodListEntradas.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Public Sub AbrirListadoBodAnticipos(numero As Byte)
    Screen.MousePointer = vbHourglass
    frmBodListAnticipos.OpcionListado = numero
    frmBodListAnticipos.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Public Sub AbrirListadoAnticipos(numero As Byte)
    Screen.MousePointer = vbHourglass
    frmListAnticipos.OpcionListado = numero
    frmListAnticipos.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Public Sub AbrirListadoTomaDatos(numero As Byte)
    Screen.MousePointer = vbHourglass
    frmListTomaDatos.OpcionListado = numero
    frmListTomaDatos.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Public Sub AbrirListadoTrazabilidad(numero As Byte)
    Screen.MousePointer = vbHourglass
    frmListTrazabilidad.OpcionListado = numero
    frmListTrazabilidad.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Public Sub AbrirListadoTraza(numero As Byte)
    Screen.MousePointer = vbHourglass
    frmListTraza.OpcionListado = numero
    frmListTraza.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Public Sub AbrirListadoAPOR(numero As Byte)
    Screen.MousePointer = vbHourglass
    frmAPOListados.OpcionListado = numero
    frmAPOListados.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Public Sub AbrirListado(numero As Byte)
    Screen.MousePointer = vbHourglass
    frmListado.OpcionListado = numero
    frmListado.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Public Sub AbrirListadoFVarias(numero As Byte)
    Screen.MousePointer = vbHourglass
    frmFVARListados.OpcionListado = numero
    frmFVARListados.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Public Sub AbrirListadoNominas(numero As Byte)
    Screen.MousePointer = vbHourglass
    frmListNomina.OpcionListado = numero
    frmListNomina.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Public Sub AbrirListadoOfer(numero As Integer)
'Abre el Form con los listados de Ofertas
    Screen.MousePointer = vbHourglass
    frmListadoOfer.OpcionListado = numero
    frmListadoOfer.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Public Sub AbrirListadoADV(numero As Byte)
    Screen.MousePointer = vbHourglass
    frmADVListados.OpcionListado = numero
    frmADVListados.Show vbModal
    Screen.MousePointer = vbDefault
End Sub


Public Function AnyadirAFormula(ByRef cadFormula As String, arg As String) As Boolean
'Concatena los criterios del WHERE para pasarlos al Crystal como FormulaSelection
    If arg = "Error" Then
        AnyadirAFormula = False
        Exit Function
    ElseIf arg <> "" Then
        If cadFormula <> "" Then
            cadFormula = cadFormula & " AND " & arg
        Else
            cadFormula = arg
        End If
    End If
    AnyadirAFormula = True
End Function


Public Function RegistrosAListar(vSQL As String, Optional vBD As Byte) As Byte
'Devuelve si hay algun registro para mostrar en el Informe con la seleccion
'realizada. Si no hay nada que mostrar devuelve 0 y no abrirá el informe
Dim Rs As ADODB.Recordset

    On Error Resume Next
    
    Set Rs = New ADODB.Recordset
    If vBD = cConta Then
        Rs.Open vSQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Else
        Rs.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    End If

    
    RegistrosAListar = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then RegistrosAListar = 1 'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        RegistrosAListar = 0
        Err.Clear
    End If
End Function




Public Function HayRegParaInforme(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    If RegistrosAListar(Sql) = 0 Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegParaInforme = False
    Else
        HayRegParaInforme = True
    End If
End Function


Public Function HayRegParaInformeNew(cTabla As String, cWhere As String, ctabla1 As String, cwhere1 As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select count(*) numero FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    Sql = Sql & " union "
    
    ctabla1 = QuitarCaracterACadena(ctabla1, "{")
    ctabla1 = QuitarCaracterACadena(ctabla1, "}")
    Sql = Sql & "Select count(*) numero FROM " & QuitarCaracterACadena(ctabla1, "_1")
    If cwhere1 <> "" Then
        cwhere1 = QuitarCaracterACadena(cwhere1, "{")
        cwhere1 = QuitarCaracterACadena(cwhere1, "}")
        cwhere1 = QuitarCaracterACadena(cwhere1, "_1")
        Sql = Sql & " WHERE " & cwhere1
    End If
    Sql = "select sum(numero) from (" & Sql & ") aaaaaaa"
    If DevuelveValor(Sql) = 0 Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegParaInformeNew = False
    Else
        HayRegParaInformeNew = True
    End If
End Function





Public Function CadenaDesdeHasta(cadDesde As String, cadhasta As String, campo As String, TipoCampo As String, Optional nomCampo As String) As String
'Devuelve la cadena de seleccion: " (campo >= cadDesde and campo<=cadHasta) "
'para Crystal Report
Dim cadAux As String

    If Trim(cadDesde) = "" And Trim(cadhasta) = "" Then
        'Campo Desde y Hasta no tiene valor
            cadAux = ""
    Else
        'Campo DESDE
        If cadDesde <> "" Then
            Select Case TipoCampo
                Case "N"
                    cadAux = campo & " >= " & Val(cadDesde)
                Case "T"
                    cadAux = campo & " >= """ & cadDesde & """"
                Case "F"
                    cadAux = campo & " >= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")"
            End Select
        End If
        
        'Campo HASTA
        If cadhasta <> "" Then
            If cadAux <> "" Then 'Hay campo Desde y campo Hasta
                'Comprobar Desde <= Hasta
                Select Case TipoCampo
                    Case "N"
                        If CSng(cadDesde) > CSng(cadhasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= " & Val(cadhasta)
                        End If
                        
                    Case "T"
                        If cadDesde > cadhasta Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= """ & cadhasta & """"
                        End If
                    
                    Case "F"
                        If CDate(cadDesde) > CDate(cadhasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= Date(" & Year(cadhasta) & "," & Month(cadhasta) & "," & Day(cadhasta) & ")"
                        End If
                End Select
            Else 'No hay campo Desde. Solo hay campo Hasta
                Select Case TipoCampo
                    Case "N"
                        cadAux = campo & " <= " & Val(cadhasta)
                    Case "T"
                        cadAux = campo & " <= """ & cadhasta & """"
                    Case "F"
                        cadAux = campo & " <= Date(" & Year(cadhasta) & "," & Month(cadhasta) & "," & Day(cadhasta) & ")"
                End Select
            End If
        End If
    End If
    If cadAux <> "" And cadAux <> "Error" Then cadAux = "(" & cadAux & ")"
    CadenaDesdeHasta = cadAux
End Function


Public Function CadenaDesdeHastaBD(cadDesde As String, cadhasta As String, campo As String, TipoCampo As String) As String
'Devuelve la cadena de seleccion: " (campo >= valor1 and campo<=valor2) "
'Para MySQL
Dim cadAux As String

    If Trim(cadDesde) = "" And Trim(cadhasta) = "" Then
        'Campo Desde y Hasta no tiene valor
            cadAux = ""
    Else
        'Campo DESDE
        If cadDesde <> "" Then
            Select Case TipoCampo
                Case "N"
                    cadAux = campo & " >= " & Val(cadDesde)
                Case "T"
                    cadAux = campo & " >= """ & cadDesde & """"
                Case "F"
                    cadAux = "(" & campo & " >= '" & Format(cadDesde, FormatoFecha) & "')"
            End Select
        End If
        
        'Campo HASTA
        If cadhasta <> "" Then
            If cadAux <> "" Then 'Hay campo Desde y campo Hasta
                'Comprobar Desde <= Hasta
                Select Case TipoCampo
                    Case "N"
                        If CSng(cadDesde) > CSng(cadhasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= " & Val(cadhasta)
                        End If
                        
                    Case "T"
                        If CSng(cadDesde) > CSng(cadhasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= """ & cadhasta & """"
                        End If
                    
                    Case "F"
                        If CDate(cadDesde) > CDate(cadhasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and (" & campo & " <= '" & Format(cadhasta, FormatoFecha) & "')"
                        End If
                End Select
                
            Else 'No hay campo Desde. Solo hay campo Hasta
                Select Case TipoCampo
                    Case "N"
                        cadAux = campo & " <= " & Val(cadhasta)
                    Case "T"
                        cadAux = campo & " <= """ & cadhasta & """"
                    Case "F"
                        cadAux = campo & " <= '" & Format(cadhasta, FormatoFecha) & "'"
                End Select
            End If
        End If
    End If
    If cadAux <> "" And cadAux <> "Error" Then cadAux = "(" & cadAux & ")"
    CadenaDesdeHastaBD = cadAux
End Function


Public Function AnyadirParametroDH(param As String, codD As String, codH As String, nomD As String, nomH As String) As String
On Error Resume Next
    
    If codD <> "" Then
        param = param & "DESDE: " & codD
        If nomD <> "" Then param = param & " - " & Replace(nomD, """", """""") 'nomD
    End If
    If codH <> "" Then
        param = param & "  HASTA: " & codH
        If nomH <> "" Then param = param & " - " & Replace(nomH, """", """""") 'nomH
    End If
    
    AnyadirParametroDH = param & """|"
    If Err.Number <> 0 Then Err.Clear
End Function



Public Function QuitarCaracterACadena(cadForm As String, Caracter As String) As String
'IN: [cadForm] es la cadena en la que se eliminara todos los caractes iguales a la vble [Caracter]
'OUT: cadena sin los caracteres
'EJEMPLO: "{scaalb.numalbar}", "{"  -->  "scaalb.numalbar}"
Dim i As Integer
Dim J As Integer
Dim Aux As String

    Aux = cadForm
    i = InStr(1, Aux, Caracter, vbTextCompare)
    While i > 0
        i = InStr(1, Aux, Caracter, vbTextCompare)
        If i > 0 Then
            J = Len(Caracter)
            Aux = Mid(Aux, 1, i - 1) & Mid(Aux, i + J, Len(Aux) - 1)
        End If
    Wend
    QuitarCaracterACadena = Aux
End Function


Public Function PonerParamRPT(Indice As Byte, CadParam As String, numParam As Byte, nomDocu As String, Optional EsAridoc As Boolean, Optional ImprimeDirecto As Integer) As Boolean
'EsAridoc = false usamos el nomdocum normal
'           true usamos el rpt para aridoc
'ImprimeDirecto = false usamos el crystal
'                 true usamos el print

Dim vParamRpt As CParamRpt 'Tipos de Documentos
Dim cad As String

    Set vParamRpt = New CParamRpt

    If vParamRpt.Leer(Indice) = 1 Then
        cad = "No se han podido cargar los Parámetros de Tipos de Documentos." & vbCrLf
        MsgBox cad & "Debe configurar la aplicación.", vbExclamation
        Set vParamRpt = Nothing
        PonerParamRPT = False
        Exit Function
    Else
        If CadParam = "" Then
            cad = "|"
        Else
            cad = ""
        End If
        cad = cad & "pCodigoISO=""" & vParamRpt.CodigoISO & """|"
        If vParamRpt.CodigoRevision = -1 Then
            cad = cad & "pCodigoRev=""" & "" & """|"
        Else
            cad = cad & "pCodigoRev=""" & Format(vParamRpt.CodigoRevision, "00") & """|"
        End If
        numParam = numParam + 2
        If vParamRpt.LineaPie1 <> "" Then
            cad = cad & "pLinea1=""" & vParamRpt.LineaPie1 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie2 <> "" Then
            cad = cad & "pLinea2=""" & vParamRpt.LineaPie2 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie3 <> "" Then
            cad = cad & "pLinea3=""" & vParamRpt.LineaPie3 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie4 <> "" Then
            cad = cad & "pLinea4=""" & vParamRpt.LineaPie4 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie5 <> "" Then
            cad = cad & "pLinea5=""" & vParamRpt.LineaPie5 & """|"
            numParam = numParam + 1
        End If
        CadParam = CadParam & cad
        If Not EsAridoc Then
            nomDocu = vParamRpt.Documento
        Else
            nomDocu = vParamRpt.AridocRpt
        End If
        
        ImprimeDirecto = vParamRpt.ImprimeDirecto
        
        PonerParamRPT = True
        Set vParamRpt = Nothing
    End If
End Function


Public Sub PonerFrameVisible(ByRef vFrame As Frame, visible As Boolean, H As Integer, W As Integer)
'Pone el Frame Visible y Ajustado al Formulario, y visualiza los controles
    
        vFrame.visible = visible
        If visible = True Then
            'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
            vFrame.Top = -90
            vFrame.Left = 0
            vFrame.Width = W
            vFrame.Height = H
        End If
End Sub


Public Function PonerParamEmpresa(CadParam As String, numParam As Byte) As Boolean
Dim DomiEmp As String
Dim WebEmp As String
Dim cad As String

        DomiEmp = vParam.DomicilioEmpresa & " - " & vParam.CPostal & " " & vParam.Poblacion
        If vParam.Provincia <> vParam.Poblacion Then DomiEmp = DomiEmp & " " & vParam.Provincia
        DomiEmp = DomiEmp & " - Telf. " & vParam.Telefono & " - Fax. " & vParam.Fax
        WebEmp = "Internet: " & vParam.WebEmpresa & " - E-mail: " & vParam.MailEmpresa
        'Resto parametros
        cad = ""
        cad = cad & "pNomEmpre=""" & vParam.NombreEmpresa & """|"
        cad = cad & "pDomEmpre=""" & DomiEmp & """|"
        cad = cad & "pWebEmpre=""" & WebEmp & """|"
        
        numParam = numParam + 3
        CadParam = CadParam & cad
        PonerParamEmpresa = True
End Function

Public Function SaltosDeLinea(ByVal cadena As String) As String
    Dim Devu As String
    Dim i As Integer
    
    Devu = ""
    Do
        i = InStr(1, cadena, vbCrLf)
        If i > 0 Then
            If Devu <> "" Then Devu = Devu & """ + chr(13) + """
            Devu = Devu & Mid(cadena, 1, i - 1)
            cadena = Mid(cadena, i + 2)
            
       Else
            Devu = Devu & cadena
       End If
    Loop While i > 0
    SaltosDeLinea = Devu
End Function

