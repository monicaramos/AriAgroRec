Attribute VB_Name = "ModHistorico"
Option Explicit
'Modulo para el traspaso de registros de cabecera y lineas de las tablas
'de OFERTAS,PEDIDOS,ALBARANES
'A las tablas del HISTORICO de Ofertas,Pedidos,Albaranes
'OFERTAS:
' scapre --> schpre
' slipre --> slhpre
'PEDIDOS:
' scaped --> schped
' sliped --> slhped


Dim CodTipoMov As String
Dim NomTabla As String 'nombre de la tabla
Dim NomTablaH As String 'nombre de la tabla del historico al que movemos
Dim NomTablaLin As String 'nombre tabla de lineas
Dim NomTablaLinH As String 'nombre tabla de lineas del historico


Public Function ActualizarElTraspaso(ByRef ADonde As String, cadWHERE As String, codMovim As String, Optional cadL As String) As Boolean
'codMovim: tipo de movimiento que estamos hacienco: OFE,PEV,ALV,PEC,ALC,....
    
    ActualizarElTraspaso = False
    CodTipoMov = codMovim
    
    'Insertamos en cabeceras Historico
    ADonde = "Insertando datos en histórico cabeceras "
    If Not InsertarCabeceraHistorico(cadWHERE, cadL) Then Exit Function
'    IncrementarProgres 2
     
    'Insertamos en lineas Historico
    ADonde = "Insertando datos en Histórico lineas "
    If Not InsertarLineasHistorico(cadWHERE) Then Exit Function
'    IncrementarProgres 2
    
    'Borramos cabeceras y lineas
    ADonde = "Borrar cabeceras y lineas"
    If Not BorrarTraspaso(False, cadWHERE) Then Exit Function
'    IncrementarProgres 2

    ActualizarElTraspaso = True
End Function


Private Function InsertarCabeceraHistorico(cadWHERE As String, Optional cadL As String) As Boolean
Dim SQL As String
On Error Resume Next

    Select Case CodTipoMov
      Case "PEV" 'pedidos de venta a clientes
        NomTabla = "scaped"
        NomTablaH = "schped"
        SQL = " SELECT numpedcl,fecpedcl," & vUsu.Codigo Mod 1000 & " as codigusu," & cadL & ","
        SQL = SQL & "fecentre,sementre,visadore,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
        SQL = SQL & "coddirec,nomdirec,referenc,codtraba,codagent,codforpa,dtoppago,dtognral,"
        SQL = SQL & "tipofact,observa01,observa02,observa03,observa04,observa05,servcomp,restoped,numofert,fecofert,observap1,observap2,recogecl"
        
      Case "ALV", "ALM" 'Albaran de venta a clientes
        NomTabla = "scaalb"
        NomTablaH = "schalb"
        SQL = " SELECT codtipom,numalbar,fechaalb," & vUsu.Codigo Mod 1000 & " as codigusu," & cadL & ","
        SQL = SQL & "factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
        SQL = SQL & "coddirec,nomdirec,referenc,facturkm,cantidkm,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,"
        SQL = SQL & "tipofact,observa01,observa02,observa03,observa04,observa05,numofert,fecofert,numpedcl,fecpedcl,fecentre,sementre,esticket"
        
        
      Case "OFE" 'Ofertas a Clientes
        NomTabla = "scapre"
        NomTablaH = "schpre"
        SQL = " SELECT numofert, fecofert," & "'" & Format(Now, FormatoFecha) & "' as fechamov, fecentre, aceptado, codclien, nomclien, domclien, codpobla, "
        SQL = SQL & "pobclien, proclien, nifclien, telclien, coddirec, nomdirec, referenc, codtraba, codagent, codforpa, dtoppago, dtognral, tipofact, "
        SQL = SQL & "plazos01, plazos02, plazos03, asunto01, asunto02, asunto03, asunto04, asunto05, observa01, observa02, observa03, observa04, observa05, "
        SQL = SQL & "concepto, seguiofe "
        
      Case "ALC" 'Albaranes a Proveedores (Compras)
        NomTabla = "scaalp"
        NomTablaH = "schalp"
        SQL = " SELECT numalbar,fechaalb,codprove," & vUsu.Codigo Mod 1000 & " as codigusu," & cadL & ","
        SQL = SQL & "nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove,"
        SQL = SQL & "codforpa,dtoppago,dtognral,"
'        SQL = SQL & "codforpa,codtraba,codtrab1,dtoppago,dtognral,"
        SQL = SQL & "observa1,observa2,observa3,observa4,observa5,numpedpr,fecpedpr"
      
      Case "PEC" 'Pedidos a Proveedores (Compras)
        NomTabla = "scappr"
        NomTablaH = "schppr"
        SQL = " SELECT numpedpr,fecpedpr," & vUsu.Codigo Mod 1000 & " as codigusu," & cadL & ","
        SQL = SQL & "codprove,nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove,"
'        SQL = SQL & "coddirea,coddiref,codforpa,codtraba,codtrab1,dtognral,dtoppago,"
        SQL = SQL & "coddirea,coddiref,codforpa,dtognral,dtoppago,"
        SQL = SQL & "restoped,codclien,observa1,observa2,observa3,observa4,observa5,tipoporte"
      
    End Select
    
    SQL = SQL & " FROM " & NomTabla & " WHERE " & cadWHERE
    SQL = "INSERT INTO " & NomTablaH & SQL
    
    Conn.Execute SQL
    
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarCabeceraHistorico = False
    Else
        InsertarCabeceraHistorico = True
    End If
End Function


Private Function InsertarLineasHistorico(cadWHERE As String) As Boolean
Dim SQL As String
On Error Resume Next

    Select Case CodTipoMov
      Case "PEV" 'pedidos ventas a clientes
        NomTablaLin = "sliped"
        NomTablaLinH = "slhped"
        SQL = " SELECT scaped.numpedcl,scaped.fecpedcl,sliped.numlinea,sliped.codalmac,sliped.codartic,sliped.nomartic,sliped.ampliaci,sliped.cantidad,servidas,precioar,dtoline1,dtoline2,importel,origpre "
        SQL = SQL & " FROM scaped INNER JOIN sliped on scaped.numpedcl=sliped.numpedcl "
        SQL = SQL & " WHERE " & cadWHERE
        
      Case "ALV", "ALM" 'Albaranes ventas a clientes
        NomTablaLin = "slialb"
        NomTablaLinH = "slhalb"
        SQL = " SELECT scaalb.codtipom,scaalb.numalbar,scaalb.fechaalb,slialb.numlinea,slialb.codalmac,slialb.codartic,slialb.nomartic,slialb.ampliaci,slialb.cantidad,precioar,dtoline1,dtoline2,importel,origpre "
        SQL = SQL & " FROM scaalb INNER JOIN slialb on scaalb.codtipom=slialb.codtipom AND scaalb.numalbar=slialb.numalbar "
        SQL = SQL & " WHERE " & cadWHERE
        
      Case "OFE" 'Ofertas a clientes
        NomTablaLin = "slipre"
        NomTablaLinH = "slhpre"
        SQL = " SELECT scapre.numofert,scapre.fecofert,slipre.numlinea,slipre.codalmac,slipre.codartic,slipre.nomartic,slipre.ampliaci,slipre.cantidad,precioar,dtoline1,dtoline2,importel,origpre "
        SQL = SQL & " FROM scapre INNER JOIN slipre on scapre.numofert=slipre.numofert"
        SQL = SQL & " WHERE " & cadWHERE
        
      Case "ALC" 'Albaranes compras a proveedores
        NomTablaLin = "slialp"
        NomTablaLinH = "slhalp"
        SQL = " SELECT scaalp.numalbar,scaalp.fechaalb,scaalp.codprove,slialp.numlinea,slialp.codartic,slialp.codalmac,slialp.nomartic,slialp.ampliaci,slialp.cantidad,precioar,dtoline1,dtoline2,importel "
        SQL = SQL & " FROM scaalp INNER JOIN slialp on scaalp.numalbar=slialp.numalbar AND scaalp.fechaalb=slialp.fechaalb AND scaalp.codprove=slialp.codprove "
        SQL = SQL & " WHERE " & cadWHERE
      
      Case "PEC" 'Pedidos compras a proveedores
        NomTablaLin = "slippr"
        NomTablaLinH = "slhppr"
        SQL = " SELECT scappr.numpedpr,scappr.fecpedpr,slippr.numlinea,slippr.codartic,slippr.codalmac,slippr.nomartic,slippr.ampliaci,slippr.cantidad,slippr.recibida,precioar,dtoline1,dtoline2,importel "
        SQL = SQL & " FROM scappr INNER JOIN slippr on scappr.numpedpr=slippr.numpedpr "
        SQL = SQL & " WHERE " & cadWHERE
      
    End Select
    
    SQL = "INSERT INTO " & NomTablaLinH & SQL
    
    Conn.Execute SQL
    
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarLineasHistorico = False
    Else
        InsertarLineasHistorico = True
    End If
End Function



Private Function BorrarTraspaso(EnHistorico As Boolean, cadWHERE As String) As Boolean
'Si EnHistorico=true borra de las tablas de historico: "schtra" y "slhtra"
'Si EnHistorico=false borra de las tablas de traspaso: "scatra" y "slitra"
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Cad As String, cadAux As String

    BorrarTraspaso = False
    On Error GoTo EBorrar
    
    
    'Eliminamos las lineas
    Select Case CodTipoMov
      Case "PEV" 'pedidos ventas  a clientes
        SQL = "Select numpedcl from scaped WHERE " & cadWHERE
        cadAux = " numpedcl IN "
      Case "ALV", "ALM" 'albaranes ventas a clientes
        SQL = "Select numalbar from scaalb WHERE " & cadWHERE
        cadAux = "codtipom=" & DBSet(CodTipoMov, "T") & " AND numalbar IN "
      Case "OFE" 'Ofertas a clientes
        SQL = "Select numofert from scapre WHERE " & cadWHERE
        cadAux = " numofert IN "
      Case "ALC" 'Albaranes compras a proveedores
'        SQL = "Select numalbar,fechaalb,codprove from scaalp WHERE " & cadWHERE
'        cadAux = " numalbar IN "
    End Select
    
    If CodTipoMov <> "ALC" And CodTipoMov <> "PEC" Then
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not RS.EOF
            If CodTipoMov <> "ALC" Then
                Cad = Cad & RS.Fields(0).Value & ","
            Else
                Cad = Cad & "numalbar="
            End If
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        'Quitar la ultima coma de la cadena
        Cad = Mid(Cad, 1, Len(Cad) - 1)
        
        cadAux = cadAux & "(" & Cad & ")"
    Else
        cadAux = Replace(cadWHERE, NomTabla, NomTablaLin)
    End If
    
    SQL = "DELETE FROM " & NomTablaLin & " WHERE " & cadAux

    Conn.Execute SQL
    
    'La cabecera
    SQL = "Delete from " & NomTabla
    SQL = SQL & " WHERE " & cadWHERE
    Conn.Execute SQL
    BorrarTraspaso = True
    
EBorrar:
    If Err.Number <> 0 Then
        BorrarTraspaso = False
    Else
        BorrarTraspaso = True
    End If
End Function



'========================================================

Public Sub CargarTagsHco(ByRef F As Form, vTabla As String, vTablaHco As String)
'Sustituye en los tags del formulario la tabla de Reparaciones (scarep)
'por la del historico de Reparaciones (schrep)
Dim Control As Object
Dim vtag As String

    For Each Control In F.Controls
        If Control.Tag <> "" Then
            vtag = Control.Tag
'            vtag = SustituirCadenas(vtag, vTabla, vTablaHco)
            vtag = Replace(vtag, vTabla, vTablaHco)
            Control.Tag = vtag
        End If
    Next Control
End Sub
