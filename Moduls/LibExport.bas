Attribute VB_Name = "LibExport"
Option Explicit

Sub CargarTodosLosCampos()
    '-- Utilidad que carga todos los campos de la base de datos
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim cmp As GRPTC_Campo
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select a.*, b.codprodu from rcampos as a , variedades as b where b.codvarie = a.codvarie"
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            '-- creamos el objeto auxiliar que montará el XML de trazatec
            Set cmp = New GRPTC_Campo
            '-- vamos cargando los diferentes valores
            cmp.Codsocio = Rs!Codsocio
            cmp.codcampo = Rs!codcampo
            cmp.codprodu = Rs!codprodu
            cmp.codvarie = Rs!codvarie
            cmp.codparti = Rs!codparti
            cmp.hanegada = 0 ' no interesa en trazatec
            cmp.numarbol = 0 ' tampoco interesa
            cmp.Poligono = Rs!Poligono
            '-- Y ahora el objeto chivato para grabar
            Set chv = New GRPTC_Chivato
            chv.Id = 0 'ya lo montará en el momento de la grabación
            chv.BD_Org = "AGRO"
            '[Monica]16/11/2011: Solo si es Alzira es SCAMP1
            If vParamAplic.Cooperativa = 4 Then
                chv.Tabla = "SCAMP1"
            Else
                chv.Tabla = "SCAMPO"
            End If
            chv.Oper = "I"
            chv.Fecha = Format(Date, "dd/mm/yyyy")
            chv.Sep = "&"
            chv.Clv_Old = ""
            chv.Clv_New = CStr(cmp.codcampo)
            chv.XML = cmp.GenXML
            chv.Grabar
            Rs.MoveNext
        Wend
    End If
End Sub

Sub CargarUnCampo(codcampo As Long, Tipo As String, Optional OldCadena As String)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim cmp As GRPTC_Campo
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select a.*, b.codprodu from rcampos as a , variedades as b where b.codvarie = a.codvarie"
    Sql = Sql & " and a.codcampo = " & CStr(codcampo)
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        '-- creamos el objeto auxiliar que montará el XML de trazatec
        Set cmp = New GRPTC_Campo
        '-- vamos cargando los diferentes valores
        cmp.Codsocio = Rs!Codsocio
        cmp.codcampo = Rs!codcampo
        cmp.codprodu = Rs!codprodu
        cmp.codvarie = Rs!codvarie
        cmp.codparti = Rs!codparti
        cmp.hanegada = 0 ' no interesa en trazatec
        cmp.numarbol = 0 ' tampoco interesa
        cmp.Poligono = Rs!Poligono
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montará en el momento de la grabación
        chv.BD_Org = "AGRO"
        '[Monica]16/11/2011: solo en Alzira es la tabla SCAMP1
        If vParamAplic.Cooperativa = 4 Then
            chv.Tabla = "SCAMP1"
        Else
            chv.Tabla = "SCAMPO"
        End If
        chv.Oper = Tipo
        chv.Fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
'[Monica] 31/12/2009 solo el campo
'        chv.Clv_New = CStr(cmp.codsocio) & _
'                            "&" & CStr(cmp.codcampo) & _
'                            "&" & CStr(cmp.codprodu) & _
'                            "&" & CStr(cmp.codvarie)
        chv.Clv_New = CStr(cmp.codcampo)
        
        
        chv.XML = cmp.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            '[Monica]17/09/2013: solo para picassent cuando se está modificando el campo
            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 7 Or vParamAplic.Cooperativa = 16 Then
                chv.Clv_Old = OldCadena
            Else
                chv.Clv_Old = chv.Clv_New
            End If
        End If
        chv.Grabar
    End If
End Sub


Sub CargarUnSocio(Codsocio As Long, Tipo As String)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim soc As GRPTC_Socio
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select * from rsocios "
    Sql = Sql & " where codsocio = " & CStr(Codsocio)
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        '-- creamos el objeto auxiliar que montará el XML de trazatec
        Set soc = New GRPTC_Socio
        '-- vamos cargando los diferentes valores
        soc.Codsocio = Rs!Codsocio
        soc.nifsocio = Rs!nifsocio
        soc.nomsocio = Rs!nomsocio
        soc.domsocio = Rs!dirsocio
        soc.telsocio = DBLet(Rs!telsoci1)
        soc.CodPobla = 0 ' algo hay que hacer
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montará en el momento de la grabación
        chv.BD_Org = "AGRO"
        chv.Tabla = "SSOCIO"
        chv.Oper = Tipo
        chv.Fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
        chv.Clv_New = CStr(soc.Codsocio)
        chv.XML = soc.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            chv.Clv_Old = chv.Clv_New
        End If
        chv.Grabar
    End If
End Sub


Sub CargarUnaPoblacion(CodPobla As String, Tipo As String)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim pob As GRPTC_Poblacion
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select * from rpueblos "
    Sql = Sql & " where codpobla = '" & CodPobla & "'"
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        '-- creamos el objeto auxiliar que montará el XML de trazatec
        Set pob = New GRPTC_Poblacion
        '-- vamos cargando los diferentes valores
        pob.CodPobla = Rs!CodPobla
        pob.desPobla = Rs!desPobla
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montará en el momento de la grabación
        chv.BD_Org = "AGRO"
        chv.Tabla = "SPOBLA"
        chv.Oper = Tipo
        chv.Fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
        chv.Clv_New = pob.CodPobla
        chv.XML = pob.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            chv.Clv_Old = chv.Clv_New
        End If
        chv.Grabar
    End If
End Sub

Sub CargarUnaCuadrilla(codcapat As Long, Tipo As String)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim cua As GRPTC_Cuadrilla
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select * from rcapataz "
    Sql = Sql & " where codcapat = " & CStr(codcapat)
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        '-- creamos el objeto auxiliar que montará el XML de trazatec
        Set cua = New GRPTC_Cuadrilla
        '-- vamos cargando los diferentes valores
        cua.codcapat = Rs!codcapat
        cua.nomcapat = Rs!nomcapat
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montará en el momento de la grabación
        chv.BD_Org = "AGRO"
        chv.Tabla = "SCAPAT"
        chv.Oper = Tipo
        chv.Fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
        chv.Clv_New = CStr(cua.codcapat)
        chv.XML = cua.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            chv.Clv_Old = chv.Clv_New
        End If
        chv.Grabar
    End If
End Sub



Sub CargarUnaPartida(codparti As Long, Tipo As String)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim par As GRPTC_Partida
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select * from rpartida "
    Sql = Sql & " where codparti = " & CStr(codparti)
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        '-- creamos el objeto auxiliar que montará el XML de trazatec
        Set par = New GRPTC_Partida
        '-- vamos cargando los diferentes valores
        par.codparti = Rs!codparti
        par.nomparti = Rs!nomparti
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montará en el momento de la grabación
        chv.BD_Org = "AGRO"
        chv.Tabla = "SPARTI"
        chv.Oper = Tipo
        chv.Fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
        chv.Clv_New = CStr(par.codparti)
        chv.XML = par.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            chv.Clv_Old = chv.Clv_New
        End If
        chv.Grabar
    End If
End Sub


Sub CargarUnVehiculo(codTrans As String, Tipo As String)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim tra As GRPTC_Vehiculo
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select * from rtransporte "
    Sql = Sql & " where codtrans = '" & codTrans & "'"
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        '-- creamos el objeto auxiliar que montará el XML de trazatec
        Set tra = New GRPTC_Vehiculo
        '-- vamos cargando los diferentes valores
        tra.nomcamio = Rs!nomtrans
        tra.matricul = Rs!codTrans
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montará en el momento de la grabación
        chv.BD_Org = "AGRO"
        chv.Tabla = "SCAMIO"
        chv.Oper = Tipo
        chv.Fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
        chv.Clv_New = CStr(tra.matricul)
        chv.XML = tra.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            chv.Clv_Old = chv.Clv_New
        End If
        chv.Grabar
    End If
End Sub



Sub CargarUnProducto(codprodu As Long, Tipo As String)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim pro As GRPTC_Producto
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select * from productos "
    Sql = Sql & " where codprodu = " & CStr(codprodu)
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        '-- creamos el objeto auxiliar que montará el XML de trazatec
        Set pro = New GRPTC_Producto
        '-- vamos cargando los diferentes valores
        pro.codprodu = Rs!codprodu
        pro.nomprodu = Rs!nomprodu
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montará en el momento de la grabación
        chv.BD_Org = "AGRO"
        chv.Tabla = "SPRODU"
        chv.Oper = Tipo
        chv.Fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
        chv.Clv_New = CStr(pro.codprodu)
        chv.XML = pro.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            chv.Clv_Old = chv.Clv_New
        End If
        chv.Grabar
    End If
End Sub

Sub CargarUnaVariedad(codvarie As Long, Tipo As String, Optional OldCadena As String)
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim vari As GRPTC_Variedad
    Dim chv As GRPTC_Chivato
    '-- leemos mediante una consulta única todos los campos
    Sql = "select * from variedades "
    Sql = Sql & " where codvarie = " & CStr(codvarie)
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        '-- creamos el objeto auxiliar que montará el XML de trazatec
        Set vari = New GRPTC_Variedad
        '-- vamos cargando los diferentes valores
        vari.codvarie = Rs!codvarie
        vari.nomvarie = Rs!nomvarie
        vari.codprodu = Rs!codprodu
        '-- Y ahora el objeto chivato para grabar
        Set chv = New GRPTC_Chivato
        chv.Id = 0 'ya lo montará en el momento de la grabación
        chv.BD_Org = "AGRO"
        chv.Tabla = "SVARIE"
        chv.Oper = Tipo
        chv.Fecha = Format(Date, "dd/mm/yyyy")
        chv.Sep = "&"
        chv.Clv_Old = ""
        chv.Clv_New = CStr(vari.codvarie)
        chv.XML = vari.GenXML
        If Tipo = "D" Then
            chv.Clv_Old = chv.Clv_New
            chv.Clv_New = ""
            chv.XML = ""
        End If
        If Tipo = "U" Then
            '[Monica]18/09/2013: Si es Picassent o Quatretonda tengo que meter producto variedad
            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 7 Or vParamAplic.Cooperativa = 16 Then
                chv.Clv_Old = OldCadena
            Else
                chv.Clv_Old = chv.Clv_New
            End If
        End If
        chv.Grabar
    End If
End Sub



