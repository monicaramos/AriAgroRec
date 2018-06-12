Attribute VB_Name = "TraspCooperativas"
Option Explicit

'#########################################################################################################
'
'################### MODULO CON LAS FUNCIONES NECESARIAS PARA COMUNICACION ENTRE COOPIC Y PICASSENT
'
'#########################################################################################################


Public Function ComunicaCooperativa(vtabla As String, vSQL As String, vOperacion As String, vObservaciones As String) As Boolean
' vOperacion: I insercion
'             U modificacion
Dim Sql As String
Dim vInsert As String
Dim vValues As String

    On Error GoTo eComunicaCooperativa
    
    ComunicaCooperativa = False
        
    Sql = "INSERT INTO comunica (fechacreacion,usuariocreacion,tipo,tabla,sqlaejecutar,  "
    Sql = Sql & "observaciones,fechadescarga,usuariodescarga) VALUES ("
    Sql = Sql & DBSet(Now, "FH") & "," & vUsu.Codigo & "," & DBSet(vOperacion, "T") & "," & DBSet(vtabla, "T") & ","
    Sql = Sql & DBSet(vSQL, "T") & "," & DBSet(vObservaciones, "T", "S") & "," & ValorNulo & "," & ValorNulo & ")"
    
    conn.Execute Sql
    
    ComunicaCooperativa = True
    Exit Function
    
eComunicaCooperativa:
    MuestraError Err.Number, "Comunica cooperativa", Err.Description
End Function


Public Function EsSocioCooperativa(vSoc As String) As Boolean
Dim Sql As String

    EsSocioCooperativa = True
    If vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16 Then Exit Function
    
    EsSocioCooperativa = (CLng(ComprobarCero(vSoc)) < cMaxSocio)

End Function

Public Function EsVariedadComercializada(vCodvarie As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from variedades  where rcampos.codvarie = " & DBSet(vCodvarie, "N")
    Sql = Sql & " and variedades.comerciocomun = 1"
    
    EsVariedadComercializada = (TotalRegistros(Sql) <> 0)

End Function



Public Function EsDeVariedadComercializada(vCodcampo As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from variedades inner join rcampos on variedades.codvarie = rcampos.codvarie where rcampos.codcampo = " & DBSet(vCodcampo, "N")
    Sql = Sql & " and variedades.comerciocomun = 1"
    
    EsDeVariedadComercializada = (TotalRegistros(Sql) <> 0)

End Function

Public Function TieneCamposVariedadComercializada(vSocio As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from rcampos inner join variedades on rcampos.codvarie = variedades.codvarie where codsocio = " & DBSet(vSocio, "N")
    Sql = Sql & " and variedades.comerciocomun = 1"
    
    TieneCamposVariedadComercializada = (TotalRegistros(Sql) <> 0)
    
End Function

