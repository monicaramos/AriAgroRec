Attribute VB_Name = "ModBasico"
Option Explicit


Public Sub AyudaAlmacenCom(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Nombre|4595|;"
    frmCom.CadenaConsulta = "SELECT salmpr.codalmac, salmpr.nomalmac "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM salmpr "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999|salmpr|codalmac|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||salmpr|nomalmac|||"
    frmCom.Tag3 = ""
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Maxlen3 = 0
    
    frmCom.pConn = cAgro
    
    frmCom.tabla = "salmpr"
    frmCom.CampoCP = "codalmac"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Almacenes Propios de Comercial"
    frmCom.DeConsulta = True

    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, -1500
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaBancosCom(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Nombre|4595|;"
    frmCom.CadenaConsulta = "SELECT banpropi.codbanpr, banpropi.nombanpr "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM banpropi "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|99|banpropi|codbanpr|00|S|"
    frmCom.Tag2 = "Nombre|T|N|||banpropi|nombanpr|||"
    frmCom.Tag3 = ""
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Maxlen3 = 0

    frmCom.pConn = cAgro

    frmCom.tabla = "banpropi"
    frmCom.CampoCP = "codbanpr"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Bancos Propios de Comercial"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, -1500
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 4900
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaClasesCom(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Nombre|4595|;"
    frmCom.CadenaConsulta = "SELECT clases.codclase, clases.nomclase "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM clases "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999|clases|codclase|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||clases|nomclase|||"
    frmCom.Tag3 = ""
    
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Maxlen3 = 0
    
    frmCom.pConn = cAgro

    frmCom.tabla = "clases"
    frmCom.CampoCP = "codclase"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Clases de Comercial"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, -1500
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaGrupoCom(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|95|;S|txtAux(1)|T|Nombre|4595|;"
    frmCom.CadenaConsulta = "SELECT grupopro.codgrupo, grupopro.nomgrupo "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM grupopro "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999|grupopro|codgrupo|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||grupopro|nomgrupo|||"
    frmCom.Tag3 = ""
    
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Maxlen3 = 0
    
    frmCom.pConn = cAgro
    
    frmCom.tabla = "grupopro"
    frmCom.CampoCP = "codgrupo"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Grupos de Producto de Comercial"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, -1500
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaHorarioCom(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Nombre|4595|;"
    frmCom.CadenaConsulta = "SELECT cchorario.codhorario, cchorario.descripc "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM cchorario "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999|cchorario|codhorario|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||cchorario|descripc|||"
    frmCom.Tag3 = ""
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Maxlen3 = 0
    
    frmCom.pConn = cAgro
    
    frmCom.tabla = "cchorario"
    frmCom.CampoCP = "codhorario"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Horarios Costes de Comercial"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, -1500
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaClienteAriges(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Descripción|4595|;"
    frmBas.CadenaConsulta = "SELECT sclien.codclien, sclien.nomclien "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM " & vParamAplic.BDAriges & ".sclien "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código|N|N|||sclien|codclien|000000|S|"
    frmBas.Tag2 = "Nombre|T|N|||sclien|nomclien|||"
    frmBas.Tag3 = ""
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 40
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = vParamAplic.BDAriges & ".sclien"
    frmBas.CampoCP = "codclien"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Clientes de Suministros"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, -1500
    
    
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaClienteCom(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Descripción|4595|;"
    frmBas.CadenaConsulta = "SELECT clientes.codclien, clientes.nomclien "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM clientes "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código|N|N|||clientes|codclien|000000|S|"
    frmBas.Tag2 = "Nombre|T|N|||clientes|nomclien|||"
    frmBas.Tag3 = ""
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 40
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "clientes"
    frmBas.CampoCP = "codclien"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Clientes de Comercial"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, -1500
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmBas.DataGrid1.Height = 7420
    frmBas.DataGrid1.Top = 870
    frmBas.FrameBotonGnral.visible = True
    frmBas.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    frmBas.Show vbModal
    
    
End Sub


Public Sub AyudaVariedad(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1405|;S|txtAux(1)|T|Descripción|3000|;S|txtAux(2)|T|Producto|2595|;"
    frmBas.CadenaConsulta = "SELECT variedades.codvarie, variedades.nomvarie, productos.nomprodu "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM variedades inner join productos on variedades.codprodu = productos.codprodu "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código |N|N|||variedades|codvarie|000000|S|"
    frmBas.Tag2 = "Descripción|T|N|||variedades|nomvarie|||"
    frmBas.Tag3 = "Producto|T|N|||variedades|nomprodu|||"
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 30
    frmBas.Maxlen3 = 100
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "variedades"
    frmBas.CampoCP = "codvarie"
    frmBas.TipoCP = "N"
    frmBas.Caption = "Variedades"
    frmBas.DeConsulta = True
    
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmBas.DataGrid1.Height = 7420
    frmBas.DataGrid1.Top = 870
    frmBas.FrameBotonGnral.visible = True
    frmBas.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    frmBas.Show vbModal

    
End Sub


Public Sub AyudaVariedadPrevio(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1205|;S|txtAux(1)|T|Descripción|2700|;S|txtAux(2)|T|Producto|2595|;S|txtAux(3)|T|Clase|1000|;S|txtAux(3)|T|Nombre|2500|;"
    frmBas.CadenaConsulta = "SELECT variedades.codvarie, variedades.nomvarie, productos.nomprodu, variedades.codclase, clases.nomclase  "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM (variedades inner join productos on variedades.codprodu = productos.codprodu) inner join clases on variedades.codclase = clases.codclase "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código |N|N|||variedades|codvarie|000000|S|"
    frmBas.Tag2 = "Descripción|T|N|||variedades|nomvarie|||"
    frmBas.Tag3 = "Producto|T|N|||variedades|nomprodu|||"
    frmBas.Tag4 = "Clase|N|N|||variedades|codclase|000000||"
    frmBas.Tag3 = "Nombre Clase|T|N|||clases|nomclase|||"
    
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 30
    frmBas.Maxlen3 = 40
    frmBas.Maxlen4 = 6
    frmBas.Maxlen5 = 30
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "variedades"
    frmBas.CampoCP = "codvarie"
    frmBas.TipoCP = "N"
    frmBas.Caption = "Variedades"
    frmBas.DeConsulta = True
    
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 3000
    
    frmBas.Show vbModal

    
End Sub







Public Sub AyudaCuadrillas(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1405|;S|txtAux(1)|T|Capataz|1200|;S|txtAux(2)|T|Nombre|4395|;"
    frmBas.CadenaConsulta = "SELECT rcuadrilla.codcuadrilla, rcuadrilla.codcapat, rcapataz.nomcapat "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM rcuadrilla inner join rcapataz on rcuadrilla.codcapat = rcapataz.codcapat "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código |N|N|||rcuadrilla|codcuadrilla|000000|S|"
    frmBas.Tag2 = "Capataz|N|N|||rcuadrilla|codcapat|000000||"
    frmBas.Tag3 = "Nombre|T|N|||rcapataz|nomcapat|||"
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 60
    frmBas.Maxlen3 = 124
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "rcuadrilla"
    frmBas.CampoCP = "codcuadrilla"
    frmBas.TipoCP = "N"
    frmBas.Caption = "Cuadrillas"
    frmBas.DeConsulta = True
    
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal

    
End Sub

Public Sub AyudaConceptos(frmBas As frmBasico2, Optional CodActual As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Descripción|4595|;"
    frmBas.CadenaConsulta = "SELECT rriego.codriego, rriego.nomriego "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM rriego "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "Código|N|N|||rriego|codriego|00|S|"
    frmBas.Tag2 = "Nombre|T|N|||rriego|nomriego|||"
    frmBas.Tag3 = ""
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 40
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "rriego"
    frmBas.CampoCP = "codriego"
    frmBas.TipoCP = "N"
    frmBas.Caption = "Conceptos"
    frmBas.DeConsulta = True
    
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, -1500
    
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaFamiliasADV(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Descripción|4595|;"
    frmCom.CadenaConsulta = "SELECT advfamia.codfamia, advfamia.nomfamia "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM advfamia "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|9999|advfamia|codfamia|0000|S|"
    frmCom.Tag2 = "Descripción|T|N|||advfamia|nomfamia|||"
    frmCom.Tag3 = ""
    
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 25
    frmCom.Maxlen3 = 0
    
    frmCom.pConn = cAgro
    
    frmCom.tabla = "advfamia"
    frmCom.CampoCP = "codfamia"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Familias de Artículos ADV"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, -1500
   
    frmCom.Show vbModal
End Sub


Public Sub AyudaTiposDocumentos(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Descripción|3500|;S|txtAux(2)|T|Fichero|2595|;"
    frmBas.CadenaConsulta = "SELECT scryst.codcryst, scryst.nomcryst, scryst.documrpt "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM scryst "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código Documento|N|N|||scryst|codcryst|0000|S|"
    frmBas.Tag2 = "Descripción|T|N|||scryst|nomcryst|||"
    frmBas.Tag3 = "Fichero rpt|T|N|||scryst|documrpt|||"
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 30
    frmBas.Maxlen3 = 100
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "scryst"
    frmBas.CampoCP = "codcryst"
    frmBas.TipoCP = "N"
    frmBas.Caption = "Tipos de Documentos"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 0
    
    frmBas.Show vbModal
End Sub




Public Sub AyudaTUnidadesCom(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Descripción|4595|;"
    frmCom.CadenaConsulta = "SELECT sunida.codunida, sunida.nomunida "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sunida "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|99|sunida|codunida|00|S|"
    frmCom.Tag2 = "Descripción|T|N|||sunida|nomunida|||"
    frmCom.Tag3 = ""
    frmCom.Maxlen1 = 2
    frmCom.Maxlen2 = 10
    frmCom.Maxlen3 = 0
    
    frmCom.pConn = cAgro
    
    frmCom.tabla = "sunida"
    frmCom.CampoCP = "codunida"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Tipos de Unidad de Comercial"
    frmCom.DeConsulta = True
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, -1500
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaProveedoresCom(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Nombre|4595|;"
    frmCom.CadenaConsulta = "SELECT proveedor.codprove, proveedor.nomprove "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM proveedor "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999999|proveedor|codprove|000000|S|"
    frmCom.Tag2 = "Nombre|T|N|||proveedor|nomprove|||"
    frmCom.Tag3 = ""
    frmCom.Maxlen1 = 6
    frmCom.Maxlen2 = 40
    frmCom.Maxlen3 = 0

    frmCom.pConn = cAgro

    frmCom.tabla = "proveedor"
    frmCom.CampoCP = "codprove"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Proveedores de Comercial"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, -1500
    
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaProductosCom(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Nombre|4595|;"
    frmCom.CadenaConsulta = "SELECT productos.codprodu, productos.nomprodu "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM productos "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999|productos|codprodu|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||productos|nomprodu|||"
    frmCom.Tag3 = ""
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Maxlen3 = 0
    
    frmCom.pConn = cAgro
    
    
    frmCom.tabla = "productos"
    frmCom.CampoCP = "codprodu"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Productos de Comercial"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, -1500
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaForfaitsCom(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|1500|;S|txtAux(1)|T|Nombre|3900|;"
    frmCom.CadenaConsulta = "SELECT forfaits.codforfait, forfaits.nomconfe "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM forfaits "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|T|N|||forfaits|codforfait||S|"
    frmCom.Tag2 = "Nombre|T|N|||forfaits|nomconfe|||"
    frmCom.Tag3 = ""
    
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Maxlen3 = 0
    
    frmCom.pConn = cAgro

    frmCom.tabla = "forfaits"
    frmCom.CampoCP = "codforfait"
    frmCom.TipoCP = "T"
    frmCom.Caption = "Forfaits de Comercial"
    frmCom.DeConsulta = True

    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, -1600
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaFPagoCom(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Nombre|4595|;"
    frmCom.CadenaConsulta = "SELECT forpago.codforpa, forpago.nomforpa "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM forpago "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999|forpago|codforpa|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||forpago|nomforpa|||"
    frmCom.Tag3 = ""
    
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Maxlen3 = 0
    
    frmCom.pConn = cAgro
    
    frmCom.tabla = "forpago"
    frmCom.CampoCP = "codforpa"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Formas Pago de Comercial"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, -1500
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaFamiliasCom(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Descripción|4595|;"
    frmCom.CadenaConsulta = "SELECT sfamia.codfamia, sfamia.nomfamia "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sfamia "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999|sfamia|codfamia|000|S|"
    frmCom.Tag2 = "Descripción|T|N|||sfamia|nomfamia|||"
    frmCom.Tag3 = ""
    
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 25
    frmCom.Maxlen3 = 0
    
    frmCom.pConn = cAgro
    
    frmCom.tabla = "sfamia"
    frmCom.CampoCP = "codfamia"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Familias de Comercial"
    frmCom.DeConsulta = True
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, -1500
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaFrasTerceros(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Nro.Factura|1505|;S|txtAux(1)|T|Fecha|1495|;S|txtAux(2)|T|Socio|1000|;S|txtAux(3)|T|Nombre|5000|;"
    
    frmBas.CadenaConsulta = "SELECT rcafter.numfactu, rcafter.fecfactu, rcafter.codsocio, rcafter.nomsocio "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM rcafter"
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Nº Factura|T|N|||rcafter|numfactu||S|"
    frmBas.Tag2 = "Fecha Factura|F|N|||rcafter|fecfactu|dd/mm/yyyy|S|"
    frmBas.Tag3 = "Cod.Tercero|N|N|0|999999|rcafter|codsocio|000000|S|"
    frmBas.Tag4 = "Nombre Tercero|T|N|||rcafter|nomsocio||N|"
    frmBas.Tag5 = ""
    frmBas.Tag6 = ""
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 10
    frmBas.Maxlen3 = 6
    frmBas.Maxlen4 = 40
    frmBas.Maxlen5 = 0
    frmBas.Maxlen6 = 0
    frmBas.Maxlen7 = 0
    
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "rcafter"
    frmBas.TipoCP = "T"
    frmBas.CampoCP = "numfactu"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Facturas Terceros"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|2|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 2000
    
    frmBas.Show vbModal
    
    
End Sub



Public Sub AyudaPartesCampo(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Parte|1100|;S|txtAux(1)|T|Fecha|1400|;S|txtAux(2)|T|Cuadrilla|1100|;S|txtAux(3)|T|Capataz|1100|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(4)|T|Nombre|4300|;"
    
    frmBas.CadenaConsulta = "SELECT rpartes.nroparte, rpartes.fechapar, rpartes.codcuadrilla, rcuadrilla.codcapat, "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " rcapataz.nomcapat "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM (rpartes INNER JOIN rcuadrilla ON rpartes.codcuadrilla=rcuadrilla.codcuadrilla) " & _
                            " INNER JOIN rcapataz ON rcuadrilla.codcapat=rcapataz.codcapat"
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Nº Parte|N|S|||rpartes|nroparte|0000000|S|"
    frmBas.Tag2 = "Fecha Parte|F|N|||rpartes|fechapar|dd/mm/yyyy||"
    frmBas.Tag3 = "Cod.Cuadrilla|N|N|0|999999|rpartes|codcuadrilla|000000||"
    frmBas.Tag4 = "Cod.Capataz|N|N|0|999999|rcuadrilla|codcapat|000000||"
    frmBas.Tag5 = "Nom.Capataz|T|S|||rcapataz|nomcapat|||"
    frmBas.Tag6 = ""
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 10
    frmBas.Maxlen3 = 6
    frmBas.Maxlen4 = 6
    frmBas.Maxlen5 = 40
    frmBas.Maxlen6 = 0
    frmBas.Maxlen7 = 0
    
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "(rpartes INNER JOIN rcuadrilla ON rpartes.codcuadrilla=rcuadrilla.codcuadrilla) " & _
                            " INNER JOIN rcapataz ON rcuadrilla.codcapat=rcapataz.codcapat"
    frmBas.CampoCP = "nroparte"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Partes de Campo"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 2000
    
    frmBas.Show vbModal
    
    
End Sub


Public Sub AyudaFVarClientes(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Seccion|2795|;S|txtAux(2)|T|Tipo|800|;S|txtAux(3)|T|Factura|1000|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(4)|T|Fecha|1500|;S|txtAux(5)|T|Socio|1000|;S|txtAux(6)|T|Cliente|1000|;"
    
    frmBas.CadenaConsulta = "SELECT fvarcabfact.codsecci, rseccion.nomsecci, fvarcabfact.codtipom, fvarcabfact.numfactu, fvarcabfact.fecfactu, "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " fvarcabfact.codsocio, fvarcabfact.codclien "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM fvarcabfact inner join rseccion on fvarcabfact.codsecci = rseccion.codsecci "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Seccion|N|N|0|999|fvarcabfact|codsecci|000||"
    frmBas.Tag2 = "Nombre|T|N|||clientes|nomclien|||"
    frmBas.Tag3 = "Tipo Movimiento|T|N|||fvarcabfact|codtipom||S|"
    frmBas.Tag4 = "Nº de Factura|N|S|0|9999999|fvarcabfact|numfactu|0000000|S|"
    frmBas.Tag5 = "Fecha Factura|F|N|||fvarcabfact|fecfactu|dd/mm/yyyy|S|"
    frmBas.Tag6 = "Cod.Socio|N|S|||fvarcabfact|codsocio|000000||"
    frmBas.Tag7 = "Cod.Cliente|N|S|||fvarcabfact|codclien|000000||"
    frmBas.Maxlen1 = 3
    frmBas.Maxlen2 = 40
    frmBas.Maxlen3 = 3
    frmBas.Maxlen4 = 7
    frmBas.Maxlen5 = 10
    frmBas.Maxlen6 = 6
    frmBas.Maxlen7 = 6
    
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "fvarcabfact inner join rseccion"
    frmBas.CampoCP = "codclien"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Facturas Varias"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|2|3|4|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 2000
    
    frmBas.Show vbModal
    
    
End Sub



Public Sub AyudaFVarProveedores(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Seccion|2795|;S|txtAux(2)|T|Tipo|800|;S|txtAux(3)|T|Factura|1000|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(4)|T|Fecha|1500|;S|txtAux(5)|T|Socio|1000|;S|txtAux(6)|T|Nombre|2500|;"
    
    frmBas.CadenaConsulta = "SELECT fvarcabfactpro.codsecci, rseccion.nomsecci, fvarcabfactpro.codtipom, fvarcabfactpro.numfactu, fvarcabfactpro.fecfactu, "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " fvarcabfactpro.codsocio, rsocios.nomsocio "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM (fvarcabfactpro inner join rseccion on fvarcabfactpro.codsecci = rseccion.codsecci) inner join rsocios on fvarcabfactpro.codsocio = rsocios.codsocio "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Seccion|N|N|0|999|fvarcabfactpro|codsecci|000||"
    frmBas.Tag2 = "Nombre|T|N|||clientes|nomclien|||"
    frmBas.Tag3 = "Tipo Movimiento|T|N|||fvarcabfactpro|codtipom||S|"
    frmBas.Tag4 = "Nº de Factura|N|S|0|9999999|fvarcabfactpro|numfactu|0000000|S|"
    frmBas.Tag5 = "Fecha Factura|F|N|||fvarcabfactpro|fecfactu|dd/mm/yyyy|S|"
    frmBas.Tag6 = "Socio|N|S|||fvarcabfactpro|codsocio|000000||"
    frmBas.Tag7 = "Nombre|T|S|||rsocios|nomsocio|||"
    frmBas.Maxlen1 = 3
    frmBas.Maxlen2 = 30
    frmBas.Maxlen3 = 3
    frmBas.Maxlen4 = 7
    frmBas.Maxlen5 = 10
    frmBas.Maxlen6 = 6
    frmBas.Maxlen7 = 40
    
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "(fvarcabfactpro inner join rseccion)"
    frmBas.CampoCP = "codsecci"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Facturas Varias Proveedor"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|2|3|4|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 3500
    
    frmBas.Show vbModal
    
    
End Sub



Public Sub AyudaGlobalGap(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Descripción|4595|;"
    frmBas.CadenaConsulta = "SELECT rglobalgap.codigo, rglobalgap.descripcion "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM rglobalgap "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "Código|T|N|||rglobalgap|codigo||S|"
    frmBas.Tag2 = "Descripción|T|N|||rglobalgap|descripcion|||"
    frmBas.Tag3 = ""
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 40
    frmBas.Maxlen3 = 0
    
    frmBas.tabla = "rglobalgap"
    frmBas.CampoCP = "codigo"
    frmBas.TipoCP = "T"
    frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "GlobalGap"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, -1500
    
    
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaIncidenciasOrdenesRecogida(frmBas As frmBasico2, Optional CodActual As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|900|;S|txtAux(1)|T|Descripción|4595|;"
    frmBas.CadenaConsulta = "SELECT rplagasaux.idplaga, rplagasaux.nomplaga "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM rplagasaux "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "Código|N|N|||rplagasaux|idplaga|000000|S|"
    frmBas.Tag2 = "Nombre|T|N|||rplagasaux|nomplaga|||"
    frmBas.Tag3 = ""
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 40
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = cAgro
    
    
    frmBas.tabla = "rplagasaux"
    frmBas.CampoCP = "idplaga"
    frmBas.TipoCP = "N"
    frmBas.Caption = "Incidencias Ordenes Recogida"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, -1500
    
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaControlDestrio(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Fecha|1305|;S|txtAux(1)|T|Socio|900|;S|txtAux(2)|T|Nombre|3430|;S|txtAux(3)|T|Variedad|1050|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(4)|T|Nombre|1600|;S|txtAux(5)|T|Campo|1200|;S|txtAux(6)|T|Número|1000|;"
    
    frmBas.CadenaConsulta = "SELECT rcontrol.fechacla, rcontrol.codsocio, rsocios.nomsocio, rcontrol.codvarie, variedades.nomvarie, "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " rcontrol.codcampo, rcontrol.nroclasif "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM (rcontrol inner join rsocios on rcontrol.codsocio = rsocios.codsocio) inner join variedades on rcontrol.codvarie = variedades.codvarie "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Fecha Entrada|F|N|||rcontrol|fechacla|dd/mm/yyyy|S|"
    frmBas.Tag2 = "Socio|N|N|||rcontrol|codsocio|000000|S|"
    frmBas.Tag3 = "NomSocio|T|N|||rsocios|nomsocio|||"
    frmBas.Tag4 = "Variedad|N|N|0|999999|rcontrol|codvarie|000000|S|"
    frmBas.Tag5 = "NomVariedad|T|N|||variedades|nomvarie|||"
    frmBas.Tag6 = "Campo|N|N|||rcontrol|codcampo|00000000|S|"
    frmBas.Tag7 = "Nro.Clasif|N|S|||rcontrol|nroclasif|0000000|S|"
    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 6
    frmBas.Maxlen3 = 40
    frmBas.Maxlen4 = 6
    frmBas.Maxlen5 = 25
    frmBas.Maxlen6 = 8
    frmBas.Maxlen7 = 7
    
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "(rcontrol inner join rsocios on rcontrol.codsocio = rsocios.codsocio) inner join variedades on rcontrol.codvarie = variedades.codvarie"
    frmBas.CampoCP = "rcontrol.codsocio"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Control de Destrio"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|3|5|6|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 3500
    
    frmBas.Show vbModal
    
    
End Sub



Public Sub AyudaVtaFruta(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Albarán|1505|;S|txtAux(1)|T|Código|1000|;S|txtAux(2)|T|Nombre Cliente/Socio|5000|;S|txtAux(3)|T|Fecha|1495|;"
    
    frmBas.CadenaConsulta = "SELECT vtafrutacab.numalbar, concat(if(vtafrutacab.codclien is null,'',vtafrutacab.codclien),if(vtafrutacab.codsocio is null,'',vtafrutacab.codsocio)) as codigo,"
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " concat(if(clientes.nomclien is null,'',clientes.nomclien), if(rsocios.nomsocio is null,'',rsocios.nomsocio)) as nombre, vtafrutacab.fecalbar "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM (vtafrutacab LEFT JOIN clientes ON vtafrutacab.codclien=clientes.codclien) left join rsocios On vtafrutacab.codsocio = rsocios.codsocio "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Albarán|N|N|||vtafrutacab|numalbar|0000000|S|"
    frmBas.Tag2 = "Cliente/Socio|N|N|0|999999|vtafrutacab|codigo|000000||"
    frmBas.Tag3 = "Nombre Cliente/Socio|T|N|||vtafrutacab|nomsocio||N|"
    frmBas.Tag4 = "Fecha|F|N|||vtafrutacab|fecalbar|dd/mm/yyyy||"
    frmBas.Tag5 = ""
    frmBas.Tag6 = ""
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 6
    frmBas.Maxlen3 = 40
    frmBas.Maxlen4 = 10
    frmBas.Maxlen5 = 0
    frmBas.Maxlen6 = 0
    frmBas.Maxlen7 = 0
    
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "(vtafrutacab LEFT JOIN clientes ON vtafrutacab.codclien=clientes.codclien) left join rsocios On vtafrutacab.codsocio = rsocios.codsocio "
    frmBas.CampoCP = "numalbar"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Venta de Fruta Báscula"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|2|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 2000
    
    frmBas.Show vbModal
    
    
End Sub



Public Sub AyudaBonificaciones(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Nombre|4595|;"
    frmCom.CadenaConsulta = "SELECT rbonifica.codvarie, variedades.nomvarie "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM rbonifica inner join variedades  on rbonifica.codvarie = variedades.codvarie "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código Variedad|N|N|1|999999|rbonifica|codvarie|000000|S|"
    frmCom.Tag2 = "Nombre|T|N|||variedades|nomvarie|||"
    frmCom.Tag3 = ""
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Maxlen3 = 0
    
    frmCom.pConn = cAgro
    
    frmCom.tabla = "rbonifica"
    frmCom.CampoCP = "codvarie"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Bonificaciones"
    frmCom.DeConsulta = True

    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, -1500
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaFrasTransporte(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Tipo|905|;S|txtAux(1)|T|Factura|1195|;S|txtAux(2)|T|Fecha|1400|;S|txtAux(3)|T|Transportista|1800|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(4)|T|Nombre|3700|;"
    
    frmBas.CadenaConsulta = "SELECT rfacttra.codtipom, rfacttra.numfactu, rfacttra.fecfactu, rfacttra.codtrans, rtransporte.nomtrans "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM rfacttra inner join rtransporte on rfacttra.codtrans = rtransporte.codtrans "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Tipo Movimiento|T|N|||rfacttra|codtipom||S|"
    frmBas.Tag2 = "Nº Factura|N|S|||rfacttra|numfactu|0000000|S|"
    frmBas.Tag3 = "Fecha Factura|F|N|||rfacttra|fecfactu|dd/mm/yyyy|S|"
    frmBas.Tag4 = "Cod.Transportista|T|N|||rfacttra|codtrans||S|"
    frmBas.Tag5 = "Descripción|T|N|||rtransporte|nomtrans|||"
    frmBas.Tag6 = ""
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 3
    frmBas.Maxlen2 = 7
    frmBas.Maxlen3 = 10
    frmBas.Maxlen4 = 10
    frmBas.Maxlen5 = 40
    frmBas.Maxlen6 = 0
    frmBas.Maxlen7 = 0
    
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "rfacttra inner join rtransporte"
    frmBas.CampoCP = "numfactu"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Facturas de Transporte"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|2|3|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 2000
    
    frmBas.Show vbModal
    
    
End Sub


Public Sub AyudaPueblos(frmBas As frmBasico2, Optional CodActual As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Descripción|4595|;"
    frmBas.CadenaConsulta = "SELECT rpueblos.codpobla, rpueblos.despobla "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM rpueblos "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "Código|T|N|||rpueblos|codpobla||S|"
    frmBas.Tag2 = "Descripción|T|N|||rpueblos|despobla|||"
    frmBas.Tag3 = ""
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 30
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "rpueblos"
    frmBas.CampoCP = "codpobla"
    frmBas.TipoCP = "T"
    frmBas.Caption = "Pueblos"
    frmBas.DeConsulta = True
    
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, -1500
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmBas.DataGrid1.Height = 7420
    frmBas.DataGrid1.Top = 870
    frmBas.FrameBotonGnral.visible = True
    frmBas.FrameBotonGnral.Enabled = True
    ' hasta aqui

    frmBas.Show vbModal

End Sub

Public Sub AyudaFrasSocios(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Tipo Factura|3295|;S|txtAux(1)|T|Tipo|705|;S|txtAux(2)|T|Factura|1000|;S|txtAux(3)|T|Fecha|1500|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(4)|T|Socio|1000|;S|txtAux(5)|T|Nombre Socio|5490|;"
    
    frmBas.CadenaConsulta = "SELECT stipom.nomtipom, rfactsoc.codtipom, rfactsoc.numfactu, rfactsoc.fecfactu, rfactsoc.codsocio, rsocios.nomsocio "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM (rfactsoc inner join rsocios on rfactsoc.codsocio = rsocios.codsocio) left join usuarios.stipom on rfactsoc.codtipom = stipom.codtipom "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Tipo Factura|T|N|||usuarios.stipom|nomtipom||N|"
    frmBas.Tag2 = "Tipo|T|N|||rfactsoc|codtipom||S|"
    frmBas.Tag3 = "Nº Factura|N|S|||rfactsoc|numfactu|0000000|S|"
    frmBas.Tag4 = "Fecha Factura|F|N|||rfactsoc|fecfactu|dd/mm/yyyy|S|"
    frmBas.Tag5 = "Socio|N|N|0|999999|rfactsoc|codsocio|000000|N|"
    frmBas.Tag6 = "Nombre Socio|T|N|||rsocios|nomsocio||N|"
    frmBas.Maxlen1 = 40
    frmBas.Maxlen2 = 3
    frmBas.Maxlen3 = 7
    frmBas.Maxlen4 = 10
    frmBas.Maxlen5 = 6
    frmBas.Maxlen6 = 40
    
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "(rfactsoc inner join rsocios on rfactsoc.codsocio = rsocios.codsocio) left join usuarios.stipom on rfactsoc.codtipom = stipom.codtipom "
    frmBas.CampoCP = "numfactu"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Facturas Socios"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "1|2|3|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 6000
    
    frmBas.Show vbModal
    
    
End Sub


Public Sub AyudaHcoFrutas(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Albarán|950|;S|txtAux(1)|T|Fecha|1300|;S|txtAux(2)|T|Codigo|900|;S|txtAux(3)|T|Nombre|2600|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(4)|T|Socio|1000|;S|txtAux(5)|T|Nombre|5100|;S|txtAux(6)|T|Campo|1150|;"
    
    frmBas.CadenaConsulta = "select rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codsocio, rsocios.nomsocio, rhisfruta.codcampo "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " from (rhisfruta inner join variedades on rhisfruta.codvarie = variedades.codvarie) inner join rsocios on rhisfruta.codsocio = rsocios.codsocio "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código|N|N|||rhisfruta|numalbar|0000000|S|"
    frmBas.Tag2 = "Fecha|F|N|||rhisfruta|fecalbar|dd/mm/yyyy||"
    frmBas.Tag3 = "Variedad|N|N|||rhisfruta|codvarie|000000||"
    frmBas.Tag4 = "Nombre|T|N|||variedades|nomvarie|||"
    frmBas.Tag5 = "Socio|N|N|||rhisfruta|codsocio|000000||"
    frmBas.Tag6 = "Nombre Socio|T|N|||rsocios|nomsocio|||"
    frmBas.Tag7 = "Campo|N|N|||rhisfruta|codcampo|00000000||"
    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 10
    frmBas.Maxlen3 = 6
    frmBas.Maxlen4 = 20
    frmBas.Maxlen5 = 6
    frmBas.Maxlen6 = 40
    frmBas.Maxlen7 = 8
    
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "(rhisfruta inner join variedades on rhisfruta.codvarie = variedades.codvarie) inner join rsocios on rhisfruta.codsocio = rsocios.codsocio "
    frmBas.CampoCP = "numalbar"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Histórico de Fruta Clasificada"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 6000
    
    frmBas.Show vbModal
    
    
End Sub


Public Sub AyudaEntradaBascula(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Nota|1100|;S|txtAux(1)|T|Fecha|1400|;S|txtAux(2)|T|Socio|1000|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(3)|T|Nombre|4670|;S|txtAux(4)|T|Codigo|1000|;S|txtAux(5)|T|Variedad|2830|;"
    
    frmBas.CadenaConsulta = "select numnotac, fechaent, rentradas.codsocio, rsocios.nomsocio, rentradas.codvarie, variedades.nomvarie from rsocios, rentradas, variedades "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE rentradas.codsocio = rsocios.codsocio and rentradas.codvarie = variedades.codvarie "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Numero de Nota|N|S|1|9999999|rentradas|numnotac|0000000|S|"
    frmBas.Tag2 = "Fecha Entrada|F|N|||rentradas|fechaent|dd/mm/yyyy||"
    frmBas.Tag3 = "Código Socio|N|N|1|999999|rentradas|codsocio|000000|N|"
    frmBas.Tag4 = "Nombre Socio|T|N|||rsocios|nomsocio|||"
    frmBas.Tag5 = "Variedad|N|N|1|999999|rentradas|codvarie|000000||"
    frmBas.Tag6 = "Nombre Variedad|T|N|||variedades|nomvarie|||"
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 10
    frmBas.Maxlen3 = 6
    frmBas.Maxlen4 = 20
    frmBas.Maxlen5 = 6
    frmBas.Maxlen6 = 40
    frmBas.Maxlen7 = 0
    
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "rsocios, rentradas, variedades  "
    frmBas.CampoCP = "numnotac"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Entrada en báscula"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 5000
    
    frmBas.Show vbModal
    
    
End Sub


Public Sub AyudaPrecios(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1000|;S|txtAux(1)|T|Variedad|3070|;S|txtAux(2)|T|Tipo|1900|;S|txtAux(3)|T|Contador|1200|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(4)|T|Fecha Inicio|1400|;S|txtAux(5)|T|Fecha Fin|1400|;N|||||;"
    
    frmBas.CadenaConsulta = "select rprecios.codvarie, variedades.nomvarie, CASE rprecios.tipofact WHEN 0 THEN ""Anticipo"" WHEN 1 THEN ""Liquidacion"" WHEN 2 THEN ""Ind.Directa"" WHEN 3 THEN ""Complementaria"" WHEN 4 THEN ""Ant.Genérico"" WHEN 5 THEN ""Ant.Retirada"" END, rprecios.contador, rprecios.fechaini, rprecios.fechafin, rprecios.tipofact from rprecios inner join variedades on rprecios.codvarie = variedades.codvarie "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código|N|N|||rprecios|codvarie|000000|S|"
    frmBas.Tag2 = "Nombre|T|N|||variedades|nomvarie|||"
    frmBas.Tag3 = "Tipo|T|N|||rprecios|tipofact|||"
    frmBas.Tag4 = "Contador|N|N|||rprecios|contador|0000000||"
    frmBas.Tag5 = "Fecha Inicio|F|N|||rprecios|fechaini|dd/mm/yyyy||"
    frmBas.Tag6 = "Fecha Fin|F|N|||rprecios|fechafin|dd/mm/yyyy||"
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 20
    frmBas.Maxlen3 = 10
    frmBas.Maxlen4 = 7
    frmBas.Maxlen5 = 10
    frmBas.Maxlen6 = 10
    frmBas.Maxlen7 = 0
    
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "rprecios inner join variedades on rprecios.codvarie = variedades.codvarie "
    frmBas.CampoCP = "codvarie"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Precios"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|3|6|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 3000
    
    frmBas.Show vbModal
    
    
End Sub


Public Sub AyudaTrabajadores(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1000|;S|txtAux(1)|T|Nombre|4370|;S|txtAux(2)|T|Nif|1450|;S|txtAux(3)|T|Teléfono|1300|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(4)|T|Móvil|1300|;S|txtAux(5)|T|Fecha Alta|1280|;"
    
    frmBas.CadenaConsulta = "select codtraba, nomtraba, niftraba, teltraba, movtraba, fechaalta from straba "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código|N|N|||straba|codtraba|000000|S|"
    frmBas.Tag2 = "Nombre|T|N|||straba|nomtraba|||"
    frmBas.Tag3 = "Nif Traba|T|N|||straba|niftraba|||"
    frmBas.Tag4 = "Teléfono|T|N|||straba|teltraba|||"
    frmBas.Tag5 = "Móvil|T|N|||straba|movtraba|||"
    frmBas.Tag6 = "Fecha Alta|F|N|||straba|fechaalta|dd/mm/yyyy||"
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 20
    frmBas.Maxlen3 = 15
    frmBas.Maxlen4 = 15
    frmBas.Maxlen5 = 15
    frmBas.Maxlen6 = 10
    frmBas.Maxlen7 = 0
    
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "straba"
    frmBas.CampoCP = "codtraba"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Trabajadores"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 3700
    
    frmBas.Show vbModal
    
    
End Sub


Public Sub AyudaEntradaPesada(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Pesada|1100|;S|txtAux(1)|T|Fecha|1300|;S|txtAux(2)|T|Código|1500|;S|txtAux(3)|T|Transportista|4000|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(4)|T|K.Bruto|1400|;S|txtAux(5)|T|K.Neto|1400|;"
    
    frmBas.CadenaConsulta = "select nropesada, fecpesada, rpesadas.codtrans, rtransporte.nomtrans, rpesadas.kilosbrut, rpesadas.kilosnet from rpesadas inner join rtransporte on rpesadas.codtrans = rtransporte.codtrans "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Numero Pesada|N|S|1|9999999|rpesadas|nropesada|0000000|S|"
    frmBas.Tag2 = "Fecha Pesada|F|N|||rpesadas|fecpesada|dd/mm/yyyy||"
    frmBas.Tag3 = "Código Transp.|T|N|||rpesadas|codtrans||N|"
    frmBas.Tag4 = "Nombre Transportista|T|N|||rtransporte|nomtrans|||"
    frmBas.Tag5 = "Kilos Bruto|N|N|||rpesadas|kilosbrut|###,###,##0||"
    frmBas.Tag6 = "Kilos Neto|N|N|||rpesadas|kilosnet|###,###,##0||"
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 10
    frmBas.Maxlen3 = 15
    frmBas.Maxlen4 = 30
    frmBas.Maxlen5 = 10
    frmBas.Maxlen6 = 10
    frmBas.Maxlen7 = 0
    
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "rpesadas inner join rtransporte on rpesadas.codtrans = rtransporte.codtrans"
    frmBas.CampoCP = "nropesada"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Entrada de Pesada"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 3700
    
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaBodEntrada(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Albaran|1150|;S|txtAux(1)|T|Fecha|1400|;S|txtAux(2)|T|Socio|5050|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(3)|T|Variedad|2500|;S|txtAux(4)|T|Campo|1300|;"
    
    frmBas.CadenaConsulta = "select rhisfruta.numalbar, rhisfruta.fecalbar, rsocios.nomsocio, variedades.nomvarie, rhisfruta.codcampo from (rhisfruta inner join rsocios on rhisfruta.codsocio = rsocios.codsocio) inner join variedades on rhisfruta.codvarie = variedades.codvarie "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Nro.Albarán|N|S|||rhisfruta|numalbar|0000000|S|"
    frmBas.Tag2 = "Fecha albarán|F|N|||rhisfruta|fecalbar|dd/mm/yyyy||"
    frmBas.Tag3 = "Socio|T|N|||rsocios|nomsocio|||"
    frmBas.Tag4 = "Variedad|T|N|||variedades|nomvarie|||"
    frmBas.Tag5 = "Campo|N|N|||rhisfruta|codcampo|00000000||"
    frmBas.Tag6 = ""
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 10
    frmBas.Maxlen3 = 30
    frmBas.Maxlen4 = 6
    frmBas.Maxlen5 = 20
    frmBas.Maxlen6 = 0
    frmBas.Maxlen7 = 0
    
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "(rhisfruta inner join rsocios on rhisfruta.codsocio = rsocios.codsocio) inner join variedades on rhisfruta.codvarie = variedades.codvarie"
    frmBas.CampoCP = "numalbar"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Entradas de Bodega"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 4430
    
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaADVTratamientos(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|900|;S|txtAux(1)|T|Nombre|4170|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(2)|T|Fec.Inicio|1450|;S|txtAux(3)|T|Fec.Fin|1450|;"
    
    frmBas.CadenaConsulta = "select codtrata, nomtrata, fechaini, fechafin from advtrata "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código|T|N|||advtrata|codtrata||S|"
    frmBas.Tag2 = "Descripción|T|N|||advtrata|nomtrata|||"
    frmBas.Tag3 = "Fecha Inicio|F|S|||advtrata|fechaini|dd/mm/yyyy||"
    frmBas.Tag4 = "Fecha Fin|F|S|||advtrata|fechafin|dd/mm/yyyy||"
    frmBas.Tag5 = ""
    frmBas.Tag6 = ""
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 30
    frmBas.Maxlen3 = 10
    frmBas.Maxlen4 = 10
    frmBas.Maxlen5 = 0
    frmBas.Maxlen6 = 0
    frmBas.Maxlen7 = 0
    
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "advtrata"
    frmBas.CampoCP = "codtrata"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Tratamientos ADV"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 1000
    
    frmBas.Show vbModal
    
End Sub



Public Sub AyudaFacturasAlmazara(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "N||||0|;S|txtAux(0)|T|Tipo|1300|;S|txtAux(1)|T|Factura|1200|;S|txtAux(2)|T|Fecha|1400|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(3)|T|Codigo|1100|;S|txtAux(4)|T|Socio|4400|;"
    
    frmBas.CadenaConsulta = "select rcabfactalmz.tipofichero, case rcabfactalmz.tipofichero when 0 then ""Aceite"" when 1 then ""Aceituna"" when 2 then ""Stock"" end, "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & "rcabfactalmz.numfactu, rcabfactalmz.fecfactu, rcabfactalmz.codsocio, rsocios.nomsocio from rcabfactalmz inner join rsocios on rcabfactalmz.codsocio = rsocios.codsocio "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código|N|N|||rcabfactalmz|tipofichero|000000|S|"
    frmBas.Tag2 = "Nº Factura|N|N|||rcabfactalmz|numfactu|0000000|S|"
    frmBas.Tag3 = "Fecha Factura|F|N|||rcabfactalmz|fecfactu|dd/mm/yyyy|S|"
    frmBas.Tag4 = "Cod.Socio|N|N|0|999999|rcabfactalmz|codsocio|000000|N|"
    frmBas.Tag5 = "Nombre|T|N|||rsocios|nomsocio|||"
    frmBas.Tag6 = ""
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 7
    frmBas.Maxlen3 = 10
    frmBas.Maxlen4 = 6
    frmBas.Maxlen5 = 30
    frmBas.Maxlen6 = 0
    frmBas.Maxlen7 = 0
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "rcabfactalmz inner join rsocios on rcabfactalmz.codsocio = rsocios.codsocio "
    frmBas.CampoCP = "numfactu"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Histórico de Facturas ADV"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|2|3|4|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 2400
    
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaAportaciones(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Aportacion|1250|;S|txtAux(1)|T|Fecha Alta|1300|;S|txtAux(2)|T|Codigo|900|;S|txtAux(3)|T|Nombre Socio|3850|;S|txtAux(4)|T|Codigo|900|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(5)|T|Variedad|2300|;S|txtAux(6)|T|Campo|1200|;"
    
    frmBas.CadenaConsulta = "select numaport, fecaport, raporhco.codsocio, rsocios.nomsocio, raporhco.codvarie, variedades.nomvarie, raporhco.codcampo "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " from (raporhco inner join rsocios on raporhco.codsocio = rsocios.codsocio) inner join variedades on raporhco.codvarie = variedades.codvarie "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Numero Aportacion|N|S|||raporhco|numaport|0000000|S|"
    frmBas.Tag2 = "Fecha Alta|F|N|||raporhco|fecaport|dd/mm/yyyy||"
    frmBas.Tag3 = "Socio|N|N|1|999999|raporhco|codsocio|000000||"
    frmBas.Tag4 = "Nombre Socio|T|N|||rsocios|nomsocio|||"
    frmBas.Tag5 = "Variedad|N|N|||raporhco|codvarie|000000||"
    frmBas.Tag6 = "Variedad|T|N|||variedades|nomvarie|||"
    frmBas.Tag7 = "Campo|N|N|||raporhco|codcampo|00000000||"
    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 10
    frmBas.Maxlen3 = 10
    frmBas.Maxlen4 = 6
    frmBas.Maxlen5 = 30
    frmBas.Maxlen6 = 6
    frmBas.Maxlen7 = 20
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "(raporhco inner join rsocios on raporhco.codsocio = rsocios.codsocio) inner join variedades on raporhco.codvarie = variedades.codvarie  "
    frmBas.CampoCP = "numaport"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Aportaciones"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 4700
    
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaCamposPrev(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1100|;S|txtAux(1)|T|Socio|900|;S|txtAux(2)|T|Nombre|3800|;S|txtAux(3)|T|Variedad|2300|;S|txtAux(4)|T|Partida|2300|;S|txtAux(5)|T|Pol|650|;S|txtAux(6)|T|Parcela|850|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(7)|T|SPar|750|;S|txtAux(8)|T|Nro.Orden|1250|;"
    
    frmBas.CadenaConsulta = "select rcampos.codcampo, rcampos.codsocio, rsocios.nomsocio, variedades.nomvarie, rpartida.nomparti, rcampos.poligono, rcampos.parcela, rcampos.subparce, rcampos.nrocampo "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " from ((rcampos INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie) INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio) INNER JOIN rpartida ON rcampos.codparti = rpartida.codparti "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código|N|N|||rcampos|codcampo|00000000|S|"
    frmBas.Tag2 = "Codsocio|N|N|||rcampos|codsocio|000000||"
    frmBas.Tag3 = "Nombre Socio|T|N|||rsocios|nomsocio|||"
    frmBas.Tag4 = "Variedad|T|N|||variedades|nomvarie|||"
    frmBas.Tag5 = "Partida|T|N|||rpartida|nomparti|||"
    frmBas.Tag6 = "Poligono|N|N|||rcampos|poligono|000||"
    frmBas.Tag7 = "Parcela|N|N|||rcampos|parcela|000000||"
    frmBas.Tag8 = "Subparcela|T|N|||rcampos|subparce|||"
    frmBas.Tag9 = "Nro Orden|N|N|||rcampos|nrocampo|000000||"
    frmBas.Maxlen1 = 8
    frmBas.Maxlen2 = 6
    frmBas.Maxlen3 = 35
    frmBas.Maxlen4 = 25
    frmBas.Maxlen5 = 25
    frmBas.Maxlen6 = 3
    frmBas.Maxlen7 = 6
    frmBas.Maxlen8 = 6
    frmBas.Maxlen9 = 6
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "(((rcampos INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie) INNER JOIN rsocios ON rcampos.codsocio = rsocios.codsocio) INNER JOIN rpartida ON rcampos.codparti = rpartida.codparti "
    frmBas.CampoCP = "codcampo"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Campos - Huertos"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 6900
    
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaEntradasClasificadasPrev(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Nota|1000|;S|txtAux(1)|T|Fecha|1300|;S|txtAux(2)|T|Codigo|950|;S|txtAux(3)|T|Socio|4100|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(4)|T|Variedad|2500|;S|txtAux(5)|T|Campo|1250|;"
    
    frmBas.CadenaConsulta = "select rclasifica.numnotac, rclasifica.fechaent, rclasifica.codsocio, rsocios.nomsocio, variedades.nomvarie, rclasifica.codcampo "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " from (rclasifica left join rsocios on rclasifica.codsocio = rsocios.codsocio) "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " left join variedades on rclasifica.codvarie = variedades.codvarie "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Nro.Nota|N|N|||rclasifica|numnotac|0000000|S|"
    frmBas.Tag2 = "Fecha|F|N|||rsocios|fechaent|dd/mm/yyyy||"
    frmBas.Tag3 = "Codigo Socio|N|N|||rclasifica|codsocio|000000||"
    frmBas.Tag4 = "Nombre Socio|T|N|||rsocios|nomsocio|||"
    frmBas.Tag5 = "Variedad|T|N|||variedades|nomvarie|||"
    frmBas.Tag6 = "Campo|N|N|||rclasifica|codcampo|00000000||"
    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 10
    frmBas.Maxlen3 = 6
    frmBas.Maxlen4 = 35
    frmBas.Maxlen5 = 25
    frmBas.Maxlen6 = 8
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "(rclasifica left join rsocios on rclasifica.codsocio = rsocios.codsocio) left join variedades on rclasifica.codvarie = variedades.codvarie "
    frmBas.CampoCP = "numnotac"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Clasificación de Campos"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 4100
    
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaSociosPrev(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|900|;S|txtAux(1)|T|Nombre|4210|;S|txtAux(2)|T|Nif|1550|;S|txtAux(3)|T|Teléfono|1500|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(4)|T|Móvil|1500|;S|txtAux(5)|T|Fecha Alta|1350|;"
    
    frmBas.CadenaConsulta = "select codsocio, nomsocio, nifsocio, telsoci1, movsocio, fechaalta  "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " from rsocios "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código|N|N|||rsocios|codsocio|000000|S|"
    frmBas.Tag2 = "Nombre|T|N|||rsocios|nomsocio|||"
    frmBas.Tag3 = "Nif Socio|T|N|||rsocios|nifsocio|||"
    frmBas.Tag4 = "Teléfono|T|N|||rsocios|telsoci1|||"
    frmBas.Tag5 = "Móvil|T|N|||rsocios|movsocio|||"
    frmBas.Tag6 = "Fecha Alta|F|N|||rsocios|fechaalta|dd/mm/yyyy||"
    frmBas.Maxlen1 = 6
    frmBas.Maxlen2 = 35
    frmBas.Maxlen3 = 15
    frmBas.Maxlen4 = 15
    frmBas.Maxlen5 = 15
    frmBas.Maxlen6 = 10
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "rsocios"
    frmBas.CampoCP = "codsocio"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Socios"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 4000
    
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaOrdenesRecogidaPrev(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Nro.Orden|1400|;S|txtAux(1)|T|Fecha|1400|;S|txtAux(2)|T|Nro.Campo|1400|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(3)|T|Variedad|1100|;S|txtAux(4)|T|Descripción|2700|;"
    
    frmBas.CadenaConsulta = "select nroorden, fecimpre, nrocampo, rordrecogida.codvarie, variedades.nomvarie "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " from rordrecogida inner join variedades on rordrecogida.codvarie = variedades.codvarie "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Nro.Nota|N|S|||rordrecogida|nroorden|000000000|S|"
    frmBas.Tag2 = "Fecha Impresión|F|N|||rordrecogida|fecimpre|dd/mm/yyyy||"
    frmBas.Tag3 = "Nro.Campo|N|N|||rordrecogida|nrocampo|00000000||"
    frmBas.Tag4 = "Variedad|N|N|1|999999|rordrecogida|codvarie|000000||"
    frmBas.Tag5 = "Variedad|T|N|||variedades|nomvarie|||"
    frmBas.Tag6 = ""
    frmBas.Maxlen1 = 8
    frmBas.Maxlen2 = 10
    frmBas.Maxlen3 = 8
    frmBas.Maxlen4 = 6
    frmBas.Maxlen5 = 25
    frmBas.Maxlen6 = 0
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "rordrecogida inner join variedades on rordrecogida.codvarie = variedades.codvarie "
    frmBas.CampoCP = "nroorden"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Ordenes de Recogida"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 1000
    
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaPOZHidrantesPrev(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Hidrante|1600|;S|txtAux(1)|T|Socio|1100|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(2)|T|Nombre|4200|;S|txtAux(3)|T|Orden|1400|;"
    
    frmBas.CadenaConsulta = "select rpozos.hidrante, rpozos.codsocio, rsocios.nomsocio,  rpozos.nroorden "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " from rpozos inner join rsocios on rpozos.codsocio = rsocios.codsocio  "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Numero Hidrante|T|N|||rpozos|hidrante||S|"
    frmBas.Tag2 = "Socio|N|N|1|999999|rpozos|codsocio|000000||"
    frmBas.Tag3 = "Nombre Socio|T|N|||rsocios|nomsocio|||"
    frmBas.Tag4 = "Nro.Orden|N|S|0|9999999|rpozos|nroorden|0000000||"
    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 6
    frmBas.Maxlen3 = 35
    frmBas.Maxlen4 = 7
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "rpozos inner join rsocios on rpozos.codsocio = rsocios.codsocio"
    frmBas.CampoCP = "hidrante"
    frmBas.TipoCP = "N"
    'frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "Hidrantes"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 1300
    
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaPOZRecibosPrev(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Tipo Fichero|3000|;S|txtAux(1)|T|Tipo|750|;S|txtAux(2)|T|Factura|1200|;S|txtAux(3)|T|Fecha|1350|;S|txtAux(4)|T|Lin|800|;"
    frmBas.CadenaTots = frmBas.CadenaTots & "S|txtAux(5)|T|Codigo|1100|;S|txtAux(6)|T|Nombre|4800|;"
    
    frmBas.CadenaConsulta = "select case rrecibpozos.codtipom when ""RCP"" then ""RCP-Consumo"" when ""RMP"" then ""RMP-Mantenim"" when ""RVP"" then ""RVP-Contadores"" when ""RMT"" then ""RMT-Manta"" when ""TAL"" then ""TAL-Talla"" when ""RRC"" then ""RRC-Rect Consumo""  when ""RRM"" then ""RRM-Rect Mto"" when ""RRT"" then ""RRT-Rect Manta""  when ""RRV"" then ""RRV-Rect Cont"" when ""RTA"" then ""RTA-Rect Talla""  when ""FIN"" then ""FIN-Internas"" end as tipo, rrecibpozos.codtipom, "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " rrecibpozos.numfactu, rrecibpozos.fecfactu, rrecibpozos.numlinea, rrecibpozos.codsocio, rsocios.nomsocio from rrecibpozos inner join rsocios on rrecibpozos.codsocio = rsocios.codsocio "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código|T|N|||rrecibpozos|codtipom||S|"
    frmBas.Tag2 = ""
    frmBas.Tag3 = "Nro.Factura|N|N|||rrecibpozos|numfactu|0000000||"
    frmBas.Tag4 = "Fecha|F|N|||rrecibpozos|fecfactu|dd/mm/yyyy||"
    frmBas.Tag5 = "Linea|N|N|||rrecibpozos|numlinea|00||"
    frmBas.Tag6 = "Codigo|N|N|||rrecibpozos|codsocio|000000||"
    frmBas.Tag7 = "Nombre|T|N|||rsocios|nomsocio|||"
    frmBas.Maxlen1 = 3
    frmBas.Maxlen2 = 0
    frmBas.Maxlen3 = 7
    frmBas.Maxlen4 = 10
    frmBas.Maxlen5 = 2
    frmBas.Maxlen6 = 6
    frmBas.Maxlen7 = 35
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "rrecibpozos inner join rsocios on rrecibpozos.codsocio = rsocios.codsocio"
    frmBas.CampoCP = "codtipom"
    frmBas.TipoCP = "N"
    If vParamAplic.Cooperativa = 10 Then
        frmBas.Caption = "Duplicado de Recibos"
    Else
        frmBas.Caption = "Histórico de Recibos "
    End If
    
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "1|2|3|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 6000
    
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaAlbaranesRetiradaBodPrev(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Albaran|1200|;S|txtAux(1)|T|Codigo|1050|;S|txtAux(2)|T|Nombre Socio|4500|;S|txtAux(3)|T|Fecha|1350|;"
    frmBas.CadenaTots = frmBas.CadenaTots & ""
    
    frmBas.CadenaConsulta = "select numalbar, rbodalbaran.codsocio, rsocios.nomsocio, rbodalbaran.fechaalb  "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " from rbodalbaran inner join rsocios on rbodalbaran.codsocio = rsocios.codsocio "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "NºAlbarán|N|S|||rbodalbaran|numalbar|0000000|S|"
    frmBas.Tag2 = "Socio|N|N|0|999999|rbodalbaran|codsocio|000000||"
    frmBas.Tag3 = "Nombre Socio|T|N|||rsocios|nomsocio|||"
    frmBas.Tag4 = "Fecha Albaran|F|N|||rbodalbaran|fechaalb|dd/mm/yyyy|N|"
    frmBas.Tag5 = ""
    frmBas.Tag6 = ""
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 6
    frmBas.Maxlen3 = 35
    frmBas.Maxlen4 = 10
    frmBas.Maxlen5 = 0
    frmBas.Maxlen6 = 0
    frmBas.Maxlen7 = 0
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "rbodalbaran inner join rsocios on rbodalbaran.codsocio = rsocios.codsocio "
    frmBas.CampoCP = "numalbar"
    frmBas.TipoCP = "N"
    frmBas.Caption = "Albaranes de Retirada de Vino y de Aceite"
    
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 1100
    
    frmBas.Show vbModal
    
End Sub




Public Sub AyudaFacturasBodPrev(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Tipo|1000|;S|txtAux(1)|T|NºFactura|1250|;S|txtAux(2)|T|Fecha|1400|;S|txtAux(3)|T|Código|1000|;S|txtAux(4)|T|Nombre Socio|4500|;"
    frmBas.CadenaTots = frmBas.CadenaTots & ""
    
    frmBas.CadenaConsulta = "select codtipom, numfactu, fecfactu, codsocio, nomsocio "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " from rbodfacturas "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Tipo|T|N|||rbodfacturas|codtipom||S|"
    frmBas.Tag2 = "Nº Factura|N|N|||rbodfacturas|numfactu|0000000|S|"
    frmBas.Tag3 = "Fecha Fra|F|N|||rbodfacturas|fecfactu|dd/mm/yyyy|S|"
    frmBas.Tag4 = "Socio|N|N|0|999999|rbodfacturas|codsocio|000000|S|"
    frmBas.Tag5 = "Nombre Socio|T|N|||rbodfacturas|nomsocio||N|"
    frmBas.Tag6 = ""
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 3
    frmBas.Maxlen2 = 7
    frmBas.Maxlen3 = 10
    frmBas.Maxlen4 = 6
    frmBas.Maxlen5 = 35
    frmBas.Maxlen6 = 0
    frmBas.Maxlen7 = 0
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "rbodfacturas"
    frmBas.CampoCP = "numfactu"
    frmBas.TipoCP = "N"
    frmBas.Caption = "Histórico de Facturas Retirada Bodega/Almazara"
    
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 2150
    
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaFamiliasADVPrev(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Tipo|1000|;S|txtAux(1)|T|Nombre|3500|;S|txtAux(2)|T|Cta.Ventas|1400|;S|txtAux(3)|T|Cta.Compras|1400|;"
    frmBas.CadenaTots = frmBas.CadenaTots & ""
    
    frmBas.CadenaConsulta = "select codfamia, nomfamia, ctaventa, ctacompr "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " from advfamia "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código|N|N|0|9999|advfamia|codfamia|0000|S|"
    frmBas.Tag2 = "Denominación|T|N|||advfamia|nomfamia||N|"
    frmBas.Tag3 = "Cta.Ventas|T|N|||advfamia|ctaventa||N|"
    frmBas.Tag4 = "Cta.Compras|T|N|||advfamia|ctacompr||N|"
    frmBas.Tag5 = ""
    frmBas.Tag6 = ""
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 35
    frmBas.Maxlen3 = 10
    frmBas.Maxlen4 = 10
    frmBas.Maxlen5 = 0
    frmBas.Maxlen6 = 0
    frmBas.Maxlen7 = 0
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "advfamia"
    frmBas.CampoCP = "codfamia"
    frmBas.TipoCP = "N"
    frmBas.Caption = "Familias de Artículos de ADV"
    
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 300
    
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaFacturasADVPrev(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Tipo|1000|;S|txtAux(1)|T|Factura|1300|;S|txtAux(2)|T|Fecha|1400|;S|txtAux(3)|T|Código|1000|;S|txtAux(4)|T|Nombre|4500|;"
    frmBas.CadenaTots = frmBas.CadenaTots & ""
    
    frmBas.CadenaConsulta = "select codtipom, numfactu, fecfactu, codsocio, nomsocio "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " from advfacturas "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Tipo|T|N|||advfacturas|codtipom||S|"
    frmBas.Tag2 = "Nº Factura|N|N|||advfacturas|numfactu|0000000|S|"
    frmBas.Tag3 = "Fecha Factura|F|N|||advfacturas|fecfactu|dd/mm/yyyy|S|"
    frmBas.Tag4 = "Socio|N|N|0|999999|advfacturas|codsocio|000000|S|"
    frmBas.Tag5 = "Nombre Socio|T|N|||advfacturas|nomsocio||N|"
    frmBas.Tag6 = ""
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 3
    frmBas.Maxlen2 = 7
    frmBas.Maxlen3 = 10
    frmBas.Maxlen4 = 6
    frmBas.Maxlen5 = 35
    frmBas.Maxlen6 = 0
    frmBas.Maxlen7 = 0
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "advfacturas"
    frmBas.CampoCP = "numfactu"
    frmBas.TipoCP = "N"
    frmBas.Caption = "Histórico de Facturas ADV"
    
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|2|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 2200
    
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaArticulosADVPrev(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Codigo|2000|;S|txtAux(1)|T|Nombre|3500|;S|txtAux(2)|T|EAN|1800|;"
    frmBas.CadenaTots = frmBas.CadenaTots & ""
    
    frmBas.CadenaConsulta = "select codartic, nomartic, codigoea "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " from advartic "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código de articulo|T|N|||advartic|codartic||S|"
    frmBas.Tag2 = "Nombre|T|N|||advartic|nomartic|||"
    frmBas.Tag3 = "Código de Barras|T|S|||advartic|codigoea||N|"
    frmBas.Tag4 = ""
    frmBas.Tag5 = ""
    frmBas.Tag6 = ""
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 16
    frmBas.Maxlen2 = 3500
    frmBas.Maxlen3 = 13
    frmBas.Maxlen4 = 0
    frmBas.Maxlen5 = 0
    frmBas.Maxlen6 = 0
    frmBas.Maxlen7 = 0
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "advartic"
    frmBas.CampoCP = "codartic"
    frmBas.TipoCP = "T"
    frmBas.Caption = "Artículos de ADV"
    
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 300
    
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaPartesADVPrev(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Parte|1000|;S|txtAux(1)|T|Fecha|1400|;S|txtAux(2)|T|Socio|1200|;S|txtAux(3)|T|Nombre|3800|;"
    frmBas.CadenaTots = frmBas.CadenaTots & ""
    
    frmBas.CadenaConsulta = "select numparte, fechapar, advpartes.codsocio, rsocios.nomsocio "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " from advpartes inner join rsocios on advpartes.codsocio = rsocios.codsocio "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " where (1=1)  "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "NºParte|N|S|||advpartes|numparte|0000000|S|"
    frmBas.Tag2 = "Fecha Parte|F|N|||advpartes|fechapar|dd/mm/yyyy|N|"
    frmBas.Tag3 = "Socio|N|N|0|999999|advpartes|codsocio|000000||"
    frmBas.Tag4 = "Nombre Socio|T|N|||rsocios|nomsocio|||"
    frmBas.Tag5 = ""
    frmBas.Tag6 = ""
    frmBas.Tag7 = ""
    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 10
    frmBas.Maxlen3 = 6
    frmBas.Maxlen4 = 35
    frmBas.Maxlen5 = 0
    frmBas.Maxlen6 = 0
    frmBas.Maxlen7 = 0
    
    frmBas.pConn = cAgro
    
    frmBas.tabla = "advpartes inner join rsocios on advpartes.codsocio = rsocios.codsocio"
    frmBas.CampoCP = "numparte"
    frmBas.TipoCP = "N"
    frmBas.Caption = "Entrada de Partes"
    
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 400
    
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaCartasPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Nombre|6595|;"
    frmCom.CadenaConsulta = "SELECT scartas.codcarta, scartas.descarta "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scartas "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Cod. Carta|N|N|0|999|scartas|codcarta|000|S|"
    frmCom.Tag2 = "Descripción|T|S|||scartas|descarta||N|"
    frmCom.Tag3 = ""
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 15
    frmCom.Maxlen3 = 0
    
    frmCom.pConn = cAgro
    
    frmCom.tabla = "scartas"
    frmCom.CampoCP = "codcarta"
    frmCom.TipoCP = "N"
    frmCom.Caption = "Cartas / SMS"
    frmCom.DeConsulta = True

    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 500
    
    frmCom.Show vbModal
End Sub





Private Sub Redimensiona(frmBas As frmBasico2, Cant As Integer)
    frmBas.Width = frmBas.Width + Cant
    frmBas.DataGrid1.Width = frmBas.DataGrid1.Width + Cant
    frmBas.cmdAceptar.Left = frmBas.cmdAceptar.Left + Cant
    frmBas.cmdCancelar.Left = frmBas.cmdCancelar.Left + Cant
    frmBas.cmdRegresar.Left = frmBas.cmdRegresar.Left + Cant

End Sub
