Attribute VB_Name = "modMenuClick"
Option Explicit

Dim DeTransporte As Boolean
Dim frmBas As frmBasico


Private Sub Construc(nom As String)
    MsgBox nom & ": en construcció..."
End Sub

' ******* DATOS BASICOS *********

Public Sub SubmnP_Generales_Click(Index As Integer)

    Select Case Index
        Case 1: frmConfParamGral.Show vbModal
                PonerDatosPpal
        Case 2: frmConfParamAplic.Show vbModal
        Case 3: conn.Close
                If AbrirConexionUsuarios Then
                    frmConfTipoMov.Show vbModal
                    CerrarConexionUsuarios
                End If
                If AbrirConexion() = False Then
                    MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
                    End
                End If
        Case 4: frmConfParamRpt.Show vbModal
        
        Case 6: frmMantenusu.Show vbModal ' mantenimiento de usuarios
        Case 10: End
    End Select
End Sub

Public Sub SubmnC_RecoleccionG_Infor2_Click(Index As Integer)
    Select Case Index
        Case 1: AbrirListado (42)       ' listado de socios
    End Select
End Sub




Public Sub submnM_Generales_click(Index As Integer)
    Select Case Index
        Case 1: 'frmManAlmProp.Show vbModal 'almacenes propios
        Case 2: 'frmManTipUnid.Show vbModal 'tipos de unidad
        Case 3: 'frmManTipArtic.Show vbModal 'tipos de articulos
        Case 4: 'frmManFamilias.Show vbModal 'familias
        Case 5: 'frmManArtic.Show vbModal 'articulos
    End Select
End Sub


' *******  RECOLECCION  *********

Public Sub SubmnC_RecoleccionG_Admon_Click(Index As Integer)
    Select Case Index
        Case 1: frmManCoope.Show vbModal    ' cooperativas
        Case 2: frmManSeccion.Show vbModal  ' secciones
        Case 3: frmManSituacion.Show vbModal 'Situaciones especiales
        Case 4: frmManSocios.Show vbModal 'Socios
        
        Case 6: frmManZonas.Show vbModal    ' zonas de cultivo
        Case 7: frmManPueblos.Show vbModal  ' pueblos
        Case 8: frmManPartidas.Show vbModal ' partidas
        Case 9: frmManTranspor.Show vbModal 'transportistas
        Case 10: frmManTarTra.Show vbModal  ' tarifas de transporte
        Case 11: frmManCapataz.Show vbModal 'Capataces
        Case 12: frmManSituCamp.Show vbModal 'Situaciones Campos
        Case 13: frmManInciden.Show vbModal 'Incidencias
        Case 14: frmManCalidades.Show vbModal 'Calidades
        Case 15: frmManPlantacion.Show vbModal 'marco de Plantacion
        Case 16: frmManTierra.Show vbModal ' tipos de tierra
        Case 17: frmManDesarrollo.Show vbModal ' desarrollo vegetativo
        Case 18: frmManConcepGasto.Show vbModal ' conceptos de gastos
        Case 19: frmManProceRiego.Show vbModal ' procedencia de riego
        Case 20: frmManPatronaPie.Show vbModal ' patron a pie
        Case 21: frmManSeguroOpc.Show vbModal ' seguro opcion
        Case 22: frmManCampos.Show vbModal 'Campos
    
        Case 24: frmManPortesPobla.Show vbModal 'Portes por Poblacion
        Case 25: frmManCalibrador.Show vbModal 'Nombres Calibradores Catadau
        Case 26: frmManDepositos.Show vbModal ' Mantenimiento de Depositos (Mogente)
        Case 27: frmManCooprop.Show vbModal ' mantenimiento de coopropiedades
        Case 28: frmManBonifEnt.Show vbModal ' mantenimiento de bonificacion para entradas bascula
        
        Case 29: 'AbrirFormularioGlobalGAP ' mantenimiento de globalgap
                 frmManGlobalGap.Show vbModal ' Mantenimiento de globalgap
        '[Monica]21/10/2013: Añadimos las variedades de comercial
        Case 30: frmManVariedad.Show vbModal
    
    End Select
End Sub


' *******  INFORMES  *********

Public Sub SubmnC_RecoleccionG_Infor_Click(Index As Integer)
    Select Case Index
        Case 1: AbrirListado (13)       ' listado de socios por seccion
        Case 3: AbrirListadoOfer (305)  ' Etiquetas a socios
        Case 4: AbrirListadoOfer (306)  ' Cartas a Socios
        
        Case 6: AbrirListado (15)       ' informe de Campos/Huerto
        Case 7: frmListSuperficies.Show vbModal ' informe de superficies de cultivo y edad de las plantaciones
        Case 9: AbrirListado (19)       ' Grabacion de fichero agriweb
        Case 10: AbrirListado (20)       ' Informe de Kilos por Producto
        Case 11: AbrirListado (25)       ' Informe de Kilos Recolectados Socio/Cooperativa
    
    End Select
End Sub

' *******  TOMA DE DATOS *********

Public Sub SubmnC_RecoleccionG_TomaDatos_Click(Index As Integer)
    Select Case Index
        Case 1: AbrirListadoTomaDatos (1) ' Construc ("Informe de toma de datos")   'Informe de toma de datos
        Case 2: frmEntKilosEst.Show vbModal 'Construc ("Entrada de kilos estimados") 'Entrada de kilos estimados
        Case 3: AbrirListadoTomaDatos (2) 'Construc ("Informe de desviacion de aforos") 'Informe de desviacion de aforos
        Case 4: AbrirListadoTomaDatos (3) 'Construc ("Informe de clasificacion socios") 'Informe de clasificacion socios
    End Select
End Sub

Public Sub SubmnC_RecoleccionPOZOS_Click(Index As Integer)
    Select Case Index
        Case 2: AbrirListadoPOZ (15)      ' listado de diferencias
        Case 3: AbrirListadoPOZ (16)      ' listado de cuentas bancarias erroneas
    End Select
End Sub

' *******  INFORMES OFICIALES *********

Public Sub SubmnC_RecoleccionG_InforOfi_Click(Index As Integer)
    Select Case Index
        Case 2: AbrirListado (43)       ' INFORME de miembros atria
    
        Case 3: AbrirListado (46) ' registro de fitosanitarios
    End Select
End Sub





' *******  ORDENES DE RECOLECCION  *********

Public Sub SubmnC_RecoleccionG_EntradasOrd_Click(Index As Integer)
    Select Case Index
        Case 1: AbrirListado (40)   ' Ordenes de Recoleccion
        Case 2: frmManOrdRecogida.Show vbModal ' mantenimiento de incidencias dentro de las ordenes
        Case 3: AbrirListado (41)   ' Ordenes de Recoleccion Emitidas
    End Select
End Sub


' *******  ENTRADAS PESADA  *********

Public Sub SubmnC_RecoleccionG_EntradasP_Click(Index As Integer)
    Select Case Index
        Case 1: frmEntPesada.Show vbModal    ' Mantenimiento de entradas de pesadas
        Case 2: AbrirListado (11)            'Listado de entradas de pesada
    End Select
End Sub

' *******  ENTRADAS EN BASCULA  *********

Public Sub SubmnC_RecoleccionG_Entradas_Click(Index As Integer)
    Select Case Index
        Case 1:
                If vParamAplic.Cooperativa = 7 Then ' Mantenimiento de entradas en bascula de quatretonda
                    frmEntBasculaQua.Show vbModal
                Else
                    If vParamAplic.Cooperativa = 9 Then
                        frmEntBasculaNat.Show vbModal
                    Else
                        frmEntBascula.Show vbModal    ' Mantenimiento de entradas en bascula
                    End If
                End If
        Case 2: AbrirListado (10)             ' Reimpresion de entradas de bascula
        Case 3: AbrirListado (14)             ' Listado de entradas en bascula
        Case 4: frmActEntradas.Show vbModal   ' Actualizacion de entradas, grabamos rclasifica
        
        Case 8: frmManPreClasifica.Show vbModal  ' Mantenimiento de preclasificacion de Anna
        Case 9: frmManClasifica.Show vbModal  ' Mantenimiento de clasificacion
        Case 10: AbrirListado (16)             ' Listado de Entradas Clasificadas
        Case 11: frmActClasifica.Show vbModal  ' Actualizacion de entradas clasificadas
        
        '[Monica]11/04/2013: dependiendo de si es Montifrut sacamos el hco de entradas general o no
        Case 13: ' Mantenimiento de hco de entradas
                 If vParamAplic.Cooperativa = 12 Then
                    frmManHcoFrutaMon.Show vbModal
                 Else
                    frmManHcoFruta.Show vbModal
                 End If
                 
        Case 14: AbrirListado (17)            ' Reimpresion de albaranes de clasificacion
        Case 15: AbrirListado (18)            ' informe de kilos /gastos --> rhisfruta
        Case 16: frmInfEntradasSocios.Show vbModal ' informe de kilos por socios de todas las tablas
        Case 17: AbrirListado (33)          ' Informe de gastos por concepto
        Case 18: frmInfAlbDestrios.Show vbModal ' Informe de destrios varios
    
        Case 20: frmVentaFruta.Show vbModal ' mantenimiento de venta fruta
        Case 21: ' hco de entrada fruta
                 frmManHcoFruta.Show vbModal
    
    End Select
End Sub

' *******  CLASIFICACION AUTOMATICA  *********

Public Sub SubmnC_RecoleccionG_Clasifica_Click(Index As Integer)
    Select Case Index
        Case 1: AbrirListado (21)             ' Traspaso desde el calibrador
        Case 2: AbrirListado (22)             ' Traspaso entradas trazabilidad
        Case 3: ' Picassent tiene distinto formulario
                If vParamAplic.Cooperativa = 2 Then
                    frmManClasAutoPic.Show vbModal   ' Mantenimiento de clasificacion automatica
                Else
                    frmManClasAuto.Show vbModal   ' Mantenimiento de clasificacion automatica
                End If
        Case 4 ' Control Destrio Picassent
                frmControlDestrio.Show vbModal ' Mantenimiento de control de destrio
                
    End Select
End Sub



' *******  PAGO A SOCIOS *********

Public Sub SubmnC_RecoleccionG_PagoSocios_Click(Index As Integer)
    Select Case Index
        Case 1: frmManPrecios.Tipo = 0
                frmManPrecios.Show vbModal    ' Mantenimiento de precios
        Case 2: frmCalculoPrecios.Show vbModal ' calculo de precios
        
        Case 4: frmManVtasCampo.Show vbModal  ' Mantenimiento de Ventas Campo
        
        Case 16: frmFactRectifSocio.Show vbModal   ' Generacion de Facturas Rectificativas
        
    
        Case 18: frmContrRecFact.Show vbModal  'Facturacion Contratos
    
    
    End Select
End Sub


' ******* ANTICIPOS SOCIOS *********

Public Sub SubmnC_RecoleccionG_Anticipos_Click(Index As Integer)

    frmListAnticipos.AnticipoGastos = False ' no son anticipos de gastos
    frmListAnticipos.LiquidacionIndustria = False ' no es liquidacion de industria
    frmListAnticipos.AnticipoGenerico = False ' no son anticipos genericos
    
    Select Case Index
        Case 1: AbrirListadoAnticipos (2) 'Construc ("Prevision de Pago") ' listado de Prevision de pago de anticipos
        Case 2: AbrirListadoAnticipos (1) 'Construc ("Informe de Anticipos") ' informe de anticipos
        Case 3: 'Construc ("Facturacion") 'facturacion de anticipos
                DesBloqueoManual ("FACANT")
                If Not BloqueoManual("FACANT", "1") Then
                    MsgBox "No se puede realizar la Facturación de Anticipos. Hay otro usuario realizándola.", vbExclamation
                    Screen.MousePointer = vbDefault
                Else
                    frmListAnticipos.AnticipoGastos = False ' no son anticipos de gastos
                    AbrirListadoAnticipos (3)
                End If
        
        Case 5: AbrirListadoAnticipos (5) 'Construc ("Deshacer Facturas de Anticipos") ' Deshacer proceso de facturacion de anticipos
                
    End Select
End Sub


' ******* ANTICIPOS GASTOS RECOLECCION *********

Public Sub SubmnC_RecoleccionG_AnticiposGastos_Click(Index As Integer)
    frmListAnticipos.LiquidacionIndustria = False ' no es liquidacion de industria
    frmListAnticipos.AnticipoGenerico = False ' no es anticipo generico
    
    Select Case Index
        Case 1: frmListAnticipos.AnticipoGastos = True
                AbrirListadoAnticipos (2) 'Construc ("Prevision de Pago") ' listado de Prevision de pago de anticipos
        Case 2: 'Construc ("Facturacion") 'facturacion de anticipos
                DesBloqueoManual ("GASANT")
                If Not BloqueoManual("GASANT", "1") Then
                    MsgBox "No se puede realizar la Facturación de Anticipos Gastos de Recolección. Hay otro usuario realizándola.", vbExclamation
                    Screen.MousePointer = vbDefault
                Else
                    frmListAnticipos.AnticipoGastos = True
                    AbrirListadoAnticipos (3)
                End If
        
'        Case : AbrirListadoAnticipos (5) 'Construc ("Deshacer Facturas de Anticipos") ' Deshacer proceso de facturacion de anticipos
                
    End Select
End Sub


' ******* ANTICIPOS GENERICOS *********
'
' Los anticipos genericos son anticipos que se realizan sobre los kilos clasificados o del hco
' todos a un mismo precio sin depender de la calidad, ya que puede que no esten clasificados
'

Public Sub SubmnC_RecoleccionG_AnticiposGene_Click(Index As Integer)
    frmListAnticipos.LiquidacionIndustria = False ' no es liquidacion de industria
    frmListAnticipos.AnticipoGastos = False ' no es anticipo de gastos
    Select Case Index
        Case 1: frmListAnticipos.AnticipoGenerico = True
                AbrirListadoAnticipos (2) 'Construc ("Prevision de Pago") ' listado de Prevision de pago de anticipos
        Case 2: 'Construc ("Facturacion") 'facturacion de anticipos
                DesBloqueoManual ("GASANT")
                If Not BloqueoManual("GASANT", "1") Then
                    MsgBox "No se puede realizar la Facturación de Anticipos Gastos de Recolección. Hay otro usuario realizándola.", vbExclamation
                    Screen.MousePointer = vbDefault
                Else
                    frmListAnticipos.AnticipoGenerico = True
                    AbrirListadoAnticipos (3)
                End If
        
'        Case : AbrirListadoAnticipos (5) 'Construc ("Deshacer Facturas de Anticipos") ' Deshacer proceso de facturacion de anticipos
                
    End Select
End Sub





' ******* LIQUIDACIONES SOCIOS : APORTACION FONDO OPERATIVO SOLO PARA PICASSENT *********

Public Sub SubmnC_RecoleccionG_LiquidFO_Click(Index As Integer)
    Select Case Index
        Case 1: frmListAnticipos.OpcionListado = 18 ' REPARTO DE LA APORTACION DE FONDO OPERATIVO
                frmListAnticipos.Show vbModal
        
        Case 2: frmManIngresos.Show vbModal ' MANENTIMIENTO DE INGRESOS DE LIQUIDACION (PICASSENT)
                
    End Select
End Sub



' ******* LIQUIDACIONES SOCIOS *********

Public Sub SubmnC_RecoleccionG_Liquidaciones_Click(Index As Integer)

    frmListAnticipos.AnticipoGastos = False ' no son gastos de recoleccion
    frmListAnticipos.LiquidacionIndustria = False ' no es un liquidacion de industria
    frmListAnticipos.AnticipoGenerico = False ' no son anticipos genericos
    

    Select Case Index
        Case 1: AbrirListadoAnticipos (13) 'Construc ("Prevision de Pago") 'AbrirListadoAnticipos (2) '' listado de Prevision de pago de anticipos
        Case 2: AbrirListadoAnticipos (12)  'Construc ("Informe de Anticipos") 'AbrirListadoAnticipos (1) ' ' informe de anticipos
        Case 3: 'Construc ("Facturacion") 'facturacion de liquidaciones
                DesBloqueoManual ("FACLIQ")
                If Not BloqueoManual("FACLIQ", "1") Then
                    MsgBox "No se puede realizar la Facturación de Liquidaciones. Hay otro usuario realizándola.", vbExclamation
                    Screen.MousePointer = vbDefault
                Else
                    AbrirListadoAnticipos (14)
                End If
        
        Case 5:  AbrirListadoAnticipos (15)  ' Construc ("Deshacer Facturas de Liquidaciones") 'AbrirListadoAnticipos (5) 'Construc ("Deshacer Facturas de Liquidaciones") ' Deshacer proceso de facturacion de anticipos
                
    End Select
End Sub



' ******* LIQUIDACIONES INDUSTRIA SOCIOS *********

Public Sub SubmnC_RecoleccionG_LiquIndustria_Click(Index As Integer)

    frmListAnticipos.AnticipoGastos = False ' no son gastos de recoleccion
    frmListAnticipos.AnticipoGenerico = False ' no son anticipos genericos
    frmListAnticipos.LiquidacionIndustria = True ' estamos en la liquidacion de industria
    
    Select Case Index
        Case 1: AbrirListadoAnticipos (13) 'Construc ("Prevision de Pago")
         
        Case 2: AbrirListadoAnticipos (12) 'Informe de liquidacion en ppio solo lo utilizará Catadau
        
        Case 3: 'Construc ("Facturacion") 'facturacion de liquidaciones de INDUSTRIA
                DesBloqueoManual ("FACLIQ")
                If Not BloqueoManual("FACLIQ", "1") Then
                    MsgBox "No se puede realizar la Facturación de Liquidaciones de Industria. Hay otro usuario realizándola.", vbExclamation
                    Screen.MousePointer = vbDefault
                Else
                    AbrirListadoAnticipos (14)
                End If
    End Select
End Sub



' *******  FACTURAS SOCIOS *********

Public Sub SubmnC_RecoleccionG_FacturasSocios_Click(Index As Integer)
    Select Case Index
        Case 1:
                frmManFactSocios.Show vbModal     ' Mantenimiento de Facturas Socios
        Case 2:
                AbrirListadoAnticipos (4)         ' Reimpresion facturas de Socios
        Case 3:
                AbrirListadoOfer (315)            ' Envio de facturas por email
        Case 4:
                AbrirListadoOfer (316)            ' Facturacion web electronica
        Case 6:
                frmContaFacSoc.Show vbModal       ' Integracion contable
        Case 7:
                frmImpAridoc.Tipo = 0 ' Integracion de aridoc: Facturas de socio
                frmImpAridoc.Caption = "Exportar Facturas Socio a Aridoc"
                frmImpAridoc.Label4(16).Caption = "Fecha Factura"
                frmImpAridoc.Show vbModal
        Case 9:
                AbrirListadoAnticipos (8)         ' Informe de Resultados
        Case 10:
                AbrirListadoAnticipos (9)         ' Listado de Retenciones
        Case 11:
                AbrirListadoAnticipos (10)        ' Grabacion de Modelo 190
        Case 12:
                AbrirListadoAnticipos (11)       ' Grabacion de Modelo 346
        Case 14:
                AbrirListadoAnticipos (20)        ' Anticipos Pendientes de Descontar
    End Select
End Sub

' *******  FACTURAS TERCEROS *********

Public Sub SubmnC_RecoleccionG_FacturasTerceros_Click(Index As Integer)
    Select Case Index
        Case 1: '3:
                frmTercRecFact.Show vbModal   ' Recepción de Facturas Terceros
        Case 2: '4:
                frmTercHcoFact.Show vbModal  ' Historico de albaran / facturas
        Case 4: '6:
                frmTercIntCont.Show vbModal  ' Integracion contable
        Case 6: '8:
                frmTercAlbPdtes.Show vbModal  ' Albaranes pendientes de facturar
        Case 7: '9:
                frmTercListFact.Show vbModal  ' Informe de facturas de terceros
    End Select
End Sub


' *******  FACTURAS VARIAS *********

Public Sub SubmnC_RecoleccionG_FacturasVarias_Click(Index As Integer)
    Select Case Index
        Case 1: 'contadores
                conn.Close
                If AbrirConexionUsuarios Then
                    frmFVARContadores.Show vbModal
                    CerrarConexionUsuarios
                End If
                If AbrirConexion() = False Then
                    MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
                    End
                End If
        Case 2: 'Conceptos
                frmFVARConceptos.Show vbModal
                
        Case 4: 'facturas
                frmFVARFacturas.Show vbModal
                
        Case 5: 'Reimpresion de facturas
                AbrirListadoFVarias 1
        
        Case 7: 'Integracion Ariconta
                AbrirListadoFVarias 4
                
        Case 10: 'diario de facturacion
                AbrirListadoFVarias 3
    End Select
End Sub



' *******  FACTURAS VARIAS PREOVEEDOR*********

Public Sub SubmnC_RecoleccionG_FacturasVariasProv_Click(Index As Integer)
    Select Case Index
        Case 1: 'facturas varias de proveedor
                frmFVARFacturasPro.Show vbModal
                
        Case 2: 'Reimpresion de facturas
                AbrirListadoFVarias 5
        
        Case 4: 'Integracion Ariconta
                AbrirListadoFVarias 7
                
        Case 5: 'Integracion Aridoc
                AbrirListadoFVarias 3
                
        Case 7: 'Diario de Facturacion
                AbrirListadoFVarias 6
    End Select
End Sub






' *******  TRANSPORTE  *********

Public Sub SubmnC_RecoleccionG_Transporte_Click(Index As Integer)
    Select Case Index
        Case 1: frmTRAFactAlb.OpcionListado = 1
                frmTRAFactAlb.Show vbModal 'frmTransListAnticipos.Show vbModal     ' facturacion de transporte
        Case 2: frmTRAFactAlb.OpcionListado = 2
                frmTRAFactAlb.Show vbModal  'frmTransListAnticipos.Show vbModal    ' reimpresion de facturas
        Case 3: frmManFactTranspor.Show vbModal 'Construc ("Hco de facturas de transporte") 'frmTransFacturas.Show vbModal ' facturas de transporte
        
        Case 4: frmTRAFactAlb.OpcionListado = 3 ' Factura a socio
                frmTRAFactAlb.Show vbModal

        Case 6: frmTRAContaFac.Show vbModal 'Construc ("Integracion contable")
        Case 7: 'Construc ("Integracion al aridoc") ' Integracion del aridoc
                frmImpAridoc.Tipo = 5 ' Integracion de aridoc: Facturas de transporte
                frmImpAridoc.Caption = "Exportar Facturas Transporte a Aridoc"
                frmImpAridoc.Label4(16).Caption = "Fecha Factura"
                frmImpAridoc.Show vbModal 'vbModalConstruc("Integracion aridoc")
    End Select
End Sub

Public Sub SubmnC_RecoleccionG_Transporte1_Click(Index As Integer)
    Select Case Index
        Case 1:  ' Informe de entradas de transporte
                frmTRAInfHcoEntradas.Show vbModal
    End Select
End Sub


' *******  GESTION DE PRENOMINA *******
Public Sub SubmnP_PreNominas_click(Index As Integer)
    Select Case Index
        Case 1: frmManSalarios.Show vbModal 'Mantenimiento de salarios
        Case 2: frmManTraba.Show vbModal  'Mantenimiento de trabajadores
        Case 3: frmManCuadrillas.Show vbModal ' Mantenimiento de cuadrillas
        Case 4: frmManCGastosNom.Show vbModal ' Construc ("Conceptos de Gastos") 'Conceptos de gastos / campo
        
        Case 5: frmManTarifaETT.Show vbModal ' Mantenimiento de tarifas ETT
        
        
        Case 7: frmManPartes.Show vbModal 'Construc ("Partes de Trabajo")
        Case 8: AbrirListadoNominas (17) 'Construc ("Pago de Partes")
        
        Case 10:
                '[Monica]29/02/2012: Natural era la cooperativa 0 junto con Catadau ahora es la 9
                If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 9 Then
                    frmManHorasDestajo.Show vbModal 'mantenimiento de horas de destajo
                Else
                    If vParamAplic.Cooperativa = 2 Then
                        frmManHorasDestajoPica.Show vbModal  'Entrada de Horas destajo picassent
                    End If
                End If
                
        Case 12: ' Solo natural tiene anticipos a trabajadores
                frmManHorasAntNat.Show vbModal
        
        Case 13: ' Natural de montaña tiene cooperativa 0
                '[Monica]29/02/2012: Natural de Montaña va a ser cooperativa 9
                If vParamAplic.Cooperativa = 9 Then
                    frmManHorasNat.Show vbModal  'Entrada de Horas de trabajadores para natural de montaña
                Else
                    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 0 Then
                        If vParamAplic.Cooperativa = 0 Then frmManHorasPica.Caption = "Entrada de Horas"
                        frmManHorasPica.Show vbModal  'Entrada de Horas de trabajadores para picassent
                    Else
                        frmManHoras.Show vbModal  'Entrada de Horas de trabajadores
                    End If
                End If
        Case 14: frmImpRecibos.Show vbModal 'Impresión de Recibos
        Case 16:
                '[Monica]29/02/2012: Natural era la cooperativa 0 junto con Catadau ahora es la 9
                 If vParamAplic.Cooperativa = 9 Then   ' caso de natural de montaña
                        frmPagoRecibos.OpcionListado = 2
                        frmPagoRecibos.Show vbModal 'Pago de Recibos
                 Else ' caso de valsur y de alzira
                        frmPagoRecibos.OpcionListado = 1
                        frmPagoRecibos.Show vbModal 'Pago de Recibos
                 End If
                 
        Case 17: frmImpAridoc.Tipo = 4 ' Integracion de aridoc: Recibos de Nóminas
                 frmImpAridoc.Caption = "Importar Recibos a Aridoc"
                 frmImpAridoc.Label4(16).Caption = "Fecha"
                 frmImpAridoc.Show vbModal 'vbModalConstruc("Integracion aridoc")
        Case 19: frmInfHorasMes.Show vbModal ' Informe de Horas Mensual
        
        Case 20: frmManHorasETT.Show vbModal ' Entrada de Horas ETT para Picassent
        
        Case 21: frmManAsesoria.Show vbModal 'Mantenimiento de diferencias asesoria
                 
        Case 22: AbrirListadoNominas (37) 'Informe mensual de horas de asesoria
                 
                 
        Case 23: AbrirListadoNominas (38) ' Rendimiento por capataz
                 
    End Select
End Sub


' *******  TRAZABILIDAD *********

Public Sub SubmnC_RecoleccionG_Trazabilidad_Click(Index As Integer)
    Select Case Index
        Case 1: frmTrzTipos.Show vbModal ' Mantenimiento de Tipos
        Case 2: frmTrzAreas.Show vbModal ' Construc ("Mantenimiento de Areas")   'Informe de toma de datos
        
        Case 4: AbrirListadoTrazabilidad (1) ' informe de palets entrados
        Case 5: AbrirListadoTrazabilidad (2) ' informe de cargas linea confeccion
        Case 6: AbrirListadoTrazabilidad (7) ' cargas linea por fecha/producto
        Case 7: AbrirListadoTrazabilidad (3) ' origenes del palet confeccionado
        Case 8: AbrirListadoTrazabilidad (4) ' destino albaranes de venta
        Case 9: AbrirListadoTrazabilidad (5) ' listado de stocks
        
        Case 11: frmTrzManPalet.Show vbModal ' Manejo de Palets
        Case 12: AbrirListadoTrazabilidad (6) ' Modificacion de cargas de confeccion
    End Select
End Sub

' *******  ALMAZARA  *********

'[Monica]20/10/2015: traspaso de campos de almazara solo para ABN
Public Sub SubmnC_RecoleccionG_Almz0_Click(Index As Integer)
    Select Case Index
        Case 1: frmAlmzTrasCampos.Show vbModal  'Construc ("Traspaso de Campos")  ' Traspaso de Campos solo para ABN
    End Select
End Sub

Public Sub SubmnC_RecoleccionG_Almazara_Click(Index As Integer)
    Select Case Index
        Case 1: frmAlmzTraspaso.Show vbModal   ' Traspaso de Almazara para Catadau
        Case 2: frmAlmzFacturas.Show vbModal   ' Hco de facturas
        Case 3: frmAlmzReimpFact.Show vbModal  ' reimpresion de facturas de almazara
        Case 4: frmAlmzContaFac.Show vbModal   ' integracion contable de facturas de almazara
    End Select
End Sub


Public Sub SubmnC_RecoleccionG_Almz_Click(Index As Integer)
    Select Case Index
        Case 3: frmAlmzTrasBascula.Show vbModal  'Construc ("Traspaso de Bascula")  ' Traspaso de Bascula para Valsur
        Case 4: frmAlmzEntradas.Show vbModal   'Construc ("Entradas bascula")  ' Hco de facturas
        Case 5: AbrirListadoBodEntradas (5) ' reparto de gastos de liquidacion almazara
        Case 6: ' mantenimiento de precios de anticipos liquidacion
                frmManPrecios.Tipo = 1 ' almazara
                frmManPrecios.Show vbModal
        Case 7: frmAlmzListEntradas.OpcionListado = 0
                frmAlmzListEntradas.Show vbModal ' informe de entradas de almazara por socios/variedad
        Case 8: frmAlmzListEntradas.OpcionListado = 1
                frmAlmzListEntradas.Show vbModal  ' extracto de entradas de almazara socio/variedad
        Case 10: frmAlmzTrasRendimiento.Show vbModal ' traspaso de rendimiento entradas
'                ' VALSUR no tiene bodega y quiere poder llamar al mantenimiento de retirada desde almazara
'        Case 10: frmBodAlbRetirada.Show vbModal ' albaranes de retirada de bodega
        Case 11: frmAlmzHcoRendimiento.Show vbModal
    End Select
End Sub

' *******  ALMAZARA : ANTICIPOS *********

Public Sub SubmnC_RecoleccionG_AlmzAnticipos_Click(Index As Integer)
    Select Case Index
        Case 1: AbrirListadoBodAnticipos (20) 'Construc ("Prevision de Pago")
        Case 2:
                DesBloqueoManual ("FACFNZ")
                If Not BloqueoManual("FACFNZ", "1") Then
                    MsgBox "No se puede realizar la Facturación de Anticipos de Almazara. Hay otro usuario realizándola.", vbExclamation
                    Screen.MousePointer = vbDefault
                Else
                    AbrirListadoBodAnticipos (30)
                End If
                
        Case 3: ' deshacer proceso de facturacion
                AbrirListadoBodAnticipos (50) 'Construc ("Deshacer Facturas de Anticipos") ' Deshacer proceso de facturacion de anticipos

    End Select
End Sub


' *******  ALMAZARA : LIQUIDACION *********

Public Sub SubmnC_RecoleccionG_AlmzLiquidacion_Click(Index As Integer)
    Select Case Index
        Case 1: AbrirListadoBodAnticipos (130) 'Construc ("Prevision de Pago") 'AbrirListadoAnticipos (2) '' listado de Prevision de pago de anticipos
        Case 2: 'Construc ("Facturacion") 'facturacion de liquidaciones
                DesBloqueoManual ("FACFLZ")
                If Not BloqueoManual("FACFLZ", "1") Then
                    MsgBox "No se puede realizar la Facturación de Liquidaciones de Almazara. Hay otro usuario realizándola.", vbExclamation
                    Screen.MousePointer = vbDefault
                Else
                    AbrirListadoBodAnticipos (140)
                End If
        
        Case 3:  AbrirListadoBodAnticipos (150)  ' Construc ("Deshacer Facturas de Liquidaciones") 'AbrirListadoAnticipos (5) 'Construc ("Deshacer Facturas de Liquidaciones") ' Deshacer proceso de facturacion de anticipos
                
        Case 4: ' informe de liquidacion de autoconsumo
                AbrirListadoBodEntradas (8)
                
    End Select
End Sub


' Retirada de almazara

Public Sub SubmnC_RecoleccionG_Almz1_Click(Index As Integer)
    Select Case Index
        Case 1: frmBodAlbRetirada.Show vbModal ' albaranes de retirada de bodega

        Case 3: ' previsio de facturacion de albaranes de retirada
                frmBodFactAlbaranes.Tipo = 0     ' Prevision Facturacion de albaranes de almazara
                frmBodFactAlbaranes.OpcionListado = 50
                frmBodFactAlbaranes.Show vbModal
                        
        Case 4:
                frmBodFactAlbaranes.Tipo = 0     ' Facturacion de albaranes de almazara
                frmBodFactAlbaranes.OpcionListado = 52
                frmBodFactAlbaranes.Label10(0).Caption = "Facturación de Albaranes Retirada Almazara"
                frmBodFactAlbaranes.Show vbModal
                
        Case 5: frmBodReimpre.Tipo = 0            ' Reimpresion de facturas de retirada (almazara)
                frmBodReimpre.Label1 = "Reimpresión de Facturas de Almazara"
                frmBodReimpre.Show vbModal
                  
        Case 6: frmBodHcoFacturas.Tipo = 0       ' Hco de Albaranes Facturas almazara
                frmBodHcoFacturas.Caption = "Histórico de Facturas Retirada Almazara"
                frmBodHcoFacturas.Show vbModal
        
        Case 8: frmBodContaFac.Tipo = 0          ' Integracion contable de facturas
                frmBodContaFac.Caption = "Integración Contable de Facturas de Retirada de Almazara"
                frmBodContaFac.Show vbModal
            
        Case 9: ' Integracion del aridoc
                frmImpAridoc.Tipo = 2 ' Integracion de aridoc: Facturas de Retirada de almazara
                frmImpAridoc.Caption = "Exportar Facturas Almazara a Aridoc"
                frmImpAridoc.Label4(16).Caption = "Fecha Factura"
                frmImpAridoc.Show vbModal 'vbModalConstruc("Integracion aridoc")
            
        Case 11: AbrirListadoBodEntradas (6) ' listado diferencia consumo/producido
        Case 12: AbrirListadoBodEntradas (9) ' diario de facturacion de retirada
    End Select
End Sub


' Retirada de Bodega

Public Sub SubmnC_RecoleccionG_Bod1_Click(Index As Integer)
    Select Case Index
        Case 1: frmBodAlbRetirada.Show vbModal ' albaranes de retirada de bodega


        Case 3: frmBodFactAlbaranes.Tipo = 1    ' Prevision Facturacion de albaranes de retirada bodega
                frmBodFactAlbaranes.OpcionListado = 50
                frmBodFactAlbaranes.Label10(0).Caption = "Previsión Facturación Retirada Bodega"
                frmBodFactAlbaranes.Show vbModal


        Case 4: frmBodFactAlbaranes.Tipo = 1    ' Facturacion de albaranes de bodega
                frmBodFactAlbaranes.OpcionListado = 52
                frmBodFactAlbaranes.Label10(0).Caption = "Facturación de Albaranes Retirada Bodega"
                frmBodFactAlbaranes.Show vbModal
        
        Case 5: frmBodReimpre.Tipo = 1              ' Reimpresion de facturas de retirada(bodega)
                frmBodReimpre.Label1 = "Reimpresión de Facturas de Bodega"
                frmBodReimpre.Show vbModal
        
        Case 6: frmBodHcoFacturas.Tipo = 1
                frmBodHcoFacturas.Caption = "Histórico de Facturas Retirada Bodega"
                frmBodHcoFacturas.Show vbModal      ' Hco de Albaranes Facturas bodega
        
        Case 8: frmBodContaFac.Tipo = 1
                frmBodContaFac.Caption = "Integración Contable de Facturas de Retirada de Bodega"
                frmBodContaFac.Show vbModal         ' Integracion contable de facturas bodega
        
        Case 9: ' Integracion del aridoc
                frmImpAridoc.Tipo = 3 ' Integracion de aridoc: Facturas de Retirada de bodega
                frmImpAridoc.Caption = "Exportar Facturas Bodega a Aridoc"
                frmImpAridoc.Label4(16).Caption = "Fecha Factura"
                frmImpAridoc.Show vbModal 'vbModalConstruc("Integracion aridoc")
        
        Case 11:  AbrirListadoBodEntradas (2)  ' listado de consumo de entradas de vino
    End Select
End Sub


' *******  ADV  *********

Public Sub SubmnC_RecoleccionG_ADV_Click(Index As Integer)
    Select Case Index
        Case 1: frmADVFamilias.Show vbModal     ' familias de articulos de adv
        Case 2: frmADVMatActivas.Show vbModal   ' Materias activas
        Case 3: frmADVArticulos.Show vbModal    ' articulos de adv
        Case 4:
                If vParamAplic.Cooperativa = 3 Then
                    frmADVTrataMoi.Show vbModal ' tipo de venta
                Else
                    frmADVTratamientos.Show vbModal ' tratamientos de adv
                End If
        Case 5: frmADVPartes.Show vbModal       ' Entradas de partes
        Case 6: frmADVReimpre.OpcionListado = 0 ' Reimpresion de partes
                frmADVReimpre.Show vbModal
                
        Case 8: frmADVFactPartes.OpcionListado = 0 ' Prevision de Facturacion de partes
                frmADVFactPartes.Show vbModal
        Case 9: frmADVFactPartes.OpcionListado = 1 ' Facturacion de partes
                frmADVFactPartes.Show vbModal
                
        Case 10: frmADVReimpre.OpcionListado = 1 ' Reimpresion de facturas de adv
                frmADVReimpre.Show vbModal
        Case 11: frmADVHcoFacturas.Show vbModal  ' Hco de Partes Facturas
        Case 13: frmADVContaFac.Show vbModal    ' Integracion contable de facturas
        Case 14: ' Integracion del aridoc
                frmImpAridoc.Tipo = 1 ' Integracion de aridoc: Facturas de adv
                frmImpAridoc.Caption = "Exportar Facturas ADV a Aridoc"
                frmImpAridoc.Label4(16).Caption = "Fecha Factura"
                frmImpAridoc.Show vbModal 'vbModalConstruc("Integracion aridoc")
    
        Case 16: ' Rendimiento por articulo
                AbrirListadoADV (0)
                
        Case 18: ' Pago de Partes por trabajador
                AbrirListadoADV (1)
                
    End Select
End Sub



' *******  BODEGA  *********

Public Sub SubmnC_RecoleccionG_Bodega_Click(Index As Integer)
    Select Case Index
        Case 1: frmBodEntradas.Show vbModal     ' entradas de bodega
        Case 2: frmBodBonifica.Show vbModal     ' tabla de bonificaciones de bodega
        Case 3: AbrirListadoBodEntradas (4)    ' reparto de gastos de liquidacion bodega
        Case 4: ' mantenimiento de precios de anticipos liquidacion
                frmManPrecios.Tipo = 2 ' bodega
                frmManPrecios.Show vbModal
        Case 5: AbrirListadoBodEntradas (0) ' informe de entradas
        Case 6: AbrirListadoBodEntradas (1)  ' extracto de entradas de bodega por socios/variedad
    End Select
End Sub


' *******  BODEGA : ANTICIPOS *********

Public Sub SubmnC_RecoleccionG_BodAnticipos_Click(Index As Integer)
    Select Case Index
        Case 1: AbrirListadoBodAnticipos (2) 'Construc ("Prevision de Pago")
        Case 2:
                DesBloqueoManual ("FACBAN")
                If Not BloqueoManual("FACBAN", "1") Then
                    MsgBox "No se puede realizar la Facturación de Anticipos de Bodega. Hay otro usuario realizándola.", vbExclamation
                    Screen.MousePointer = vbDefault
                Else
                    AbrirListadoBodAnticipos (3)
                End If
                
        Case 3: ' deshacer proceso de facturacion
                AbrirListadoBodAnticipos (5) 'Construc ("Deshacer Facturas de Anticipos") ' Deshacer proceso de facturacion de anticipos

    End Select
End Sub


' *******  BODEGA : LIQUIDACION *********

Public Sub SubmnC_RecoleccionG_BodLiquidacion_Click(Index As Integer)
    Select Case Index
        Case 1: AbrirListadoBodAnticipos (13) 'Construc ("Prevision de Pago") 'AbrirListadoAnticipos (2) '' listado de Prevision de pago de anticipos
        Case 2: 'Construc ("Facturacion") 'facturacion de liquidaciones
                DesBloqueoManual ("FACFLB")
                If Not BloqueoManual("FACFLB", "1") Then
                    MsgBox "No se puede realizar la Facturación de Liquidaciones de Bodega. Hay otro usuario realizándola.", vbExclamation
                    Screen.MousePointer = vbDefault
                Else
                    AbrirListadoBodAnticipos (14)
                End If
        
        Case 3:  AbrirListadoBodAnticipos (15)  ' Construc ("Deshacer Facturas de Liquidaciones") 'AbrirListadoAnticipos (5) 'Construc ("Deshacer Facturas de Liquidaciones") ' Deshacer proceso de facturacion de anticipos
                
    End Select
End Sub

' *******  POZOS  *********
Public Sub SubmnC_RecoleccionG_Pozos_Click(Index As Integer)
    Select Case Index
        Case 1: frmPOZPozos.Show vbModal ' tipos de pozos
                'construc ("Tipos de Pozos")
        Case 2:
                If vParamAplic.Cooperativa = 10 Or vParamAplic.Cooperativa = 8 Then
                    frmPOZHidrantesIndefa.Show vbModal
                Else
                    frmPOZHidrantes.Show vbModal
                End If
        
        Case 4: AbrirListadoPOZ (1) ' Toma de lectura de contador
                
        Case 5: frmPOZLecturas.Show vbModal ' Introduccion de lectura
        
        Case 6: AbrirListadoPOZ (2)  ' Listado de comprobacion
        Case 8: AbrirListadoPOZ (3)  ' Generacion Recibos de consumo
        Case 9: AbrirListadoPOZ (4)  ' Generacion Recibos de mantenimiento
        Case 10: AbrirListadoPOZ (5) ' Generacion Recibos de contadores
        
        '[Monica]05/05/2014: Nuevo punto de facturacion de recibos de consumo a manta
        Case 12: 'AbrirListadoPOZ (17) ' Generacion Recibos de consumo a manta
                 frmPOZMantaTickets.Show vbModal

        Case 13: ' Informe de recibos riego a manta
                AbrirListadoPOZ (19)      ' listado de recibos por fecha de riego
        
        Case 15: AbrirListadoPOZ (10) ' Listado de tallas (recibos mantenimiento) sólo para Escalona
        Case 16: AbrirListadoPOZ (11) ' Generacion Recibos de Talla (solo para Escalona)
                 
        Case 17: AbrirListadoPOZ (12) ' Calculo bonificacion recibos talla sólo para Escalona
                 
                
        Case 18: frmPOZRecibos.Show vbModal ' mantenimiento de historico de Recibos
        
        Case 20: AbrirListadoPOZ (6) ' Reimpresion de recibos de pozos
        
        Case 21: ' Impresion de carta de reclamacion para escalona
                Screen.MousePointer = vbHourglass
                frmPOZListadoOfer.OpcionListado = 1
                frmPOZListadoOfer.Show vbModal
                Screen.MousePointer = vbDefault
        
        '[Monica]11/01/2016: facturas de recargo solo para escalona
        Case 22: frmPOZFraRecargo.Show vbModal
        
        Case 24:
        
                Load frmPOZIntTesor
                If frmPOZIntTesor.Combo1(0).ListCount = 0 Then
                    Unload frmPOZIntTesor
                Else
                    frmPOZIntTesor.Show vbModal ' Integracion contable
                End If
                
        Case 25: ' Integracion aridoc
                frmImpAridoc.Tipo = 6 ' Integracion de aridoc: Facturas de socio de pozos
                frmImpAridoc.Caption = "Exportar Facturas Socio a Aridoc"
                frmImpAridoc.Label4(16).Caption = "Fecha Factura"
                frmImpAridoc.Show vbModal 'vbModalConstruc("Integracion aridoc")
                
        Case 27: AbrirListadoPOZ (7) ' Informe de Facturas generadas por Hidrante
                
        Case 28: AbrirListadoPOZ (18) ' listado de recibos pendientes de cobro
                
        Case 29: AbrirListadoPOZ (20) ' listado de recibos de consumo pendientes de cobro
                
        Case 30: AbrirListadoPOZ (8) ' Impresion de Etiquetas de contadores
                 
        Case 32: AbrirListadoPOZ (9) ' Rectificación de Lecturas
        
        Case 33: frmPOZIndefa.Show vbModal  ' prueba de indefa
    End Select
End Sub

' *******  APORTACIONES  *********

Public Sub SubmnC_RecoleccionG_Aport_Click(Index As Integer)
    Select Case Index
        Case 1: frmAPOTipos.Show vbModal  ' tipos de aportaciones
        
        Case 2: '[Monica]12/01/2012:
                ' Para el caso de Quatretonda tiene un mto de aportaciones diferente del resto de cooperativas
                If vParamAplic.Cooperativa = 7 Then
                    frmAPOMtoQua.Show vbModal
                '[Monica]25/11/2013: entra bolbaite con las aportaciones
                ElseIf vParamAplic.Cooperativa = 14 Then
                    frmAPOMtoBol.Show vbModal
                Else
                    frmAPOAportacion.Show vbModal ' mantenimiento de aportaciones
                End If
    
        Case 4: 'informe de aportaciones
                If vParamAplic.Cooperativa = 14 Then
                    AbrirListadoAPOR (4)
                Else
                    frmAPOListados.OpcionListado = 1
                    frmAPOListados.Show vbModal
                End If
                
        Case 5: frmAPOListados.OpcionListado = 2     ' Regularizaciones de aportacion
                frmAPOListados.Show vbModal
                'Construc ("Regularizaciones de aportacion")
        Case 6: ' Certificado de aportaciones
                'Construc ("Certificado de aportaciones")
                If vParamAplic.Cooperativa = 14 Then ' Bolbaite
                    AbrirListadoAPOR (15)
                Else
                    frmAPOListados.OpcionListado = 3
                    frmAPOListados.Show vbModal
                End If
    
        Case 7: ' aportacion obligatoria
                AbrirListadoAPOR (13)
                
        Case 8: ' integracion en tesoreria para bolbaite
                AbrirListadoAPOR (14)
                
        Case 9: ' devolucion aportaciones para bolbaite
                AbrirListadoAPOR (16)
                
                
    End Select
End Sub




' *******  UTILIDADES *********

Public Sub SubmnE_Util_Click(Index As Integer)
    Select Case Index
        Case 1: frmCaracteresMB.Show vbModal
        Case 3: AbrirListado (24) ' traspaso de facturas de liquidacion (valsur)
        Case 4: AbrirListado (26) ' traspaso de ropas solo para catadau
        Case 5: frmSumActSoc.Show vbModal ' Acturalizacion datos socios del Ariges
        Case 6: AbrirListado (27) ' traspaso de datos a almazara
        Case 8: frmTelTrasFras.Show vbModal ' Construc ("Trapaso facturas telefonia")
        Case 9: frmTelFacturas.Show vbModal ' Construc ("Facturas Telefonia")
        Case 10: frmTelContaFac.Show vbModal ' Construc ("Integracion contable facturas telefonia")
        Case 12: frmBackUP.Show vbModal ' Copia de seguridad previa al reparto de albaranes
        Case 13: frmRepartoAlb.Show vbModal ' Construc ("Reparto albaranes coopropietarios")
    
        Case 15: frmTrasTraza.Show vbModal ' traspaso de entradas desde traza (Castelduc)
        Case 16: AbrirListado (47) ' traspaso datos a trazabilidad (Castelduc)
        Case 17: AbrirListado (48) ' traspaso datos a trazabilidad (Castelduc)
    
        Case 19: frmLog.Show vbModal ' ver acciones
        
    End Select
End Sub

Public Sub BloqueoMenusSegunCooperativa()
Dim b As Boolean
Dim I As Integer

    ' traspaso de facturas a cooperativas solo para Valsur
    MDIppal.mnE_Util(3).Enabled = MDIppal.mnE_Util(3).visible And (vParamAplic.Cooperativa = 1)
    MDIppal.mnE_Util(3).visible = MDIppal.mnE_Util(3).visible And (vParamAplic.Cooperativa = 1)
    
    ' traspaso de ROPAS solo para Catadau
    MDIppal.mnE_Util(4).Enabled = MDIppal.mnE_Util(4).visible 'And (vParamAplic.Cooperativa = 0)
    MDIppal.mnE_Util(4).visible = MDIppal.mnE_Util(4).visible 'And (vParamAplic.Cooperativa = 0)
    
' luego lo descomento
    ' telefonia solo para Valsur
    For I = 8 To 10
        MDIppal.mnE_Util(I).Enabled = MDIppal.mnE_Util(I).visible And (vParamAplic.Cooperativa = 1)
        MDIppal.mnE_Util(I).visible = MDIppal.mnE_Util(I).visible And (vParamAplic.Cooperativa = 1)
    Next I
    
'    ' traspaso de datos a Almazara solo para Moixent
'    MDIppal.mnE_Util(6).Enabled = (vParamAplic.Cooperativa = 3)
'    MDIppal.mnE_Util(6).visible = (vParamAplic.Cooperativa = 3)
'
    'traspaso a almazara solo para catadau
    MDIppal.mnRec_Almazara(1).Enabled = MDIppal.mnRec_Almazara(1).visible And (vParamAplic.Cooperativa = 0)
    MDIppal.mnRec_Almazara(1).visible = MDIppal.mnRec_Almazara(1).visible And (vParamAplic.Cooperativa = 0)
    
    '[Monica]29/02/2012: Natural era la cooperativa 0 junto con Catadau ahora es la 9
    ' mantenimiento de calibradores solo para catadau , alzira, castelduc y picassent
    MDIppal.mnRecG_Admon(25).Enabled = MDIppal.mnRecG_Admon(25).visible And ((vParamAplic.Cooperativa = 0) Or (vParamAplic.Cooperativa = 2) Or (vParamAplic.Cooperativa = 4) Or (vParamAplic.Cooperativa = 5) Or (vParamAplic.Cooperativa = 9))
    MDIppal.mnRecG_Admon(25).visible = MDIppal.mnRecG_Admon(25).visible And ((vParamAplic.Cooperativa = 0) Or (vParamAplic.Cooperativa = 2) Or (vParamAplic.Cooperativa = 4) Or (vParamAplic.Cooperativa = 5) Or (vParamAplic.Cooperativa = 9))
    
    
    MDIppal.mnRec_AlmzLiquidacion(4).Enabled = MDIppal.mnRec_AlmzLiquidacion(4).visible And vParamAplic.Cooperativa = 1
    MDIppal.mnRec_AlmzLiquidacion(4).visible = MDIppal.mnRec_AlmzLiquidacion(4).visible And vParamAplic.Cooperativa = 1
    
    MDIppal.mnRec_LiquFO(1).Enabled = (vParamAplic.Cooperativa = 2) 'Solo para Picassent
    MDIppal.mnRec_LiquFO(1).visible = (vParamAplic.Cooperativa = 2) 'Solo para Picassent
    '[Monica]10/01/2014: incremento de liquidacion para picassent
    MDIppal.mnRec_LiquFO(2).Enabled = (vParamAplic.Cooperativa = 2) 'Solo para Picassent
    MDIppal.mnRec_LiquFO(2).visible = (vParamAplic.Cooperativa = 2) 'Solo para Picassent
    
    '[Monica]23/09/2011: la rectificacion de facturas de consumo de momento solo para Quatretonda
    MDIppal.mnRec_Pozos(28).Enabled = (vParamAplic.Cooperativa = 7)
    MDIppal.mnRec_Pozos(28).visible = (vParamAplic.Cooperativa = 7)


    '[Monica]05/05/2014: los recibos de consumo a manta solo para utxera y escalona
    MDIppal.mnRec_Pozos(12).Enabled = (vParamAplic.Cooperativa = 10)
    MDIppal.mnRec_Pozos(12).visible = (vParamAplic.Cooperativa = 10)


    '[Monica]19/06/2012: generacion de recibos de talla (solo para escalona)
    MDIppal.mnRec_Pozos(15).Enabled = (vParamAplic.Cooperativa = 10)
    MDIppal.mnRec_Pozos(15).visible = (vParamAplic.Cooperativa = 10)
    '[Monica]19/06/2012: el informe de talla (recibos de mto) solo es para Escalona
    MDIppal.mnRec_Pozos(16).Enabled = (vParamAplic.Cooperativa = 10)
    MDIppal.mnRec_Pozos(16).visible = (vParamAplic.Cooperativa = 10)
    '[Monica]19/06/2012: actualizacion recibos de mto solo es para Escalona
    MDIppal.mnRec_Pozos(17).Enabled = (vParamAplic.Cooperativa = 10)
    MDIppal.mnRec_Pozos(17).visible = (vParamAplic.Cooperativa = 10)
    '[Monica]26/11/2012: Cartas de reclamacion solo para escalona
    MDIppal.mnRec_Pozos(21).Enabled = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
    MDIppal.mnRec_Pozos(21).visible = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)

    '[Monica]16/05/2014: los listados de recibos pendientes de cobro y recibos por riego a manta solo lo ve Escalona
    MDIppal.mnRec_Pozos(13).Enabled = (vParamAplic.Cooperativa = 10)
    MDIppal.mnRec_Pozos(13).visible = (vParamAplic.Cooperativa = 10)
    MDIppal.mnRec_Pozos(28).Enabled = (vParamAplic.Cooperativa = 10)
    MDIppal.mnRec_Pozos(28).visible = (vParamAplic.Cooperativa = 10)
    MDIppal.mnRec_Pozos(29).Enabled = (vParamAplic.Cooperativa = 10)
    MDIppal.mnRec_Pozos(29).visible = (vParamAplic.Cooperativa = 10)

    '[Monica]11/01/2016: las facturas de recargo solo las ve Escalona
    MDIppal.mnRec_Pozos(22).Enabled = (vParamAplic.Cooperativa = 10)
    MDIppal.mnRec_Pozos(22).visible = (vParamAplic.Cooperativa = 10)

    '[Monica]12/01/2016: la rectificacion de lectura no la ve ni Escalona ni Utxera
    MDIppal.mnRec_Pozos(32).Enabled = (vParamAplic.Cooperativa = 7)
    MDIppal.mnRec_Pozos(32).visible = (vParamAplic.Cooperativa = 7)



    '[Monica]29/10/2012: la preclasificacion es solo para Anna/Castelduc [Monica]24/10/2013: tambien para bolbaite
    MDIppal.mnRec_Entradas(8).Enabled = (vParamAplic.Cooperativa = 5 Or vParamAplic.Cooperativa = 14)
    MDIppal.mnRec_Entradas(8).visible = (vParamAplic.Cooperativa = 5 Or vParamAplic.Cooperativa = 14)

    '[Monica]22/12/2011: la venta fruta la bloqueamos excepto para Alzira
    MDIppal.mnRec_Entradas(20).Enabled = (vParamAplic.Cooperativa = 4)
    MDIppal.mnRec_Entradas(20).visible = (vParamAplic.Cooperativa = 4)

    '[Monica]18/01/2012: bloqueamos todo lo de aportaciones que no sea el mantenimiento para Quatretonda
    MDIppal.mnE_Aport(1).Enabled = (vParamAplic.Cooperativa <> 7)
    MDIppal.mnE_Aport(1).visible = (vParamAplic.Cooperativa <> 7)
    MDIppal.mnE_Aport(4).Enabled = (vParamAplic.Cooperativa <> 7)
    MDIppal.mnE_Aport(4).visible = (vParamAplic.Cooperativa <> 7)
    MDIppal.mnE_Aport(5).Enabled = (vParamAplic.Cooperativa <> 7) And (vParamAplic.Cooperativa <> 14)
    MDIppal.mnE_Aport(5).visible = (vParamAplic.Cooperativa <> 7) And (vParamAplic.Cooperativa <> 14)
    MDIppal.mnE_Aport(6).Enabled = (vParamAplic.Cooperativa <> 7)
    MDIppal.mnE_Aport(6).visible = (vParamAplic.Cooperativa <> 7)
    MDIppal.mnE_Aport(7).Enabled = (vParamAplic.Cooperativa = 14)
    MDIppal.mnE_Aport(7).visible = (vParamAplic.Cooperativa = 14)
    MDIppal.mnE_Aport(8).Enabled = (vParamAplic.Cooperativa = 14)
    MDIppal.mnE_Aport(8).visible = (vParamAplic.Cooperativa = 14)
    MDIppal.mnE_Aport(9).Enabled = (vParamAplic.Cooperativa = 14)
    MDIppal.mnE_Aport(9).visible = (vParamAplic.Cooperativa = 14)

    '[Monica]29/04/2013: bloqueo de pago de facturas terceros + facturas socios, solo para montifrut
    MDIppal.mnRec_PagoSocios(18).Enabled = (vParamAplic.Cooperativa = 12)
    MDIppal.mnRec_PagoSocios(18).visible = (vParamAplic.Cooperativa = 12)

    MDIppal.mnRec_Entradas(21).Enabled = (vParamAplic.Cooperativa = 12)
    MDIppal.mnRec_Entradas(21).visible = (vParamAplic.Cooperativa = 12)

    '[Monica]10/07/2013: los listados de diferencias con indefa solo para utxera y escalona
    MDIppal.mnRec_DifPozos(2).Enabled = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)
    MDIppal.mnRec_DifPozos(2).visible = (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10)

    '[Monica]12/09/2013: solo natural tiene anticipos de nomina a trabajadores
    MDIppal.mnP_PreNominas(12).Enabled = vParamAplic.Cooperativa = 9
    MDIppal.mnP_PreNominas(12).visible = vParamAplic.Cooperativa = 9

    '[Monica]12/12/2013: el informe de atria solo lo ve Picassent
    MDIppal.mnRec_InforOfi(2).Enabled = (vParamAplic.Cooperativa = 2)
    MDIppal.mnRec_InforOfi(2).visible = (vParamAplic.Cooperativa = 2)
    
    '[Monica]12/12/2013: el informe de fitosanitarios solo para catadau
    MDIppal.mnRec_InforOfi(3).Enabled = (vParamAplic.Cooperativa = 0) Or (vParamAplic.Cooperativa = 4)
    MDIppal.mnRec_InforOfi(3).visible = (vParamAplic.Cooperativa = 0) Or (vParamAplic.Cooperativa = 4)
    
    
    '[Monica]04/05/2015: traspaso de ropas solo para Castelduc
    MDIppal.mnE_Util(16).Enabled = MDIppal.mnE_Util(16).visible And (vParamAplic.Cooperativa = 5)
    MDIppal.mnE_Util(16).visible = MDIppal.mnE_Util(16).visible And (vParamAplic.Cooperativa = 5)

    
    '[Monica]29/06/2015: traspaso de albaranes de retirada solo para ABN
    MDIppal.mnE_Util(17).Enabled = MDIppal.mnE_Util(17).visible And (vParamAplic.Cooperativa = 1)
    MDIppal.mnE_Util(17).visible = MDIppal.mnE_Util(17).visible And (vParamAplic.Cooperativa = 1)
    
    '[Monica]20/10/2015: solo para ABN traspaso de campos de almazara
    MDIppal.mnRec_AlmTrasCampos(1).Enabled = (vParamAplic.Cooperativa = 1)
    MDIppal.mnRec_AlmTrasCampos(1).visible = (vParamAplic.Cooperativa = 1)
    
    
End Sub

Public Sub BloqueoMenusSegunNivelUsuario()

    MDIppal.mnE_Util(19).visible = (vUsu.Login = "root")
    MDIppal.mnE_Util(19).Enabled = (vUsu.Login = "root")
    
    
End Sub


Public Sub BloqueoMenusSegunCampanya()
Dim b As Boolean

    ' actualizacion de datos de socios de suministros :
    '   si es campaña actual
    '   y si hay suministros
    MDIppal.mnE_Util(5).Enabled = MDIppal.mnE_Util(5).visible And (EsCampanyaActual(vEmpresa.BDAriagro) And vParamAplic.BDAriges <> "")
    MDIppal.mnE_Util(5).visible = MDIppal.mnE_Util(5).visible And (EsCampanyaActual(vEmpresa.BDAriagro) And vParamAplic.BDAriges <> "")
     
End Sub



Private Sub AbrirFormularioGlobalGAP()
    
    Set frmBas = New frmBasico
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|800|;S|txtAux(1)|T|Descripción|3930|;"
    frmBas.CadenaConsulta = "SELECT rglobalgap.codigo, rglobalgap.descripcion "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM rglobalgap "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    frmBas.Tag1 = "Código|T|N|||rglobalgap|codigo||S|"
    frmBas.Tag2 = "Descripción|T|N|||rglobalgap|descripcion|||"
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 40
    frmBas.Tabla = "rglobalgap"
    frmBas.CampoCP = "codigo"
    frmBas.Report = "rManGlobalGap.rpt"
    frmBas.Caption = "GlobalGap"
    frmBas.Show vbModal
    
    Set frmBas = Nothing

End Sub


