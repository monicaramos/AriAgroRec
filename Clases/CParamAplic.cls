VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParamAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

'Variables locales que contienen valores de propiedad
'Variables asociadas a cada campo de la tabla de parametros SPARAM

'Datos de la Contabilidad
Private mvarSerConta As String 'Servidor BD de Contabilidad
Private mvarUsuConta As String 'Usuario BD de Contabilidad
Private mvarPasConta As String 'Password BD de Contabilidad
Private mvarNumConta As Byte 'Numero BD de Contabilidad

Private mvarWebSoporte As String  'Web Soporte
Private mvarDireMail As String 'direccion mail
Private mvarSmtphost As String 'Host
Private mvarSmtpUser As String 'User
Private mvarSmtpPass As String 'Password

Private mvarEnvioDesdeOutlook As Boolean  'Si envia por outlook
Private mvarExeEnvioMail As String

Private mvarSMSemail As String 'SMS email
Private mvarSMSclave As String 'SMS clave
Private mvarSMSremitente As String 'SMS remitente


' entradas de almacen
Private mvarTipoCaja1 As String ' Tipo Caja 1
Private mvarTipoCaja2 As String ' Tipo Caja 2
Private mvarTipoCaja3 As String ' Tipo Caja 3
Private mvarTipoCaja4 As String ' Tipo Caja 4
Private mvarTipoCaja5 As String ' Tipo Caja 5

Private mvarPesoCaja1 As Currency ' Peso Caja 1
Private mvarPesoCaja2 As Currency ' Peso Caja 2
Private mvarPesoCaja3 As Currency ' Peso Caja 3
Private mvarPesoCaja4 As Currency ' Peso Caja 4
Private mvarPesoCaja5 As Currency ' Peso Caja 5

Private mvarPesoCaja11 As Currency ' Peso Caja 1 transportista
Private mvarPesoCaja12 As Currency ' Peso Caja 2 transportista
Private mvarPesoCaja13 As Currency ' Peso Caja 3 transportista
Private mvarPesoCaja14 As Currency ' Peso Caja 4 transportista
Private mvarPesoCaja15 As Currency ' Peso Caja 5 transportista

Private mvarEsCaja1 As Boolean ' es Caja 1
Private mvarEsCaja2 As Boolean ' es Caja 2
Private mvarEsCaja3 As Boolean ' es Caja 3
Private mvarEsCaja4 As Boolean ' es Caja 4
Private mvarEsCaja5 As Boolean ' es Caja 5
Private mvarKilosCajaMin As Currency ' Kilos Caja minimo
Private mvarKilosCajaMax As Currency ' Kilos Caja maximo
Private mvarPesoCajaLlena As Currency ' Peso Caja LLena

Private mvarEsVtaFruta1 As Boolean ' es VtaFruta 1
Private mvarEsVtaFruta2 As Boolean ' es VtaFruta 2
Private mvarEsVtaFruta3 As Boolean ' es VtaFruta 3
Private mvarEsVtaFruta4 As Boolean ' es VtaFruta 4
Private mvarEsVtaFruta5 As Boolean ' es VtaFruta 5


Private mvarPorcIncreAforo As String ' porcentaje de incremento/decremento de aforo

Private mvarSeTaraTractor As Boolean  'Si se tara el tractor en entrada si o no
Private mvarHayTraza As Boolean   'Si hay Trazabilidad o no
Private mvarCajasporPalet As Long ' Nro de Cajas por Palet
Private mvarSeAgrupanNotas As Boolean  'Si se agrupan notas en albaran
Private mvarSeRespetaNota As Boolean  'Si se respeta el nro de nota como nro de albaran
Private mvarNroNotaManual As Boolean  'Si el nro de nota es manual o no en entradas de bascula
Private mvarCooproenEntradas As Boolean  '0 = Si la entrada se desdobla en la actualizacion de entradas clasificadas
                                         '1 = Si la entrada se desdobla en la actualizacion de entradas bascula

Private mvarImpresoraEntradas As String  'impresora de entradas
Private mvarCodIncidPlaga As Integer ' codigo de incidencia de plaga

' aridoc
Private mvarHayAridoc As Byte  'Hay Aridoc
Private mvarCarpetaAnt As Integer 'Carpeta Anticipo
Private mvarCarpetaLiq As Integer 'Carpeta Liquidacion
Private mvarCarpetaADV As Integer 'Carpeta ADV
Private mvarCarpetaRecAlmacen As Integer 'Carpeta Recibos de Almacen
Private mvarCarpetaRecCampo As Integer 'Carpeta Recibos de Campo
Private mvarCarpetaAlmz As Integer 'Carpeta Almazara
Private mvarCarpetaBOD As Integer 'Carpeta Bodega
Private mvarCarpetaTra As Integer 'Carpeta Facturas de transporte
Private mvarExtension As Byte  'Extension del documento

Private mvarC1Anticipo As Byte  'Campo 1 Anticipo
Private mvarC2Anticipo As Byte  'Campo 2 Anticipo
Private mvarC3Anticipo As Byte  'Campo 3 Anticipo
Private mvarC4Anticipo As Byte  'Campo 4 Anticipo

Private mvarC1Liquidacion As Byte  'Campo 1 Liquidacion
Private mvarC2Liquidacion As Byte  'Campo 2 Liquidacion
Private mvarC3Liquidacion As Byte  'Campo 3 Liquidacion
Private mvarC4Liquidacion As Byte  'Campo 4 Liquidacion

Private mvarC1ADV As Byte  'Campo 1 ADV
Private mvarC2ADV As Byte  'Campo 2 ADV
Private mvarC3ADV As Byte  'Campo 3 ADV
Private mvarC4ADV As Byte  'Campo 4 ADV

Private mvarC1Recibo As Byte  'Campo 1 Recibo
Private mvarC2Recibo As Byte  'Campo 2 Recibo
Private mvarC3Recibo As Byte  'Campo 3 Recibo
Private mvarC4Recibo As Byte  'Campo 4 Recibo

Private mvarC1Almazara As Byte  'Campo 1 Almazara
Private mvarC2Almazara As Byte  'Campo 2 Almazara
Private mvarC3Almazara As Byte  'Campo 3 Almazara
Private mvarC4Almazara As Byte  'Campo 4 Almazara

Private mvarC1Bodega As Byte  'Campo 1 Bodega
Private mvarC2Bodega As Byte  'Campo 2 Bodega
Private mvarC3Bodega As Byte  'Campo 3 Bodega
Private mvarC4Bodega As Byte  'Campo 4 Bodega


Private mvarC1Transporte As Byte  'Campo 1 Transporte
Private mvarC2Transporte As Byte  'Campo 2 Transporte
Private mvarC3Transporte As Byte  'Campo 3 Transporte
Private mvarC4Transporte As Byte  'Campo 4 Transporte


Private mvarFaneca As String ' constante faneca

Private mvarCosteHora As Currency ' coste hora
Private mvarCosteSegSo As Currency ' coste seguridad social
Private mvarSeccionHorto As String ' seccion de hortofruticola
Private mvarSeccionAlmaz As String ' seccion de almazara
Private mvarSeccionADV As String ' seccion de adv
Private mvarSeccionSumi As String ' seccion de suministros
Private mvarSeccionBodega As String ' seccion de bodega
Private mvarSeccionPozos As String ' seccion de pozos


Private mvarPorcReteFacSoc As String ' porcentaje de retencion

Private mvarPrimFactAnt As String ' primera factura de anticipo
Private mvarUltFactAnt As String ' ultima factura de anticipo

Private mvarPrimFactAntVC As String ' primera factura de anticipo venta campo
Private mvarUltFactAntVC As String ' ultima factura de anticipo venta campo

Private mvarPrimFactLiq As String ' primera factura de Liquidación
Private mvarUltFactLiq As String ' ultima factura de Liquidación

Private mvarPrimFactLiqVC As String ' primera factura de liquidación venta campo
Private mvarUltFactLiqVC As String ' ultima factura de liquidación venta campo

Private mvarPorcenAFO As String ' porcentaje de aportacion fondo operativo

Private mvarTTomaDatos As String ' Texto para la toma de datos
Private mvarPieTomaDatos As String ' Texto pie para la toma de datos

Private mvarCooperativa As Byte ' 0 = Catadau
                                ' 1 = Valsur
                                ' 2 = Picassent
                                ' 3 = Mogente
                                ' 4 = Alzira
                                ' 5 = Castelduc
                                ' 6 = Macastre
                                ' 7 = Quatretonda
                                ' 8 = utxera
                                ' 9 = Natural
                                '10 = Escalona
                                '11 = Belgida
                                '12 = Montifrut
                                '14 = Bolbaite

Private mvarCodIvaIntra As String ' codigo de iva aplicado a intracomunitarios
Private mvarCtaTerReten As String ' cuenta contable de retencion de facturas de terceros

Private mvarCtaTraReten As String ' cuenta contable de retencion de facturas de transporte


'anticipos / liquidaciones
Private mvarCtaRetenSoc As String ' cta de retencion de socios
Private mvarCtaAportaSoc As String ' cta de aportacion de socios
Private mvarCtaBancoSoc As String ' cta de banco contabilizacion facturas de socios
Private mvarForpaPosi As Integer ' forma de pago de anti/liqui positivos
Private mvarForpaNega As Integer ' forma de pago de anti/liqui negativos


'Almazara
Private mvarCtaRetenAlmz As String ' cta de retencion de almazara
Private mvarCtaBancoAlmz As String ' cta de banco contabilizacion almazara
Private mvarForpaPosiAlmz As Integer ' forma de pago de almazara positivos
Private mvarForpaNegaAlmz As Integer ' forma de pago de almazara negativos
Private mvarCtaVentasAlmz As String ' cta ventas de almazara
Private mvarCtaGastosAlmz As String ' cta gastos de almazara
Private mvarLetraSerieAlmz As String ' letra de serie de almazara

Private mvarCodGastoAlmz As Integer ' codigo de concepto gasto para reparto de gasto de liquidacion almazara

Private mvarPrimFactAntAlmz As String ' primera factura de anticipo almazara
Private mvarUltFactAntAlmz As String ' ultima factura de anticipo almazara

Private mvarPrimFactLiqAlmz As String ' primera factura de Liquidación almazara
Private mvarUltFactLiqAlmz As String ' ultima factura de Liquidación almazara

Private mvarGtoMoltura As String ' precio por gastos de molturacion
Private mvarGtoEnvasado As String ' precio por gastos de envasado

'ADV
Private mvarAlmacenADV As Integer ' codigo de almacen de adv
Private mvarCtaBancoADV As String ' cta de banco contabilizacion adv
Private mvarCodIvaExeADV As String ' codigo de iva facturas internas de adv


Private mvarPathTraza As String     ' path de trazabilidad

Private mvarSigpac As String     ' direccion de sigpac
Private mvarGoolZoom As String     ' direccion de goolzoom


'Suministros
Private mvarBDAriges As String ' base de datos de ariges


'Bodega
Private mvarAlbRetiradaManual As Boolean  'Nro Albaran de retirada manual
Private mvarCtaBancoBOD As String ' cta de banco prevista de bodega
Private mvarCtaVentasBOD As String ' cta de ventas de bodega
Private mvarPorcGtoMantBOD As String ' porcentaje de gto de mantenimiento en liquidacion bodega
Private mvarCodGastoBOD As Integer ' codigo de concepto gasto para reparto de gasto de liquidacion bodega

Private mvarPrimFactAntBOD As String ' primera factura de anticipo bodega
Private mvarUltFactAntBOD As String ' ultima factura de anticipo bodega

Private mvarPrimFactLiqBOD As String ' primera factura de Liquidación bodega
Private mvarUltFactLiqBOD As String ' ultima factura de Liquidación bodega

Private mvarPorcKilosBOD As String ' porcentaje de incremento de kilos entrados de bodega (Para Mogente)

'Telefonia
Private mvarLetraSerieTel As String ' letra de serie de telefonia
Private mvarCtaVentasTel As String ' cta de ventas de telefonia

'Nomina
Private mvarAlmacenNOMI As Integer ' codigo de almacen de gestion de nominas
Private mvarEurosTrabdiaNOMI As Currency ' euros trabajador dia

Private mvarPorcIRPFNOMI As String ' porcentaje irpf
Private mvarPorcSegSo1NOMI As String ' porcentaje seguridad social 1
Private mvarPorcSegSo2NOMI As String ' porcentaje seguridad social 2
Private mvarPorcJornadaNOMI As String ' porcentaje jornada nominas
Private mvarMaxJornadasNOMI As String  ' nro maximo de jornadas nominas

'Pozos
Private mvarCodIvaPOZ As String ' codigo de iva aplicado a pozos

Private mvarPrecio1POZ As String ' precio rango 1
Private mvarPrecio2POZ As String ' precio rango 2

Private mvarConsumo1POZ As String ' hasta m3 en rango 1
Private mvarConsumo2POZ As String ' hasta m3 en rango 2

Private mvarCuotaPOZ As String ' cuota pozos
Private mvarDerramaPOZ As String ' derrama pozos
Private mvarCtaVentasConsPOZ As String ' cta ventas consumo de pozos
Private mvarCtaVentasCuoPOZ As String ' cta ventas cuotas de pozos
Private mvarCtaVentasTalPOZ As String ' cta ventas talla de pozos
Private mvarCtaVentasMtoPOZ As String ' cta ventas mantenimiento de pozos
'[Monica]21/01/2016: cuenta de recargos (escalona)
Private mvarCtaVentasMantaPOZ As String ' cta ventas consumo a manta pozos

Private mvarCtaRecargosPOZ As String ' cta recargos pozos



Private mvarCodCCostPoz As String ' centro de coste de pozos
 
Private mvarForpaConPOZ As String ' forma de pago contado pozos
Private mvarForpaRecPOZ As String ' forma de pago recibo pozos
 
Private mvarCartaPOZ As Integer ' codigo de carta de reclamacion de pozos (solo utxera y escalona)
Private mvarNomZonaPOZ As String ' nombre de la zona de pozos (braçal)
 
 
 
'Transporte
Private mvarTipoPortesTRA As Byte ' 0 = Portes por Poblacion 1 = Portes por Tarifa
Private mvarCodGastoTRA As Integer ' codigo de concepto gasto para transporte
Private mvarCodTarifTRA As Integer ' codigo de tarifa de transporte de tipo 1
Private mvarCodTarifTRA2 As Integer ' codigo de tarifa de transporte de tipo 2

Private mvarPorcReteFacTra As String ' porcentaje de retencion para modulo transportistas
Private mvarPreciokgTra As String ' precio por kilo de transporte
Private mvarTipoContadorTRA As Byte ' 0 = Contador Global  1 = Contador por Transportista


'Variable de sparam de control de costes
Private mvarHayCCostes As Byte  'Hay Control de costes


'Tipo de Lectura que leemos de pozos (Escalona)
Private mvarTipoLecturaPoz As Byte  'tipo de lectura de pozos 0=lectura_equipo 1=lectura_bd


Private mvarLiqDirecta As Byte 'liquidacion directa

'Variable de rparam de facturas varias
Private mvarHayFVarias As Byte  'Hay Facturas Varias


Private mvarPathFacturaE As String  'FacturaE


Private Sql As String  'Para la sentencia
Private Mens As String    'Para los mensajes

Private mvarConsumoMinPoz As Long ' Consumo minimo de pozos
Private mvarConsumoMaxPoz As Long ' Consumo maximo de pozos


Private mvarClaveAcceso As String ' Clave de acceso

'[Monica]21/04/2016: servidor sqlserver de Castellduc
Private mvarSqlServer As String ' Servidor de SqlServer de Castellduc CASTELLDUCS1-O5
Private mvarSqlServer1 As String ' Servidor de SqlServer de Castellduc CASTELLDUCS2-O5

'[Monica]10/10/2016: hay resumen por variedad de cajas, kilos e importes en el mto de partes de trabajo
Private mvarHayResumenCajas As Byte

Private mvarPathEntradas As String     ' path de impresion de entradas



'------------------------------------------------
'Propiedades del modulo CParamAplic
'------------------------------------------------

'servidor conta
Public Property Let ServidorConta(ByVal vData As String)
     mvarSerConta = vData
End Property

Public Property Get ServidorConta() As String
     ServidorConta = mvarSerConta
End Property

Public Property Let UsuarioConta(ByVal vData As String)
     mvarUsuConta = vData
End Property

Public Property Get UsuarioConta() As String
     UsuarioConta = mvarUsuConta
End Property

Public Property Let PasswordConta(ByVal vData As String)
     mvarPasConta = vData
End Property

Public Property Get PasswordConta() As String
     PasswordConta = mvarPasConta
End Property

Public Property Let NumeroConta(ByVal vData As Byte)
     mvarNumConta = vData
End Property

Public Property Get NumeroConta() As Byte
     NumeroConta = mvarNumConta
End Property

Public Property Let WebSoporte(ByVal vData As String)
     mvarWebSoporte = vData
End Property

Public Property Get WebSoporte() As String
     WebSoporte = mvarWebSoporte
End Property

Public Property Let DireMail(ByVal vData As String)
     mvarDireMail = vData
End Property

Public Property Get DireMail() As String
     DireMail = mvarDireMail
End Property

Public Property Let Smtphost(ByVal vData As String)
     mvarSmtphost = vData
End Property

Public Property Get Smtphost() As String
     Smtphost = mvarSmtphost
End Property

Public Property Let SmtpUser(ByVal vData As String)
     mvarSmtpUser = vData
End Property

Public Property Get SmtpUser() As String
     SmtpUser = mvarSmtpUser
End Property

Public Property Let Smtppass(ByVal vData As String)
     mvarSmtpPass = vData
End Property

Public Property Get Smtppass() As String
     Smtppass = mvarSmtpPass
End Property

' Envio de email por outlook o arigesmail

Public Property Get EnvioDesdeOutlook() As Boolean
     EnvioDesdeOutlook = mvarEnvioDesdeOutlook
End Property

Public Property Let EnvioDesdeOutlook(ByVal vData As Boolean)
     mvarEnvioDesdeOutlook = vData
End Property

Public Property Get ExeEnvioMail() As String
     ExeEnvioMail = mvarExeEnvioMail
End Property

Public Property Let ExeEnvioMail(ByVal vData As String)
     mvarExeEnvioMail = vData
End Property


'########## SMS

Public Property Let SMSemail(ByVal vData As String)
     mvarSMSemail = vData
End Property

Public Property Get SMSemail() As String
     SMSemail = mvarSMSemail
End Property

Public Property Let SMSclave(ByVal vData As String)
     mvarSMSclave = vData
End Property

Public Property Get SMSclave() As String
     SMSclave = mvarSMSclave
End Property

Public Property Let SMSremitente(ByVal vData As String)
     mvarSMSremitente = vData
End Property

Public Property Get SMSremitente() As String
     SMSremitente = mvarSMSremitente
End Property

' Entradas de almacen
Public Property Let SeTaraTractor(ByVal vData As Boolean)
     mvarSeTaraTractor = vData
End Property

Public Property Get SeTaraTractor() As Boolean
     SeTaraTractor = mvarSeTaraTractor
End Property

Public Property Let SeAgrupanNotas(ByVal vData As Boolean)
     mvarSeAgrupanNotas = vData
End Property

Public Property Get SeAgrupanNotas() As Boolean
     SeAgrupanNotas = mvarSeAgrupanNotas
End Property

Public Property Let SeRespetaNota(ByVal vData As Boolean)
     mvarSeRespetaNota = vData
End Property

Public Property Get SeRespetaNota() As Boolean
     SeRespetaNota = mvarSeRespetaNota
End Property

Public Property Let HayTraza(ByVal vData As Boolean)
     mvarHayTraza = vData
End Property

Public Property Get HayTraza() As Boolean
     HayTraza = mvarHayTraza
End Property


Public Property Let NroNotaManual(ByVal vData As Boolean)
     mvarNroNotaManual = vData
End Property

Public Property Get NroNotaManual() As Boolean
     NroNotaManual = mvarNroNotaManual
End Property


Public Property Let CooproenEntradas(ByVal vData As Boolean)
     mvarCooproenEntradas = vData
End Property

Public Property Get CooproenEntradas() As Boolean
     CooproenEntradas = mvarCooproenEntradas
End Property


Public Property Let CajasporPalet(ByVal vData As Long)
     mvarCajasporPalet = vData
End Property

Public Property Get CajasporPalet() As Long
     CajasporPalet = mvarCajasporPalet
End Property

Public Property Let TipoCaja1(ByVal vData As String)
     mvarTipoCaja1 = vData
End Property

Public Property Get TipoCaja1() As String
     TipoCaja1 = mvarTipoCaja1
End Property

Public Property Let TipoCaja2(ByVal vData As String)
     mvarTipoCaja2 = vData
End Property

Public Property Get TipoCaja2() As String
     TipoCaja2 = mvarTipoCaja2
End Property

Public Property Let TipoCaja3(ByVal vData As String)
     mvarTipoCaja3 = vData
End Property

Public Property Get TipoCaja3() As String
     TipoCaja3 = mvarTipoCaja3
End Property

Public Property Let TipoCaja4(ByVal vData As String)
     mvarTipoCaja4 = vData
End Property

Public Property Get TipoCaja4() As String
     TipoCaja4 = mvarTipoCaja4
End Property

Public Property Let TipoCaja5(ByVal vData As String)
     mvarTipoCaja5 = vData
End Property

Public Property Get TipoCaja5() As String
     TipoCaja5 = mvarTipoCaja5
End Property

Public Property Let PesoCaja1(ByVal vData As Currency)
     mvarPesoCaja1 = vData
End Property

Public Property Get PesoCaja1() As Currency
     PesoCaja1 = mvarPesoCaja1
End Property

Public Property Let PesoCaja2(ByVal vData As Currency)
     mvarPesoCaja2 = vData
End Property

Public Property Get PesoCaja2() As Currency
     PesoCaja2 = mvarPesoCaja2
End Property

Public Property Let PesoCaja3(ByVal vData As Currency)
     mvarPesoCaja3 = vData
End Property

Public Property Get PesoCaja3() As Currency
     PesoCaja3 = mvarPesoCaja3
End Property

Public Property Let PesoCaja4(ByVal vData As Currency)
     mvarPesoCaja4 = vData
End Property

Public Property Get PesoCaja4() As Currency
     PesoCaja4 = mvarPesoCaja4
End Property

Public Property Let PesoCaja5(ByVal vData As Currency)
     mvarPesoCaja5 = vData
End Property

Public Property Get PesoCaja5() As Currency
     PesoCaja5 = mvarPesoCaja5
End Property


Public Property Let PesoCaja11(ByVal vData As Currency)
     mvarPesoCaja11 = vData
End Property

Public Property Get PesoCaja11() As Currency
     PesoCaja11 = mvarPesoCaja11
End Property

Public Property Let PesoCaja12(ByVal vData As Currency)
     mvarPesoCaja12 = vData
End Property

Public Property Get PesoCaja12() As Currency
     PesoCaja12 = mvarPesoCaja12
End Property

Public Property Let PesoCaja13(ByVal vData As Currency)
     mvarPesoCaja13 = vData
End Property

Public Property Get PesoCaja13() As Currency
     PesoCaja13 = mvarPesoCaja13
End Property

Public Property Let PesoCaja14(ByVal vData As Currency)
     mvarPesoCaja14 = vData
End Property

Public Property Get PesoCaja14() As Currency
     PesoCaja14 = mvarPesoCaja14
End Property

Public Property Let PesoCaja15(ByVal vData As Currency)
     mvarPesoCaja15 = vData
End Property

Public Property Get PesoCaja15() As Currency
     PesoCaja15 = mvarPesoCaja15
End Property


' Entradas de almacen
Public Property Let EsCaja1(ByVal vData As Boolean)
     mvarEsCaja1 = vData
End Property

Public Property Get EsCaja1() As Boolean
     EsCaja1 = mvarEsCaja1
End Property

Public Property Let EsCaja2(ByVal vData As Boolean)
     mvarEsCaja2 = vData
End Property

Public Property Get EsCaja2() As Boolean
     EsCaja2 = mvarEsCaja2
End Property

Public Property Let EsCaja3(ByVal vData As Boolean)
     mvarEsCaja3 = vData
End Property

Public Property Get EsCaja3() As Boolean
     EsCaja3 = mvarEsCaja3
End Property

Public Property Let EsCaja4(ByVal vData As Boolean)
     mvarEsCaja4 = vData
End Property

Public Property Get EsCaja4() As Boolean
     EsCaja4 = mvarEsCaja4
End Property

Public Property Let EsCaja5(ByVal vData As Boolean)
     mvarEsCaja5 = vData
End Property

Public Property Get EsCaja5() As Boolean
     EsCaja5 = mvarEsCaja5
End Property


Public Property Let EsVtaFruta1(ByVal vData As Boolean)
     mvarEsVtaFruta1 = vData
End Property

Public Property Get EsVtaFruta1() As Boolean
     EsVtaFruta1 = mvarEsVtaFruta1
End Property

Public Property Let EsVtaFruta2(ByVal vData As Boolean)
     mvarEsVtaFruta2 = vData
End Property

Public Property Get EsVtaFruta2() As Boolean
     EsVtaFruta2 = mvarEsVtaFruta2
End Property

Public Property Let EsVtaFruta3(ByVal vData As Boolean)
     mvarEsVtaFruta3 = vData
End Property

Public Property Get EsVtaFruta3() As Boolean
     EsVtaFruta3 = mvarEsVtaFruta3
End Property

Public Property Let EsVtaFruta4(ByVal vData As Boolean)
     mvarEsVtaFruta4 = vData
End Property

Public Property Get EsVtaFruta4() As Boolean
     EsVtaFruta4 = mvarEsVtaFruta4
End Property

Public Property Let EsVtaFruta5(ByVal vData As Boolean)
     mvarEsVtaFruta5 = vData
End Property

Public Property Get EsVtaFruta5() As Boolean
     EsVtaFruta5 = mvarEsVtaFruta5
End Property


Public Property Let KilosCajaMin(ByVal vData As Currency)
     mvarKilosCajaMin = vData
End Property

Public Property Get KilosCajaMin() As Currency
     KilosCajaMin = mvarKilosCajaMin
End Property


Public Property Let KilosCajaMax(ByVal vData As Currency)
     mvarKilosCajaMax = vData
End Property

Public Property Get KilosCajaMax() As Currency
     KilosCajaMax = mvarKilosCajaMax
End Property

Public Property Let PesoCajaLLena(ByVal vData As Currency)
     mvarPesoCajaLlena = vData
End Property

Public Property Get PesoCajaLLena() As Currency
     PesoCajaLLena = mvarPesoCajaLlena
End Property


Public Property Let ImpresoraEntradas(ByVal vData As String)
     mvarImpresoraEntradas = vData
End Property

Public Property Get ImpresoraEntradas() As String
     ImpresoraEntradas = mvarImpresoraEntradas
End Property


Public Property Let CodIncidPlaga(ByVal vData As Integer)
     mvarCodIncidPlaga = vData
End Property

Public Property Get CodIncidPlaga() As Integer
     CodIncidPlaga = mvarCodIncidPlaga
End Property






Public Property Let PorcIncreAforo(ByVal vData As String)
     mvarPorcIncreAforo = vData
End Property

Public Property Get PorcIncreAforo() As String
     PorcIncreAforo = mvarPorcIncreAforo
End Property

' Aridoc
Public Property Let HayAridoc(ByVal vData As Byte)
     mvarHayAridoc = vData
End Property

Public Property Get HayAridoc() As Byte
     HayAridoc = mvarHayAridoc
End Property

Public Property Let CarpetaAnt(ByVal vData As Integer)
     mvarCarpetaAnt = vData
End Property

Public Property Get CarpetaAnt() As Integer
     CarpetaAnt = mvarCarpetaAnt
End Property

Public Property Let CarpetaLiq(ByVal vData As Integer)
     mvarCarpetaLiq = vData
End Property

Public Property Get CarpetaLiq() As Integer
     CarpetaLiq = mvarCarpetaLiq
End Property


Public Property Let CarpetaADV(ByVal vData As Integer)
     mvarCarpetaADV = vData
End Property

Public Property Get CarpetaADV() As Integer
     CarpetaADV = mvarCarpetaADV
End Property

Public Property Let CarpetaRecAlmacen(ByVal vData As Integer)
     mvarCarpetaRecAlmacen = vData
End Property

Public Property Get CarpetaRecAlmacen() As Integer
     CarpetaRecAlmacen = mvarCarpetaRecAlmacen
End Property

Public Property Let CarpetaRecCampo(ByVal vData As Integer)
     mvarCarpetaRecCampo = vData
End Property

Public Property Get CarpetaRecCampo() As Integer
     CarpetaRecCampo = mvarCarpetaRecCampo
End Property

Public Property Let CarpetaAlmz(ByVal vData As Integer)
     mvarCarpetaAlmz = vData
End Property

Public Property Get CarpetaAlmz() As Integer
     CarpetaAlmz = mvarCarpetaAlmz
End Property

Public Property Let CarpetaBOD(ByVal vData As Integer)
     mvarCarpetaBOD = vData
End Property

Public Property Get CarpetaBOD() As Integer
     CarpetaBOD = mvarCarpetaBOD
End Property


Public Property Let CarpetaTra(ByVal vData As Integer)
     mvarCarpetaTra = vData
End Property

Public Property Get CarpetaTra() As Integer
     CarpetaTra = mvarCarpetaTra
End Property



Public Property Let Extension(ByVal vData As Byte)
     mvarExtension = vData
End Property

Public Property Get Extension() As Byte
     Extension = mvarExtension
End Property


Public Property Let C1Anticipo(ByVal vData As Byte)
     mvarC1Anticipo = vData
End Property

Public Property Get C1Anticipo() As Byte
     C1Anticipo = mvarC1Anticipo
End Property

Public Property Let C2Anticipo(ByVal vData As Byte)
     mvarC2Anticipo = vData
End Property

Public Property Get C2Anticipo() As Byte
     C2Anticipo = mvarC2Anticipo
End Property

Public Property Let C3Anticipo(ByVal vData As Byte)
     mvarC3Anticipo = vData
End Property

Public Property Get C3Anticipo() As Byte
     C3Anticipo = mvarC3Anticipo
End Property

Public Property Let C4Anticipo(ByVal vData As Byte)
     mvarC4Anticipo = vData
End Property

Public Property Get C4Anticipo() As Byte
     C4Anticipo = mvarC4Anticipo
End Property


Public Property Let C1Liquidacion(ByVal vData As Byte)
     mvarC1Liquidacion = vData
End Property

Public Property Get C1Liquidacion() As Byte
     C1Liquidacion = mvarC1Liquidacion
End Property

Public Property Let C2Liquidacion(ByVal vData As Byte)
     mvarC2Liquidacion = vData
End Property

Public Property Get C2Liquidacion() As Byte
     C2Liquidacion = mvarC2Liquidacion
End Property

Public Property Let C3Liquidacion(ByVal vData As Byte)
     mvarC3Liquidacion = vData
End Property

Public Property Get C3Liquidacion() As Byte
     C3Liquidacion = mvarC3Liquidacion
End Property

Public Property Let C4Liquidacion(ByVal vData As Byte)
     mvarC4Liquidacion = vData
End Property

Public Property Get C4Liquidacion() As Byte
     C4Liquidacion = mvarC4Liquidacion
End Property


Public Property Let C1ADV(ByVal vData As Byte)
     mvarC1ADV = vData
End Property

Public Property Get C1ADV() As Byte
     C1ADV = mvarC1ADV
End Property

Public Property Let C2ADV(ByVal vData As Byte)
     mvarC2ADV = vData
End Property

Public Property Get C2ADV() As Byte
     C2ADV = mvarC2ADV
End Property

Public Property Let C3ADV(ByVal vData As Byte)
     mvarC3ADV = vData
End Property

Public Property Get C3ADV() As Byte
     C3ADV = mvarC3ADV
End Property

Public Property Let C4ADV(ByVal vData As Byte)
     mvarC4ADV = vData
End Property

Public Property Get C4ADV() As Byte
     C4ADV = mvarC4ADV
End Property

Public Property Let C1Recibo(ByVal vData As Byte)
     mvarC1Recibo = vData
End Property

Public Property Get C1Recibo() As Byte
     C1Recibo = mvarC1Recibo
End Property

Public Property Let C2Recibo(ByVal vData As Byte)
     mvarC2Recibo = vData
End Property

Public Property Get C2Recibo() As Byte
     C2Recibo = mvarC2Recibo
End Property

Public Property Let C3Recibo(ByVal vData As Byte)
     mvarC3Recibo = vData
End Property

Public Property Get C3Recibo() As Byte
     C3Recibo = mvarC3Recibo
End Property

Public Property Let C4Recibo(ByVal vData As Byte)
     mvarC4Recibo = vData
End Property

Public Property Get C4Recibo() As Byte
     C4Recibo = mvarC4Recibo
End Property

Public Property Let C1Almz(ByVal vData As Byte)
     mvarC1Almazara = vData
End Property

Public Property Get C1Almz() As Byte
     C1Almz = mvarC1Almazara
End Property

Public Property Let C2Almz(ByVal vData As Byte)
     mvarC2Almazara = vData
End Property

Public Property Get C2Almz() As Byte
     C2Almz = mvarC2Almazara
End Property

Public Property Let C3Almz(ByVal vData As Byte)
     mvarC3Almazara = vData
End Property

Public Property Get C3Almz() As Byte
     C3Almz = mvarC3Almazara
End Property

Public Property Let C4Almz(ByVal vData As Byte)
     mvarC4Almazara = vData
End Property

Public Property Get C4Almz() As Byte
     C4Almz = mvarC4Almazara
End Property


Public Property Let C1BOD(ByVal vData As Byte)
     mvarC1Bodega = vData
End Property

Public Property Get C1BOD() As Byte
     C1BOD = mvarC1Bodega
End Property

Public Property Let C2BOD(ByVal vData As Byte)
     mvarC2Bodega = vData
End Property

Public Property Get C2BOD() As Byte
     C2BOD = mvarC2Bodega
End Property

Public Property Let C3BOD(ByVal vData As Byte)
     mvarC3Bodega = vData
End Property

Public Property Get C3BOD() As Byte
     C3BOD = mvarC3Bodega
End Property

Public Property Let C4BOD(ByVal vData As Byte)
     mvarC4Bodega = vData
End Property

Public Property Get C4BOD() As Byte
     C4BOD = mvarC4Bodega
End Property

Public Property Let C1Transporte(ByVal vData As Byte)
     mvarC1Transporte = vData
End Property

Public Property Get C1Transporte() As Byte
     C1Transporte = mvarC1Transporte
End Property

Public Property Let C2Transporte(ByVal vData As Byte)
     mvarC2Transporte = vData
End Property

Public Property Get C2Transporte() As Byte
     C2Transporte = mvarC2Transporte
End Property

Public Property Let C3Transporte(ByVal vData As Byte)
     mvarC3Transporte = vData
End Property

Public Property Get C3Transporte() As Byte
     C3Transporte = mvarC3Transporte
End Property

Public Property Let C4Transporte(ByVal vData As Byte)
     mvarC4Transporte = vData
End Property

Public Property Get C4Transporte() As Byte
     C4Transporte = mvarC4Transporte
End Property

Public Property Let CosteHora(ByVal vData As Currency)
     mvarCosteHora = vData
End Property

Public Property Get CosteHora() As Currency
     CosteHora = mvarCosteHora
End Property

Public Property Let CosteSegSo(ByVal vData As Currency)
     mvarCosteSegSo = vData
End Property

Public Property Get CosteSegSo() As Currency
     CosteSegSo = mvarCosteSegSo
End Property

Public Property Let Seccionhorto(ByVal vData As String)
     mvarSeccionHorto = vData
End Property

Public Property Get Seccionhorto() As String
     Seccionhorto = mvarSeccionHorto
End Property

Public Property Let SeccionPOZOS(ByVal vData As String)
     mvarSeccionPozos = vData
End Property

Public Property Get SeccionPOZOS() As String
     SeccionPOZOS = mvarSeccionPozos
End Property



Public Property Let SeccionAlmaz(ByVal vData As String)
     mvarSeccionAlmaz = vData
End Property

Public Property Get SeccionAlmaz() As String
     SeccionAlmaz = mvarSeccionAlmaz
End Property


Public Property Let SeccionADV(ByVal vData As String)
     mvarSeccionADV = vData
End Property

Public Property Get SeccionADV() As String
     SeccionADV = mvarSeccionADV
End Property

Public Property Let SeccionSumi(ByVal vData As String)
     mvarSeccionSumi = vData
End Property

Public Property Get SeccionSumi() As String
     SeccionSumi = mvarSeccionSumi
End Property


Public Property Let SeccionBodega(ByVal vData As String)
     mvarSeccionBodega = vData
End Property

Public Property Get SeccionBodega() As String
     SeccionBodega = mvarSeccionBodega
End Property

Public Property Let Faneca(ByVal vData As String)
     mvarFaneca = vData
End Property

Public Property Get Faneca() As String
     Faneca = mvarFaneca
End Property



Public Property Let PorcreteFacSoc(ByVal vData As String)
     mvarPorcReteFacSoc = vData
End Property

Public Property Get PorcreteFacSoc() As String
     PorcreteFacSoc = mvarPorcReteFacSoc
End Property

' FACTURACION DE ANTICIPOS
' primera factura de anticipo
Public Property Let PrimFactAnt(ByVal vData As String)
     mvarPrimFactAnt = vData
End Property

Public Property Get PrimFactAnt() As String
     PrimFactAnt = mvarPrimFactAnt
End Property

' ultima factura de anticipo
Public Property Let UltFactAnt(ByVal vData As String)
     mvarUltFactAnt = vData
End Property

Public Property Get UltFactAnt() As String
     UltFactAnt = mvarUltFactAnt
End Property

' FACTURACION DE ANTICIPOS VENTA CAMPO
' primera factura de anticipo
Public Property Let PrimFactAntVC(ByVal vData As String)
     mvarPrimFactAntVC = vData
End Property

Public Property Get PrimFactAntVC() As String
     PrimFactAntVC = mvarPrimFactAntVC
End Property

' ultima factura de anticipo
Public Property Let UltFactAntVC(ByVal vData As String)
     mvarUltFactAntVC = vData
End Property

Public Property Get UltFactAntVC() As String
     UltFactAntVC = mvarUltFactAntVC
End Property

' FACTURACION DE LIQUIDACION
' primera factura de liquidacion
Public Property Let PrimFactLiq(ByVal vData As String)
     mvarPrimFactLiq = vData
End Property

Public Property Get PrimFactLiq() As String
     PrimFactLiq = mvarPrimFactLiq
End Property

' ultima factura de liquidacion
Public Property Let UltFactLiq(ByVal vData As String)
     mvarUltFactLiq = vData
End Property

Public Property Get UltFactLiq() As String
     UltFactLiq = mvarUltFactLiq
End Property

' FACTURACION DE LIQUIDACION VENTA CAMPO
' primera factura de liquidacion venta campo
Public Property Let PrimFactLiqVC(ByVal vData As String)
     mvarPrimFactLiqVC = vData
End Property

Public Property Get PrimFactLiqVC() As String
     PrimFactLiqVC = mvarPrimFactLiqVC
End Property

' ultima factura de liquidacion venta campo
Public Property Let UltFactLiqVC(ByVal vData As String)
     mvarUltFactLiqVC = vData
End Property

Public Property Get UltFactLiqVC() As String
     UltFactLiqVC = mvarUltFactLiqVC
End Property

' Porcentaje de fondo operativo
Public Property Let PorcenAFO(ByVal vData As String)
     mvarPorcenAFO = vData
End Property

Public Property Get PorcenAFO() As String
     PorcenAFO = mvarPorcenAFO
End Property


' Texto cabecera de toma de datos
Public Property Get TTomaDatos() As String
     TTomaDatos = mvarTTomaDatos
End Property

Public Property Let TTomaDatos(ByVal vData As String)
     mvarTTomaDatos = vData
End Property

' Texto pie de toma de datos
Public Property Get PieTomaDatos() As String
     PieTomaDatos = mvarPieTomaDatos
End Property

Public Property Let PieTomaDatos(ByVal vData As String)
     mvarPieTomaDatos = vData
End Property


Public Property Let Cooperativa(ByVal vData As Byte)
     mvarCooperativa = vData
End Property

Public Property Get Cooperativa() As Byte
     Cooperativa = mvarCooperativa
End Property

' codigo de iva aplicado a intracomunitarios

Public Property Let CodIvaIntra(ByVal vData As String)
     mvarCodIvaIntra = vData
End Property

Public Property Get CodIvaIntra() As String
     CodIvaIntra = mvarCodIvaIntra
End Property

' cuenta contable de retencion de facturas de terceros

Public Property Let CtaTerReten(ByVal vData As String)
     mvarCtaTerReten = vData
End Property

Public Property Get CtaTerReten() As String
     CtaTerReten = mvarCtaTerReten
End Property


' cuenta contable de retencion de facturas de transporte

Public Property Let CtaTraReten(ByVal vData As String)
     mvarCtaTraReten = vData
End Property

Public Property Get CtaTraReten() As String
     CtaTraReten = mvarCtaTraReten
End Property




'path de Trazabilidad

Public Property Let PathTraza(ByVal vData As String)
     mvarPathTraza = vData
End Property

Public Property Get PathTraza() As String
     PathTraza = mvarPathTraza
End Property


'direccion sigpac

Public Property Let SigPac(ByVal vData As String)
     mvarSigpac = vData
End Property

Public Property Get SigPac() As String
     SigPac = mvarSigpac
End Property


'direccion goolzoom

Public Property Let GoolZoom(ByVal vData As String)
     mvarGoolZoom = vData
End Property

Public Property Get GoolZoom() As String
     GoolZoom = mvarGoolZoom
End Property




' cuenta contable de retencion de facturas de socios

Public Property Let CtaRetenSoc(ByVal vData As String)
     mvarCtaRetenSoc = vData
End Property

Public Property Get CtaRetenSoc() As String
     CtaRetenSoc = mvarCtaRetenSoc
End Property

' cuenta contable de aportacion de facturas de socios

Public Property Let CtaAportaSoc(ByVal vData As String)
     mvarCtaAportaSoc = vData
End Property

Public Property Get CtaAportaSoc() As String
     CtaAportaSoc = mvarCtaAportaSoc
End Property


' cuenta contable de banco de facturas de socios

Public Property Let CtaBancoSoc(ByVal vData As String)
     mvarCtaBancoSoc = vData
End Property

Public Property Get CtaBancoSoc() As String
     CtaBancoSoc = mvarCtaBancoSoc
End Property


' forma de pago de facturas positivas

Public Property Let ForpaPosi(ByVal vData As Integer)
     mvarForpaPosi = vData
End Property

Public Property Get ForpaPosi() As Integer
     ForpaPosi = mvarForpaPosi
End Property

' forma de pago de facturas negativas

Public Property Let ForpaNega(ByVal vData As Integer)
     mvarForpaNega = vData
End Property

Public Property Get ForpaNega() As Integer
     ForpaNega = mvarForpaNega
End Property

'*********************
'      ALMAZARA
'*********************

' cuenta contable de retencion
Public Property Let CtaRetenAlmz(ByVal vData As String)
     mvarCtaRetenAlmz = vData
End Property

Public Property Get CtaRetenAlmz() As String
     CtaRetenAlmz = mvarCtaRetenAlmz
End Property


' cuenta contable de banco de almazara

Public Property Let CtaBancoAlmz(ByVal vData As String)
     mvarCtaBancoAlmz = vData
End Property

Public Property Get CtaBancoAlmz() As String
     CtaBancoAlmz = mvarCtaBancoAlmz
End Property


' forma de pago de facturas positivas almazara

Public Property Let ForpaPosiAlmz(ByVal vData As Integer)
     mvarForpaPosiAlmz = vData
End Property

Public Property Get ForpaPosiAlmz() As Integer
     ForpaPosiAlmz = mvarForpaPosiAlmz
End Property

' forma de pago de facturas negativas almazara

Public Property Let ForpaNegaAlmz(ByVal vData As Integer)
     mvarForpaNegaAlmz = vData
End Property

Public Property Get ForpaNegaAlmz() As Integer
     ForpaNegaAlmz = mvarForpaNegaAlmz
End Property

' cuenta contable ventas

Public Property Let CtaVentasAlmz(ByVal vData As String)
     mvarCtaVentasAlmz = vData
End Property

Public Property Get CtaVentasAlmz() As String
     CtaVentasAlmz = mvarCtaVentasAlmz
End Property

' cuenta contable gastos

Public Property Let CtaGastosAlmz(ByVal vData As String)
     mvarCtaGastosAlmz = vData
End Property

Public Property Get CtaGastosAlmz() As String
     CtaGastosAlmz = mvarCtaGastosAlmz
End Property


' letra de serie de almazara

Public Property Let LetraSerieAlmz(ByVal vData As String)
     mvarLetraSerieAlmz = vData
End Property

Public Property Get LetraSerieAlmz() As String
     LetraSerieAlmz = mvarLetraSerieAlmz
End Property

Public Property Let CodGastoAlmz(ByVal vData As Integer)
     mvarCodGastoAlmz = vData
End Property

Public Property Get CodGastoAlmz() As Integer
     CodGastoAlmz = mvarCodGastoAlmz
End Property

'precio de gasto de molturacion para valsur
Public Property Let GtoMoltura(ByVal vData As String)
     mvarGtoMoltura = vData
End Property

Public Property Get GtoMoltura() As String
     GtoMoltura = mvarGtoMoltura
End Property

'precio de gasto de envasado para valsur
Public Property Let GtoEnvasado(ByVal vData As String)
     mvarGtoEnvasado = vData
End Property

Public Property Get GtoEnvasado() As String
     GtoEnvasado = mvarGtoEnvasado
End Property



' FACTURACION DE ANTICIPOS ALMAZARA
' primera factura de anticipo almazara
Public Property Let PrimFactAntAlmz(ByVal vData As String)
     mvarPrimFactAntAlmz = vData
End Property

Public Property Get PrimFactAntAlmz() As String
     PrimFactAntAlmz = mvarPrimFactAntAlmz
End Property

' ultima factura de anticipo
Public Property Let UltFactAntAlmz(ByVal vData As String)
     mvarUltFactAntAlmz = vData
End Property

Public Property Get UltFactAntAlmz() As String
     UltFactAntAlmz = mvarUltFactAntAlmz
End Property

' FACTURACION DE liquidacion ALMAZARA
' primera factura de liquidacion almazara
Public Property Let PrimFactLiqAlmz(ByVal vData As String)
     mvarPrimFactLiqAlmz = vData
End Property

Public Property Get PrimFactLiqAlmz() As String
     PrimFactLiqAlmz = mvarPrimFactLiqAlmz
End Property

' ultima factura de anticipo almazara
Public Property Let UltFactLiqAlmz(ByVal vData As String)
     mvarUltFactLiqAlmz = vData
End Property

Public Property Get UltFactLiqAlmz() As String
     UltFactLiqAlmz = mvarUltFactLiqAlmz
End Property

'*********************
'      ADV
'*********************

' almacen de adv

Public Property Let AlmacenADV(ByVal vData As Integer)
     mvarAlmacenADV = vData
End Property

Public Property Get AlmacenADV() As Integer
     AlmacenADV = mvarAlmacenADV
End Property


' cuenta contable de banco de ADV

Public Property Let CtaBancoADV(ByVal vData As String)
     mvarCtaBancoADV = vData
End Property

Public Property Get CtaBancoADV() As String
     CtaBancoADV = mvarCtaBancoADV
End Property

' codigo de iva de facturas internas de adv

Public Property Let CodIvaExeADV(ByVal vData As String)
     mvarCodIvaExeADV = vData
End Property

Public Property Get CodIvaExeADV() As String
     CodIvaExeADV = mvarCodIvaExeADV
End Property




'*************************
'      BODEGA
'*************************

Public Property Let AlbRetiradaManual(ByVal vData As Boolean)
     mvarAlbRetiradaManual = vData
End Property

Public Property Get AlbRetiradaManual() As Boolean
     AlbRetiradaManual = mvarAlbRetiradaManual
End Property

' cuenta contable de banco
Public Property Let CtaBancoBOD(ByVal vData As String)
     mvarCtaBancoBOD = vData
End Property

Public Property Get CtaBancoBOD() As String
     CtaBancoBOD = mvarCtaBancoBOD
End Property

' cuenta ventas
Public Property Let CtaVentasBOD(ByVal vData As String)
     mvarCtaVentasBOD = vData
End Property

Public Property Get CtaVentasBOD() As String
     CtaVentasBOD = mvarCtaVentasBOD
End Property

' Porcentaje de gasto de mantenimiento de liquidacion bodega
Public Property Let PorcGtoMantBOD(ByVal vData As String)
     mvarPorcGtoMantBOD = vData
End Property

Public Property Get PorcGtoMantBOD() As String
     PorcGtoMantBOD = mvarPorcGtoMantBOD
End Property

Public Property Let CodGastoBOD(ByVal vData As Integer)
     mvarCodGastoBOD = vData
End Property

Public Property Get CodGastoBOD() As Integer
     CodGastoBOD = mvarCodGastoBOD
End Property

' FACTURACION DE ANTICIPOS bodega
' primera factura de anticipo bodega
Public Property Let PrimFactAntBOD(ByVal vData As String)
     mvarPrimFactAntBOD = vData
End Property

Public Property Get PrimFactAntBOD() As String
     PrimFactAntBOD = mvarPrimFactAntBOD
End Property

' ultima factura de anticipo
Public Property Let UltFactAntBOD(ByVal vData As String)
     mvarUltFactAntBOD = vData
End Property

Public Property Get UltFactAntBOD() As String
     UltFactAntBOD = mvarUltFactAntBOD
End Property

' FACTURACION DE liquidacion bodega
' primera factura de liquidacion bodega
Public Property Let PrimFactLiqBOD(ByVal vData As String)
     mvarPrimFactLiqBOD = vData
End Property

Public Property Get PrimFactLiqBOD() As String
     PrimFactLiqBOD = mvarPrimFactLiqBOD
End Property

' ultima factura de liquidacion bodega
Public Property Let UltFactLiqBOD(ByVal vData As String)
     mvarUltFactLiqBOD = vData
End Property

Public Property Get UltFactLiqBOD() As String
     UltFactLiqBOD = mvarUltFactLiqBOD
End Property

' Porcentaje de incremento de kilos netos
Public Property Let PorcKilosBOD(ByVal vData As String)
     mvarPorcKilosBOD = vData
End Property

Public Property Get PorcKilosBOD() As String
     PorcKilosBOD = mvarPorcKilosBOD
End Property


'*************************
'      SUMINISTROS
'*************************

Public Property Let BDAriges(ByVal vData As String)
     mvarBDAriges = vData
End Property

Public Property Get BDAriges() As String
     BDAriges = mvarBDAriges
End Property


'*************************
'      TELEFONIA
'*************************
' cuenta ventas
Public Property Let CtaVentasTel(ByVal vData As String)
     mvarCtaVentasTel = vData
End Property

Public Property Get CtaVentasTel() As String
     CtaVentasTel = mvarCtaVentasTel
End Property


Public Property Let LetraSerieTel(ByVal vData As String)
     mvarLetraSerieTel = vData
End Property

Public Property Get LetraSerieTel() As String
     LetraSerieTel = mvarLetraSerieTel
End Property


'*********************
'      NOMINAS
'*********************

' almacen de gestion de nominas

Public Property Let AlmacenNOMI(ByVal vData As Integer)
     mvarAlmacenNOMI = vData
End Property

Public Property Get AlmacenNOMI() As Integer
     AlmacenNOMI = mvarAlmacenNOMI
End Property


' Euros por trabajador dia

Public Property Let EurosTrabdiaNOMI(ByVal vData As Currency)
     mvarEurosTrabdiaNOMI = vData
End Property

Public Property Get EurosTrabdiaNOMI() As Currency
     EurosTrabdiaNOMI = mvarEurosTrabdiaNOMI
End Property

Public Property Let PorcSegSo1NOMI(ByVal vData As String)
     mvarPorcSegSo1NOMI = vData
End Property

Public Property Get PorcSegSo1NOMI() As String
     PorcSegSo1NOMI = mvarPorcSegSo1NOMI
End Property

Public Property Let PorcSegSo2NOMI(ByVal vData As String)
     mvarPorcSegSo2NOMI = vData
End Property

Public Property Get PorcSegSo2NOMI() As String
     PorcSegSo2NOMI = mvarPorcSegSo2NOMI
End Property

Public Property Let PorcIRPFNOMI(ByVal vData As String)
     mvarPorcIRPFNOMI = vData
End Property

Public Property Get PorcIRPFNOMI() As String
     PorcIRPFNOMI = mvarPorcIRPFNOMI
End Property

Public Property Let PorcJornadaNOMI(ByVal vData As String)
     mvarPorcJornadaNOMI = vData
End Property

Public Property Get PorcJornadaNOMI() As String
     PorcJornadaNOMI = mvarPorcJornadaNOMI
End Property

Public Property Let NroMaxJornadasNOMI(ByVal vData As String)
     mvarMaxJornadasNOMI = vData
End Property

Public Property Get NroMaxJornadasNOMI() As String
     NroMaxJornadasNOMI = mvarMaxJornadasNOMI
End Property



'*********************
'      POZOS
'*********************

Public Property Let CodIvaPOZ(ByVal vData As String)
     mvarCodIvaPOZ = vData
End Property

Public Property Get CodIvaPOZ() As String
     CodIvaPOZ = mvarCodIvaPOZ
End Property


' Precio de rango 1
Public Property Let Precio1POZ(ByVal vData As String)
     mvarPrecio1POZ = vData
End Property

Public Property Get Precio1POZ() As String
     Precio1POZ = mvarPrecio1POZ
End Property

' Precio de rango 2
Public Property Let Precio2POZ(ByVal vData As String)
     mvarPrecio2POZ = vData
End Property

Public Property Get Precio2POZ() As String
     Precio2POZ = mvarPrecio2POZ
End Property

' Consumo de rango 1
Public Property Let Consumo1POZ(ByVal vData As String)
     mvarConsumo1POZ = vData
End Property

Public Property Get Consumo1POZ() As String
     Consumo1POZ = mvarConsumo1POZ
End Property

' Consumo de rango 2
Public Property Let Consumo2POZ(ByVal vData As String)
     mvarConsumo2POZ = vData
End Property

Public Property Get Consumo2POZ() As String
     Consumo2POZ = mvarConsumo2POZ
End Property


' Cuota pozos
Public Property Let CuotaPOZ(ByVal vData As String)
     mvarCuotaPOZ = vData
End Property

Public Property Get CuotaPOZ() As String
     CuotaPOZ = mvarCuotaPOZ
End Property

' Derrama pozos
Public Property Let DerramaPOZ(ByVal vData As String)
     mvarDerramaPOZ = vData
End Property

Public Property Get DerramaPOZ() As String
     DerramaPOZ = mvarDerramaPOZ
End Property

' cuenta contable consumo ventas pozos

Public Property Let CtaVentasConsPOZ(ByVal vData As String)
     mvarCtaVentasConsPOZ = vData
End Property

Public Property Get CtaVentasConsPOZ() As String
     CtaVentasConsPOZ = mvarCtaVentasConsPOZ
End Property

' cuenta contable cuotas ventas pozos

Public Property Let CtaVentasCuoPOZ(ByVal vData As String)
     mvarCtaVentasCuoPOZ = vData
End Property

Public Property Get CtaVentasCuoPOZ() As String
     CtaVentasCuoPOZ = mvarCtaVentasCuoPOZ
End Property

' cuenta contable talla ventas pozos

Public Property Let CtaVentasTalPOZ(ByVal vData As String)
     mvarCtaVentasTalPOZ = vData
End Property

Public Property Get CtaVentasTalPOZ() As String
     CtaVentasTalPOZ = mvarCtaVentasTalPOZ
End Property

' cuenta contable mantenimiento ventas pozos

Public Property Let CtaVentasMtoPOZ(ByVal vData As String)
     mvarCtaVentasMtoPOZ = vData
End Property

Public Property Get CtaVentasMtoPOZ() As String
     CtaVentasMtoPOZ = mvarCtaVentasMtoPOZ
End Property


' cuenta contable consumo a manta pozos

Public Property Let CtaVentasMantaPOZ(ByVal vData As String)
     mvarCtaVentasMantaPOZ = vData
End Property

Public Property Get CtaVentasMantaPOZ() As String
     CtaVentasMantaPOZ = mvarCtaVentasMantaPOZ
End Property


'[Monica]21/01/2016: cuenta de recargos (escalona)
' cuenta recargos pozos

Public Property Let CtaRecargosPOZ(ByVal vData As String)
     mvarCtaRecargosPOZ = vData
End Property

Public Property Get CtaRecargosPOZ() As String
     CtaRecargosPOZ = mvarCtaRecargosPOZ
End Property


' centro de coste  pozos

Public Property Let CodCCostPOZ(ByVal vData As String)
     mvarCodCCostPoz = vData
End Property

Public Property Get CodCCostPOZ() As String
     CodCCostPOZ = mvarCodCCostPoz
End Property

' forma de pago contado pozos

Public Property Let ForpaConPOZ(ByVal vData As Integer)
     mvarForpaConPOZ = vData
End Property

Public Property Get ForpaConPOZ() As Integer
     ForpaConPOZ = mvarForpaConPOZ
End Property

' forma de pago recibo pozos

Public Property Let ForpaRecPOZ(ByVal vData As Integer)
     mvarForpaRecPOZ = vData
End Property

Public Property Get ForpaRecPOZ() As Integer
     ForpaRecPOZ = mvarForpaRecPOZ
End Property


Public Property Let CartaPOZ(ByVal vData As Integer)
     mvarCartaPOZ = vData
End Property

Public Property Get CartaPOZ() As Integer
     CartaPOZ = mvarCartaPOZ
End Property

Public Property Let NomZonaPOZ(ByVal vData As String)
     mvarNomZonaPOZ = vData
End Property

Public Property Get NomZonaPOZ() As String
     NomZonaPOZ = mvarNomZonaPOZ
End Property




'*********************
'      TRANSPORTE
'*********************

Public Property Let TipoPortesTRA(ByVal vData As Byte)
     mvarTipoPortesTRA = vData
End Property

Public Property Get TipoPortesTRA() As Byte
     TipoPortesTRA = mvarTipoPortesTRA
End Property

' Tarifa Base de Tipo 1
Public Property Let TarifaTRA(ByVal vData As Integer)
     mvarCodTarifTRA = vData
End Property

Public Property Get TarifaTRA() As Integer
     TarifaTRA = mvarCodTarifTRA
End Property


' Tarifa Base de Tipo 2
Public Property Let TarifaTRA2(ByVal vData As Integer)
     mvarCodTarifTRA2 = vData
End Property

Public Property Get TarifaTRA2() As Integer
     TarifaTRA2 = mvarCodTarifTRA2
End Property


Public Property Let CodGastoTRA(ByVal vData As Integer)
     mvarCodGastoTRA = vData
End Property

Public Property Get CodGastoTRA() As Integer
     CodGastoTRA = mvarCodGastoTRA
End Property

Public Property Let PorcreteFacTra(ByVal vData As String)
     mvarPorcReteFacTra = vData
End Property

Public Property Get PorcreteFacTra() As String
     PorcreteFacTra = mvarPorcReteFacTra
End Property

Public Property Let PrecioKgTra(ByVal vData As String)
     mvarPreciokgTra = vData
End Property

Public Property Get PrecioKgTra() As String
     PrecioKgTra = mvarPreciokgTra
End Property


Public Property Let TipoContadorTRA(ByVal vData As Byte)
     mvarTipoContadorTRA = vData
End Property

Public Property Get TipoContadorTRA() As Byte
     TipoContadorTRA = mvarTipoContadorTRA
End Property


' Hay control de costes
Public Property Let HayCCostes(ByVal vData As Byte)
     mvarHayCCostes = vData
End Property

Public Property Get HayCCostes() As Byte
     HayCCostes = mvarHayCCostes
End Property


' Tipo de lectura de Pozos (Escalona)
Public Property Let TipoLecturaPoz(ByVal vData As Byte)
     mvarTipoLecturaPoz = vData
End Property

Public Property Get TipoLecturaPoz() As Byte
     TipoLecturaPoz = mvarTipoLecturaPoz
End Property

Public Property Let LiqDirecta(ByVal vData As Byte)
     mvarLiqDirecta = vData
End Property

Public Property Get LiqDirecta() As Byte
     LiqDirecta = mvarLiqDirecta
End Property


' Hay facturas varias
Public Property Let HayFacVarias(ByVal vData As Byte)
     mvarHayFVarias = vData
End Property

Public Property Get HayFacVarias() As Byte
     HayFacVarias = mvarHayFVarias
End Property


'Facturacion Electrónica FacturaE
Public Property Get PathFacturaE() As String
     PathFacturaE = mvarPathFacturaE
End Property

Public Property Let PathFacturaE(ByVal vData As String)
     mvarPathFacturaE = vData
End Property


'Consumo Minimo de pozos
Public Property Let ConsumoMinPOZ(ByVal vData As Long)
     mvarConsumoMinPoz = vData
End Property

Public Property Get ConsumoMinPOZ() As Long
     ConsumoMinPOZ = mvarConsumoMinPoz
End Property

'Consumo Maximo de Pozos
Public Property Let ConsumoMaxPOZ(ByVal vData As Long)
     mvarConsumoMaxPoz = vData
End Property

Public Property Get ConsumoMaxPOZ() As Long
     ConsumoMaxPOZ = mvarConsumoMaxPoz
End Property


Public Property Let ClaveAcceso(ByVal vData As String)
     mvarClaveAcceso = vData
End Property

Public Property Get ClaveAcceso() As String
     ClaveAcceso = mvarClaveAcceso
End Property

Public Property Let SqlServer(ByVal vData As String)
     mvarSqlServer = vData
End Property

Public Property Get SqlServer() As String
     SqlServer = mvarSqlServer
End Property


Public Property Let SqlServer1(ByVal vData As String)
     mvarSqlServer1 = vData
End Property

Public Property Get SqlServer1() As String
     SqlServer1 = mvarSqlServer1
End Property

' Hay facturas varias
Public Property Let HayResumenCajas(ByVal vData As Byte)
     mvarHayResumenCajas = vData
End Property

Public Property Get HayResumenCajas() As Byte
     HayResumenCajas = mvarHayResumenCajas
End Property



'path de impresion de entradas

Public Property Let PathEntradas(ByVal vData As String)
     mvarPathEntradas = vData
End Property

Public Property Get PathEntradas() As String
     PathEntradas = mvarPathEntradas
End Property





'------------------------------------------------
'FUNCIONES del modulo Cparametros
'------------------------------------------------

'****   LEER  ***
Public Function Leer() As Byte
Dim Rs As ADODB.Recordset
On Error GoTo Err3
    
    Sql = "SELECT * FROM rparam "
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic

    If Rs.EOF Then
        Leer = 1
    Else
        'Datos Contabilidad
        mvarSerConta = DBLet(Rs!serConta, "T") 'servidor conta
        mvarUsuConta = DBLet(Rs!UsuConta, "T") 'usuario conta
        mvarPasConta = DBLet(Rs!pasconta, "T") 'password
        mvarNumConta = DBLet(Rs!NumConta, "N") 'numero conta
        mvarWebSoporte = DBLet(Rs!WebSoporte, "T") 'web soporte
        mvarDireMail = DBLet(Rs!DireMail, "T") 'mail
        mvarSmtphost = DBLet(Rs!Smtphost, "T") 'host
        mvarSmtpUser = DBLet(Rs!SmtpUser, "T") 'user
        mvarSmtpPass = DBLet(Rs!Smtppass, "T") 'pass
        
        mvarEnvioDesdeOutlook = DBLet(Rs!EnvioDesdeOutlook, "N") = 1
        mvarExeEnvioMail = DBLet(Rs!arigesmail, "T")
        
        mvarSMSemail = DBLet(Rs!SMSemail, "T") 'email
        mvarSMSclave = DBLet(Rs!SMSclave, "T") 'clave
        mvarSMSremitente = DBLet(Rs!SMSremitente, "T") 'remitente
        
        
        'entrada en bascula
        mvarTipoCaja1 = DBLet(Rs!TipoCaja1, "T") ' tipo de caja 1
        mvarTipoCaja2 = DBLet(Rs!TipoCaja2, "T") ' tipo de caja 2
        mvarTipoCaja3 = DBLet(Rs!TipoCaja3, "T") ' tipo de caja 3
        mvarTipoCaja4 = DBLet(Rs!TipoCaja4, "T") ' tipo de caja 4
        mvarTipoCaja5 = DBLet(Rs!TipoCaja5, "T") ' tipo de caja 5
        
        mvarPesoCaja1 = DBLet(Rs!PesoCaja1, "N") ' peso caja 1
        mvarPesoCaja2 = DBLet(Rs!PesoCaja2, "N") ' peso caja 2
        mvarPesoCaja3 = DBLet(Rs!PesoCaja3, "N") ' peso caja 3
        mvarPesoCaja4 = DBLet(Rs!PesoCaja4, "N") ' peso caja 4
        mvarPesoCaja5 = DBLet(Rs!PesoCaja5, "N") ' peso caja 5
        
        mvarPesoCaja11 = DBLet(Rs!PesoCaja11, "N") ' peso caja 1 transportista
        mvarPesoCaja12 = DBLet(Rs!PesoCaja12, "N") ' peso caja 2 transportista
        mvarPesoCaja13 = DBLet(Rs!PesoCaja13, "N") ' peso caja 3 transportista
        mvarPesoCaja14 = DBLet(Rs!PesoCaja14, "N") ' peso caja 4 transportista
        mvarPesoCaja15 = DBLet(Rs!PesoCaja15, "N") ' peso caja 5 transportista
        
        
        
        
        mvarKilosCajaMin = DBLet(Rs!KilosCajaMin, "N") ' limite kilos caja minimo
        mvarKilosCajaMax = DBLet(Rs!KilosCajaMax, "N") ' limite kilos caja maximo
        
        mvarPesoCajaLlena = DBLet(Rs!PesoCajaLLena, "N") ' peso caja llena
        
        mvarEsCaja1 = DBLet(Rs!EsCaja1, "N") ' es caja 1
        mvarEsCaja2 = DBLet(Rs!EsCaja2, "N") ' es caja 2
        mvarEsCaja3 = DBLet(Rs!EsCaja3, "N") ' es caja 3
        mvarEsCaja4 = DBLet(Rs!EsCaja4, "N") ' es caja 4
        mvarEsCaja5 = DBLet(Rs!EsCaja5, "N") ' es caja 5
        
        
        mvarEsVtaFruta1 = DBLet(Rs!EsVtaFruta1, "N") ' es VtaFruta 1
        mvarEsVtaFruta2 = DBLet(Rs!EsVtaFruta2, "N") ' es VtaFruta 2
        mvarEsVtaFruta3 = DBLet(Rs!EsVtaFruta3, "N") ' es VtaFruta 3
        mvarEsVtaFruta4 = DBLet(Rs!EsVtaFruta4, "N") ' es VtaFruta 4
        mvarEsVtaFruta5 = DBLet(Rs!EsVtaFruta5, "N") ' es VtaFruta 5
        
        mvarSeTaraTractor = DBLet(Rs!SeTaraTractor, "B") 'se tara el tractor o no
        mvarSeAgrupanNotas = DBLet(Rs!agruparnotas, "B") 'se agrupan notas si o no
        mvarSeRespetaNota = DBLet(Rs!SeRespetaNota, "B") 'se respeta el nro de nota si o no
        mvarHayTraza = DBLet(Rs!HayTraza, "B") 'Si hay o no trazabilidad
        mvarNroNotaManual = DBLet(Rs!NroNotaManual, "B") 'nro de nota de entrada manual o no
        mvarCooproenEntradas = DBLet(Rs!CooproEntradas, "B") '0=entrada desdoblada en act.entradas clasificadas/1=entrada desdoblada en act.entradas bascula
        mvarCajasporPalet = DBLet(Rs!CajasporPalet, "N") 'Nro de Cajas por Palet
        
        mvarPorcIncreAforo = DBLet(Rs!PorcIncreAforo, "N") 'porcentaje de incremento/decremento aforo
        
        mvarImpresoraEntradas = DBLet(Rs!ImpresoraEntradas, "T") ' impresora de entradas
        mvarCodIncidPlaga = DBLet(Rs!CodIncidPlaga, "N") ' codigo de incidencia para entradas de quatretonda
        
        'aridoc
        mvarHayAridoc = DBLet(Rs!HayAridoc, "N") ' hay aridoc
        mvarCarpetaAnt = DBLet(Rs!codcarpetaanti, "N") 'carpeta albaran
        mvarCarpetaLiq = DBLet(Rs!codcarpetaliqu, "N") 'carpeta factura
        mvarCarpetaADV = DBLet(Rs!codcarpetaADV, "N") 'carpeta ADV
        mvarCarpetaRecAlmacen = DBLet(Rs!codcarpetarecalm, "N") 'carpeta recibo almacen
        mvarCarpetaRecCampo = DBLet(Rs!codcarpetareccamp, "N") 'carpeta recibo campo
        mvarCarpetaAlmz = DBLet(Rs!codcarpetaAlmz, "N") 'carpeta Almazara
        mvarCarpetaBOD = DBLet(Rs!codcarpetabode, "N") 'carpeta Bodega
        mvarCarpetaTra = DBLet(Rs!codcarpetatran, "N") 'carpeta transporte
        mvarExtension = DBLet(Rs!codextension, "N") 'extension documento
        mvarC1Anticipo = DBLet(Rs!c1antiaridoc, "N") 'campo 1 anticipo
        mvarC2Anticipo = DBLet(Rs!c2antiaridoc, "N") 'campo 2 anticipo
        mvarC3Anticipo = DBLet(Rs!c3antiaridoc, "N") 'campo 3 anticipo
        mvarC4Anticipo = DBLet(Rs!c4antiaridoc, "N") 'campo 4 anticipo
        mvarC1Liquidacion = DBLet(Rs!c1liquaridoc, "N") 'campo 1 liquidacion
        mvarC2Liquidacion = DBLet(Rs!c2liquaridoc, "N") 'campo 2 liquidacion
        mvarC3Liquidacion = DBLet(Rs!c3liquaridoc, "N") 'campo 3 liquidacion
        mvarC4Liquidacion = DBLet(Rs!c4liquaridoc, "N") 'campo 4 liquidacion
        mvarC1ADV = DBLet(Rs!c1ADVaridoc, "N") 'campo 1 ADV
        mvarC2ADV = DBLet(Rs!c2ADVaridoc, "N") 'campo 2 ADV
        mvarC3ADV = DBLet(Rs!c3ADVaridoc, "N") 'campo 3 ADV
        mvarC4ADV = DBLet(Rs!c4ADVaridoc, "N") 'campo 4 ADV
        mvarC1Recibo = DBLet(Rs!c1recaridoc, "N") 'campo 1 recibo
        mvarC2Recibo = DBLet(Rs!c2recaridoc, "N") 'campo 2 recibo
        mvarC3Recibo = DBLet(Rs!c3recaridoc, "N") 'campo 3 recibo
        mvarC4Recibo = DBLet(Rs!c4recaridoc, "N") 'campo 4 recibo
        
        mvarC1Almazara = DBLet(Rs!c1almzaridoc, "N") 'campo 1 Almazara
        mvarC2Almazara = DBLet(Rs!c2almzaridoc, "N") 'campo 2 Almazara
        mvarC3Almazara = DBLet(Rs!c3almzaridoc, "N") 'campo 3 Almazara
        mvarC4Almazara = DBLet(Rs!c4almzaridoc, "N") 'campo 4 Almazara
        mvarC1Bodega = DBLet(Rs!c1bodearidoc, "N") 'campo 1 bodega
        mvarC2Bodega = DBLet(Rs!c2bodearidoc, "N") 'campo 2 bodega
        mvarC3Bodega = DBLet(Rs!c3bodearidoc, "N") 'campo 3 bodega
        mvarC4Bodega = DBLet(Rs!c4bodearidoc, "N") 'campo 4 bodega
        
        mvarC1Transporte = DBLet(Rs!c1tranaridoc, "N") 'campo 1 Transporte
        mvarC2Transporte = DBLet(Rs!c2tranaridoc, "N") 'campo 2 Transporte
        mvarC3Transporte = DBLet(Rs!c3tranaridoc, "N") 'campo 3 Transporte
        mvarC4Transporte = DBLet(Rs!c4tranaridoc, "N") 'campo 4 Transporte
        
        mvarCosteHora = DBLet(Rs!CosteHora, "N") ' coste hora
        mvarCosteSegSo = DBLet(Rs!CosteSegSo, "N") ' coste seguridad social
        mvarSeccionHorto = DBLet(Rs!Seccionhorto, "N") ' seccion hortofruticola
        mvarSeccionPozos = DBLet(Rs!SeccionPOZOS, "N") ' seccion pozos
        mvarSeccionAlmaz = DBLet(Rs!SeccionAlmaz, "N") ' seccion almazara
        mvarSeccionSumi = DBLet(Rs!SeccionSumi, "N") ' seccion suministros
        mvarSeccionADV = DBLet(Rs!SeccionADV, "N") ' seccion adv
        mvarSeccionBodega = DBLet(Rs!SeccionBodega, "N") ' seccion bodega
        
        mvarFaneca = DBLet(Rs!Faneca, "N") 'constante faneca
        mvarPorcReteFacSoc = DBLet(Rs!PorcretenFacSoc, "N") 'porcentaje de retencion de la factura
        mvarPrimFactAnt = DBLet(Rs!PrimFactAnt, "N") 'primera factura de anticipo
        mvarUltFactAnt = DBLet(Rs!UltFactAnt, "N") 'ultima factura de anticipo
        mvarPrimFactAntVC = DBLet(Rs!PrimFactAntVC, "N") 'primera factura de anticipo venta campo
        mvarUltFactAntVC = DBLet(Rs!UltFactAntVC, "N") 'ultima factura de anticipo venta campo
        mvarPrimFactLiq = DBLet(Rs!PrimFactLiq, "N") 'primera factura de liquidacion
        mvarUltFactLiq = DBLet(Rs!UltFactLiq, "N") 'ultima factura de liquidacion
        mvarPrimFactLiqVC = DBLet(Rs!PrimFactLiqVC, "N") 'primera factura de liquidacion venta campo
        mvarUltFactLiqVC = DBLet(Rs!UltFactLiqVC, "N") 'ultima factura de liquidacion venta campo
        mvarPorcenAFO = DBLet(Rs!PorcenAFO, "N") 'porcentaje de fondo operativo
        mvarTTomaDatos = DBLet(Rs!texttomadatos, "T")  ' texto de impresion de toma de datos
        mvarPieTomaDatos = DBLet(Rs!PieTomaDatos, "T") ' texto pie impresion de toma de datos
        
        mvarCooperativa = DBLet(Rs!Cooperativa, "N") 'cooperativa 0=catadau 1=valsur 2=picassent
        
        mvarCodIvaIntra = DBLet(Rs!codivaintracom, "T") ' codigo de iva intracomunitario
        mvarCtaTerReten = DBLet(Rs!CtaTerReten, "T") ' cta contable de retencion de facturas de terceros
        mvarCtaTraReten = DBLet(Rs!CtaTraReten, "T") ' cta contable de retencion de facturas de transporte
        
        mvarPathTraza = DBLet(Rs!directoriotraza, "T") ' directorio de trazabilidad
        
        mvarSigpac = DBLet(Rs!SigPac, "T") ' direccion de sigpac
        mvarGoolZoom = DBLet(Rs!GoolZoom, "T") ' direccion de goolzoom
        
        
        'contabilizacion de facturas de socios (liquidaciones / anticipos)
        mvarCtaRetenSoc = DBLet(Rs!CtaRetenSoc, "T") ' cta contable de retencion de socios
        mvarCtaAportaSoc = DBLet(Rs!CtaAportaSoc, "T") ' cta contable de aportacion de facturas de socios
        mvarCtaBancoSoc = DBLet(Rs!CtaBanco, "T") ' cta contable de banco prevista de facturas de socios
        
        mvarForpaPosi = DBLet(Rs!codforpaposi, "N") ' forma de pago de facturas en positivo
        mvarForpaNega = DBLet(Rs!Codforpanega, "N") ' forma de pago de facturas en negativo
        
        
        'contabilizacion de facturas de liquidacion de almazara
        mvarCtaRetenAlmz = DBLet(Rs!CtaRetenAlmz, "T") ' cta contable de retencion de almazara
        mvarCtaBancoAlmz = DBLet(Rs!CtaBancoAlmz, "T") ' cta contable de banco prevista de almazara
        mvarCtaVentasAlmz = DBLet(Rs!CtaVentasAlmz, "T") ' cta contable de ventas de almazara
        mvarCtaGastosAlmz = DBLet(Rs!CtaGastosAlmz, "T") ' cta contable de gastos de almazara
        mvarLetraSerieAlmz = DBLet(Rs!LetraSerieAlmz, "T") ' letra de serie de almazara
        
        mvarForpaPosiAlmz = DBLet(Rs!codforpaposalmz, "N") ' forma de pago de facturas en positivo almazara
        mvarForpaNegaAlmz = DBLet(Rs!Codforpanegalmz, "N") ' forma de pago de facturas en negativo almazara
        
        mvarPrimFactAntAlmz = DBLet(Rs!PrimFactAntAlmz, "N") 'primera factura de anticipo almazara
        mvarUltFactAntAlmz = DBLet(Rs!UltFactAntAlmz, "N") 'ultima factura de anticipo almazara
        mvarPrimFactLiqAlmz = DBLet(Rs!PrimFactLiqAlmz, "N") 'primera factura de liquidacion almazara
        mvarUltFactLiqAlmz = DBLet(Rs!UltFactLiqAlmz, "N") 'ultima factura de liquidacion almazara
        
        mvarCodGastoAlmz = DBLet(Rs!CodGastoAlmz, "N") ' codigo de gasto para liquidacion de almazara
        mvarGtoMoltura = DBLet(Rs!GtoMoltura, "N") 'precio por kilo de gastos de molturacion
        mvarGtoEnvasado = DBLet(Rs!GtoEnvasado, "N") 'precio por litro de gastos de envasado
        
        mvarAlmacenADV = DBLet(Rs!codalmacadv, "N") ' codigo de almacen de adv
        mvarCtaBancoADV = DBLet(Rs!CtaBancoADV, "T") ' cta contable de banco prevista de adv
        mvarCodIvaExeADV = DBLet(Rs!CodIvaExeADV, "N") ' codigo de iva de facturas internas de pozos
        
        
        'suministros
        mvarBDAriges = DBLet(Rs!BDAriges, "T") ' bdariges si tiene suministros
        
        'bodega
        mvarAlbRetiradaManual = DBLet(Rs!albretiradabodman, "B") 'nro de albaran de retirada de bodega manual
        mvarCtaBancoBOD = DBLet(Rs!CtaBancoBOD, "T") ' cta contable de banco prevista para bodega
        mvarCtaVentasBOD = DBLet(Rs!CtaVentasBOD, "T") ' cta contable de ventas para bodega
        mvarPorcGtoMantBOD = DBLet(Rs!bodporcenmant, "N") 'porcentaje de gasto de mantenimiento de liquidacion bodega
        mvarCodGastoBOD = DBLet(Rs!CodGastoBOD, "N") ' codigo de gasto para liquidacion de bodega
        
        mvarPrimFactAntBOD = DBLet(Rs!PrimFactAntBOD, "N") 'primera factura de anticipo bodega
        mvarUltFactAntBOD = DBLet(Rs!UltFactAntBOD, "N") 'ultima factura de anticipo bodega
        mvarPrimFactLiqBOD = DBLet(Rs!PrimFactLiqBOD, "N") 'primera factura de liquidacion bodega
        mvarUltFactLiqBOD = DBLet(Rs!UltFactLiqBOD, "N") 'ultima factura de liquidacion bodega
        mvarPorcKilosBOD = DBLet(Rs!PorcKilosBOD, "N") 'porcentaje de incremento de kilos netos de entradas de bodega
        
        'telefonia
        mvarLetraSerieTel = DBLet(Rs!LetraSerieTel, "T") ' letra de serie de telefonia
        mvarCtaVentasTel = DBLet(Rs!CtaVentasTel, "T") ' cta contable de ventas para telefonia
        
        'gestion de nominas
        mvarAlmacenNOMI = DBLet(Rs!codalmacnomi, "N") ' codigo de almacen de gestion de nominas
        mvarEurosTrabdiaNOMI = DBLet(Rs!EurosTrabdia, "N") ' euros por trabajador dia nomina
        mvarPorcSegSo1NOMI = DBLet(Rs!porcsegso1, "N") 'porcentaje de seguridad social de nomina
        mvarPorcSegSo2NOMI = DBLet(Rs!porcsegso2, "N") 'porcentaje de seguridad social de nomina
        mvarPorcIRPFNOMI = DBLet(Rs!porcirpf, "N") ' porcentaje de irpf de nominas
        mvarPorcJornadaNOMI = DBLet(Rs!porcjornada, "N") ' porcentaje de irpf de nominas
        mvarMaxJornadasNOMI = DBLet(Rs!nromaxjornadas, "N") ' nro maximo de jornadas de nominas
        
        ' pozos
        mvarConsumo1POZ = DBLet(Rs!hastametcub1poz, "N") ' hasta m3 del primer rango
        mvarConsumo2POZ = DBLet(Rs!hastametcub2poz, "N") ' hasta m3 del segundo rango
        mvarPrecio1POZ = DBLet(Rs!Precio1POZ, "N") ' precio del rango 1
        mvarPrecio2POZ = DBLet(Rs!Precio2POZ, "N") ' precio del rango 2
        mvarCodIvaPOZ = DBLet(Rs!CodIvaPOZ, "N") ' codigo de iva de pozos
        mvarCuotaPOZ = DBLet(Rs!imporcuotapoz, "N") ' importe de cuota
        mvarDerramaPOZ = DBLet(Rs!imporderramapoz, "N") ' importe derrama
        mvarCtaVentasConsPOZ = DBLet(Rs!CtaVentasConsPOZ, "T") ' cta contable de ventas consumo de pozos
        mvarCtaVentasCuoPOZ = DBLet(Rs!CtaVentasCuoPOZ, "T") ' cta contable de ventas cuota de pozos
        mvarCtaVentasTalPOZ = DBLet(Rs!CtaVentasTalPOZ, "T") ' cta contable de ventas talla de pozos
        mvarCtaVentasMtoPOZ = DBLet(Rs!CtaVentasMtoPOZ, "T") ' cta contable de ventas mantenimiento de pozos
        mvarCtaVentasMantaPOZ = DBLet(Rs!CtaVentasMantaPOZ, "T") ' cta contable de ventas consumo a manta de pozos
        '[Monica]21/01/2016: cuenta de recargos (escalona)
        mvarCtaRecargosPOZ = DBLet(Rs!CtaRecargosPOZ, "T") ' cta contable de recargos de pozos
        
        mvarCodCCostPoz = DBLet(Rs!CodCCostPOZ, "T") ' centro de coste de pozos
        mvarForpaConPOZ = DBLet(Rs!ForpaConPOZ, "N") ' forma de pago contado pozos
        mvarForpaRecPOZ = DBLet(Rs!ForpaRecPOZ, "N") ' forma de pago recibo pozos
        mvarCartaPOZ = DBLet(Rs!codcartapoz, "N") ' codigo de carta de reclamacion de pozos
        mvarNomZonaPOZ = DBLet(Rs!NomZonaPOZ, "T") ' nombre de la zona de pozos para Escalona y Utxera
        
        
        ' transporte
        mvarTipoPortesTRA = DBLet(Rs!tratipoportes, "N") ' tipo de portes de transporte
        mvarCodTarifTRA = DBLet(Rs!tracodtarif, "N") ' codigo de tarifa de transporte de tipo 1
        mvarCodTarifTRA2 = DBLet(Rs!tracodtarif2, "N") ' codigo de tarifa de transporte de tipo 2
        mvarCodGastoTRA = DBLet(Rs!tracodgasto, "N") ' codigo de gasto de transporte
        mvarPorcReteFacTra = DBLet(Rs!PorcretenFactra, "N") 'porcentaje de retencion de la factura de transporte
        mvarPreciokgTra = DBLet(Rs!Preciotra, "N") 'precio por kilo de transporte transporte
        mvarTipoContadorTRA = DBLet(Rs!tratipocontador, "N") ' tipo de contador de transporte
        
        
        mvarHayCCostes = DevuelveValor("select hayccostes from sparam") ' hay control de costes en la parte de recoleccion
        
        mvarTipoLecturaPoz = DBLet(Rs!TipoLecturaPoz, "N") ' tipo de lectura de pozos
        
        mvarLiqDirecta = DBLet(Rs!LiqDirecta, "N") ' liquidacion directa
        
        mvarHayFVarias = DBLet(Rs!hayfactvarias, "N") ' Hya Facturas Varias
        
        '[Monica]07/03/2013: necesitamos el path de la facturacion electronica
        mvarPathFacturaE = DevuelveValor("select pathfacturae from sparam")
        If mvarPathFacturaE = "0" Then mvarPathFacturaE = ""
        
        mvarConsumoMinPoz = DBLet(Rs!ConsumoMinPOZ, "N") 'Consumo Minimo de Pozos
        mvarConsumoMaxPoz = DBLet(Rs!ConsumoMaxPOZ, "N") 'Consumo Maximo de Pozos
        
        '[Monica]02/12/2014:
        mvarClaveAcceso = DBLet(Rs!ClaveAcceso, "T") ' Clave de Acceso
        
        '[Monica]21/04/2016:
        mvarSqlServer = DBLet(Rs!SqlServer, "T") ' SqlServer CASTELLDUCS1-O5
        mvarSqlServer1 = DBLet(Rs!SqlServer1, "T") ' SqlServer1 CASTELLDUCS1-O5
        
        '[Monica]10/10/2016: resumen de cajas en el mto de partes de trabajo
        mvarHayResumenCajas = DBLet(Rs!HayResumenCajas, "N") ' Hay Resumen Cajas
                
        '[Monica]17/10/2016: directorio de grabacion de entradas
        mvarPathEntradas = DBLet(Rs!directorioentradas, "T") ' directorio de impresion de entradas
        
        
        
        Leer = 0
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Rs.Close
    Set Rs = Nothing
    Leer = 1
End Function


Public Function Modificar() As Byte
'Modifica los parametros de la Aplicación
'Modifica la Tabla: sparam, BD: Arigasol
On Error GoTo eModificar

    Sql = "UPDATE rparam SET "
    Sql = Sql & " serconta= '" & mvarSerConta & "'"
    Sql = Sql & ", usuconta= " & DBSet(mvarUsuConta, "T")
    Sql = Sql & ", pasconta= " & DBSet(mvarPasConta, "T")
    Sql = Sql & ", numconta= " & DBSet(mvarNumConta, "N", "S")
    Sql = Sql & ", websoporte= " & DBSet(mvarWebSoporte, "T")
    Sql = Sql & ", diremail= " & DBSet(mvarDireMail, "T")
    Sql = Sql & ", smtphost= " & DBSet(mvarSmtphost, "T")
    Sql = Sql & ", smtpuser= " & DBSet(mvarSmtpUser, "T")
    Sql = Sql & ", smtppass= " & DBSet(mvarSmtpPass, "T")
    
    '[Monica]06/02/2012: arigesmail
    Sql = Sql & ", EnvioDesdeOutlook = " & DBSet(Me.EnvioDesdeOutlook, "B")
    Sql = Sql & ", arigesmail = " & DBSet(ExeEnvioMail, "T", "S")
    
    Sql = Sql & ", smsemail= " & DBSet(mvarSMSemail, "T")
    Sql = Sql & ", smsclave= " & DBSet(mvarSMSclave, "T")
    Sql = Sql & ", smsremitente= " & DBSet(mvarSMSremitente, "T")
    
    'entrada en bascula
    Sql = Sql & ", setaratractor= " & DBSet(mvarSeTaraTractor, "B")
    Sql = Sql & ", agruparnotas= " & DBSet(mvarSeAgrupanNotas, "B")
    Sql = Sql & ", serespetanota= " & DBSet(mvarSeRespetaNota, "B")
    Sql = Sql & ", haytraza= " & DBSet(mvarHayTraza, "B")
    Sql = Sql & ", nronotamanual= " & DBSet(mvarNroNotaManual, "B")
    Sql = Sql & ", cooproentradas= " & DBSet(mvarCooproenEntradas, "B")
    Sql = Sql & ", cajasporpalet= " & DBSet(mvarCajasporPalet, "N")
    
    
    Sql = Sql & ", tipocaja1= " & DBSet(mvarTipoCaja1, "T")
    Sql = Sql & ", tipocaja2= " & DBSet(mvarTipoCaja2, "T")
    Sql = Sql & ", tipocaja3= " & DBSet(mvarTipoCaja3, "T")
    Sql = Sql & ", tipocaja4= " & DBSet(mvarTipoCaja4, "T")
    Sql = Sql & ", tipocaja5= " & DBSet(mvarTipoCaja5, "T")
    
    Sql = Sql & ", pesocaja1= " & DBSet(mvarPesoCaja1, "N")
    Sql = Sql & ", pesocaja2= " & DBSet(mvarPesoCaja2, "N")
    Sql = Sql & ", pesocaja3= " & DBSet(mvarPesoCaja3, "N")
    Sql = Sql & ", pesocaja4= " & DBSet(mvarPesoCaja4, "N")
    Sql = Sql & ", pesocaja5= " & DBSet(mvarPesoCaja5, "N")
    
    Sql = Sql & ", pesocaja11= " & DBSet(mvarPesoCaja11, "N")
    Sql = Sql & ", pesocaja12= " & DBSet(mvarPesoCaja12, "N")
    Sql = Sql & ", pesocaja13= " & DBSet(mvarPesoCaja13, "N")
    Sql = Sql & ", pesocaja14= " & DBSet(mvarPesoCaja14, "N")
    Sql = Sql & ", pesocaja15= " & DBSet(mvarPesoCaja15, "N")
    
    
    Sql = Sql & ", pesocajallena= " & DBSet(mvarPesoCajaLlena, "N")
    
    Sql = Sql & ", escaja1= " & DBSet(mvarEsCaja1, "B")
    Sql = Sql & ", escaja2= " & DBSet(mvarEsCaja2, "B")
    Sql = Sql & ", escaja3= " & DBSet(mvarEsCaja3, "B")
    Sql = Sql & ", escaja4= " & DBSet(mvarEsCaja4, "B")
    Sql = Sql & ", escaja5= " & DBSet(mvarEsCaja5, "B")
    
    Sql = Sql & ", esvtafruta1= " & DBSet(mvarEsVtaFruta1, "B")
    Sql = Sql & ", esvtafruta2= " & DBSet(mvarEsVtaFruta2, "B")
    Sql = Sql & ", esvtafruta3= " & DBSet(mvarEsVtaFruta3, "B")
    Sql = Sql & ", esvtafruta4= " & DBSet(mvarEsVtaFruta4, "B")
    Sql = Sql & ", esvtafruta5= " & DBSet(mvarEsVtaFruta5, "B")
    
    Sql = Sql & ", kiloscajamin= " & DBSet(mvarKilosCajaMin, "N")
    Sql = Sql & ", kiloscajamax= " & DBSet(mvarKilosCajaMax, "N")
    
    Sql = Sql & ", impresoraentradas= " & DBSet(mvarImpresoraEntradas, "T")
    Sql = Sql & ", codincidplaga= " & DBSet(mvarCodIncidPlaga, "N") ' codigo de incidencia de plagas en entradas bascula de quatretonda
    
    Sql = Sql & ", porcincreaforo= " & DBSet(mvarPorcIncreAforo, "N") 'porcentaje de incremento / decremento de aforo
    
    'aridoc
    Sql = Sql & ", hayaridoc= " & DBSet(mvarHayAridoc, "N") ' hay aridoc
    Sql = Sql & ", codcarpetaanti = " & DBSet(mvarCarpetaAnt, "N") 'carpeta anticipo
    Sql = Sql & ", codcarpetaliqu = " & DBSet(mvarCarpetaLiq, "N") 'carpeta liquidacion
    Sql = Sql & ", codcarpetarecalm = " & DBSet(mvarCarpetaRecAlmacen, "N") 'carpeta recibo almacen
    Sql = Sql & ", codcarpetareccamp = " & DBSet(mvarCarpetaRecCampo, "N") 'carpeta recibo campo
    Sql = Sql & ", codcarpetaadv = " & DBSet(mvarCarpetaADV, "N") 'carpeta adv
    
    Sql = Sql & ", codcarpetaalmz = " & DBSet(mvarCarpetaAlmz, "N") 'carpeta almazara
    Sql = Sql & ", codcarpetabode = " & DBSet(mvarCarpetaBOD, "N") 'carpeta bodega
    
    Sql = Sql & ", codcarpetatran = " & DBSet(mvarCarpetaTra, "N") 'carpeta transportista
    
    Sql = Sql & ", codextension = " & DBSet(mvarExtension, "N") 'extension documento
    Sql = Sql & ", c1antiaridoc = " & DBSet(mvarC1Anticipo, "N") 'campo 1 anticipo
    Sql = Sql & ", c2antiaridoc = " & DBSet(mvarC2Anticipo, "N") 'campo 2 anticipo
    Sql = Sql & ", c3antiaridoc = " & DBSet(mvarC3Anticipo, "N") 'campo 3 anticipo
    Sql = Sql & ", c4antiaridoc = " & DBSet(mvarC4Anticipo, "N") 'campo 4 anticipo
    
    Sql = Sql & ", c1liquaridoc = " & DBSet(mvarC1Liquidacion, "N") 'campo 1 liquidacion
    Sql = Sql & ", c2liquaridoc = " & DBSet(mvarC2Liquidacion, "N") 'campo 2 liquidacion
    Sql = Sql & ", c3liquaridoc = " & DBSet(mvarC3Liquidacion, "N") 'campo 3 liquidacion
    Sql = Sql & ", c4liquaridoc = " & DBSet(mvarC4Liquidacion, "N") 'campo 4 liquidacion
    
    Sql = Sql & ", c1advaridoc = " & DBSet(mvarC1ADV, "N") 'campo 1 adv
    Sql = Sql & ", c2advaridoc = " & DBSet(mvarC2ADV, "N") 'campo 2 adv
    Sql = Sql & ", c3advaridoc = " & DBSet(mvarC3ADV, "N") 'campo 3 adv
    Sql = Sql & ", c4advaridoc = " & DBSet(mvarC4ADV, "N") 'campo 4 adv
    
    Sql = Sql & ", c1recaridoc = " & DBSet(mvarC1Recibo, "N") 'campo 1 recibo
    Sql = Sql & ", c2recaridoc = " & DBSet(mvarC2Recibo, "N") 'campo 2 recibo
    Sql = Sql & ", c3recaridoc = " & DBSet(mvarC3Recibo, "N") 'campo 3 recibo
    Sql = Sql & ", c4recaridoc = " & DBSet(mvarC4Recibo, "N") 'campo 4 recibo
    
    Sql = Sql & ", c1almzaridoc = " & DBSet(mvarC1Almazara, "N") 'campo 1 almazara
    Sql = Sql & ", c2almzaridoc = " & DBSet(mvarC2Almazara, "N") 'campo 2 almazara
    Sql = Sql & ", c3almzaridoc = " & DBSet(mvarC3Almazara, "N") 'campo 3 almazara
    Sql = Sql & ", c4almzaridoc = " & DBSet(mvarC4Almazara, "N") 'campo 4 almazara
    
    Sql = Sql & ", c1bodearidoc = " & DBSet(mvarC1Bodega, "N") 'campo 1 bodega
    Sql = Sql & ", c2bodearidoc = " & DBSet(mvarC2Bodega, "N") 'campo 2 bodega
    Sql = Sql & ", c3bodearidoc = " & DBSet(mvarC3Bodega, "N") 'campo 3 bodega
    Sql = Sql & ", c4bodearidoc = " & DBSet(mvarC4Bodega, "N") 'campo 4 bodega
    
    Sql = Sql & ", c1tranaridoc = " & DBSet(mvarC1Transporte, "N") 'campo 1 Transporte
    Sql = Sql & ", c2tranaridoc = " & DBSet(mvarC2Transporte, "N") 'campo 2 Transporte
    Sql = Sql & ", c3tranaridoc = " & DBSet(mvarC3Transporte, "N") 'campo 3 Transporte
    Sql = Sql & ", c4tranaridoc = " & DBSet(mvarC4Transporte, "N") 'campo 4 Transporte
    
    
    Sql = Sql & ", costehora= " & DBSet(mvarCosteHora, "N") 'coste hora
    Sql = Sql & ", costesegso= " & DBSet(mvarCosteSegSo, "N") 'coste seguridad social
    If mvarSeccionHorto = "" Then
        Sql = Sql & ", seccionhorto= null "
    Else
        Sql = Sql & ", seccionhorto= " & DBSet(mvarSeccionHorto, "N") 'seccion hortofruticola
    End If
    If mvarSeccionPozos = "" Then
        Sql = Sql & ", seccionpozos= null "
    Else
        Sql = Sql & ", seccionpozos= " & DBSet(mvarSeccionPozos, "N") 'seccion pozos
    End If
    If mvarSeccionAlmaz = "" Then
        Sql = Sql & ", seccionalmaz= null "
    Else
        Sql = Sql & ", seccionalmaz= " & DBSet(mvarSeccionAlmaz, "N") 'seccion almazara
    End If
    If mvarSeccionADV = "" Then
        Sql = Sql & ", seccionadv= null "
    Else
        Sql = Sql & ", seccionadv= " & DBSet(mvarSeccionADV, "N") 'seccion adv
    End If
    If mvarSeccionSumi = "" Then
        Sql = Sql & ", seccionsumi= null "
    Else
        Sql = Sql & ", seccionsumi= " & DBSet(mvarSeccionSumi, "N") 'seccion suministros
    End If
    If mvarSeccionBodega = "" Then
        Sql = Sql & ", seccionbodega= null "
    Else
        Sql = Sql & ", seccionbodega= " & DBSet(mvarSeccionBodega, "N") 'seccion bodega
    End If
    
    Sql = Sql & ", faneca= " & DBSet(mvarFaneca, "N") ' constante faneca
    Sql = Sql & ", porcretenfacsoc= " & DBSet(mvarPorcReteFacSoc, "N", "S")
    Sql = Sql & ", primfactant= " & DBSet(mvarPrimFactAnt, "N", "S") ' primera factura anticipos
    Sql = Sql & ", ultfactant= " & DBSet(mvarUltFactAnt, "N", "S")   ' ultima factura anticipos
    Sql = Sql & ", primfactantvc= " & DBSet(mvarPrimFactAntVC, "N", "S") ' primera factura anticipos venta campo
    Sql = Sql & ", ultfactantvc= " & DBSet(mvarUltFactAntVC, "N", "S")   ' ultima factura anticipos venta campo
    Sql = Sql & ", primfactliq= " & DBSet(mvarPrimFactLiq, "N", "S") ' primera factura liquidacion
    Sql = Sql & ", ultfactliq= " & DBSet(mvarUltFactLiq, "N", "S")   ' ultima factura liquidacion
    Sql = Sql & ", primfactliqvc= " & DBSet(mvarPrimFactLiqVC, "N", "S") ' primera factura liquidacion venta campo
    Sql = Sql & ", ultfactliqvc= " & DBSet(mvarUltFactLiqVC, "N", "S")   ' ultima factura liquidacion venta campo
    Sql = Sql & ", porcenafo= " & DBSet(mvarPorcenAFO, "N", "S") ' porcentaje de fondo operativo
    Sql = Sql & ", texttomadatos= " & DBSet(mvarTTomaDatos, "T", "S") ' texto de impresion de toma de datos
    Sql = Sql & ", pietomadatos= " & DBSet(mvarPieTomaDatos, "T", "S") ' texto de impresion de toma de datos
    
    Sql = Sql & ", cooperativa= " & DBSet(mvarCooperativa, "N") ' cooperativa en la que estoy
                                                                '0=catadau   1=valsur    2=catadau
    Sql = Sql & ", codivaintracom= " & DBSet(mvarCodIvaIntra, "T") ' codigo de iva intracomunitario
    Sql = Sql & ", ctaterreten= " & DBSet(mvarCtaTerReten, "T") ' cuenta contable de retencion de facturas de terceros
    Sql = Sql & ", ctatrareten= " & DBSet(mvarCtaTraReten, "T") ' cuenta contable de retencion de facturas de transporte
    
    Sql = Sql & ", directoriotraza= " & DBSet(Replace(mvarPathTraza, "\", "\\"), "T") ' path ficheros para edicom
    
    Sql = Sql & ", sigpac= " & DBSet(Replace(mvarSigpac, "\", "\\"), "T") ' direccion de sigpac
    Sql = Sql & ", goolzoom= " & DBSet(Replace(mvarGoolZoom, "\", "\\"), "T") ' direccion de goolzoom
    
    Sql = Sql & ", ctaretensoc= " & DBSet(mvarCtaRetenSoc, "T") ' cuenta contable de retencion facturas socios
    Sql = Sql & ", ctaaportasoc= " & DBSet(mvarCtaAportaSoc, "T") ' cuenta contable de aportacion facturas socios
    Sql = Sql & ", ctabanco= " & DBSet(mvarCtaBancoSoc, "T") ' cuenta contable de banco de facturas socios
    Sql = Sql & ", codforpaposi= " & DBSet(mvarForpaPosi, "N") ' forma de pago de facturas de socio positivas
    Sql = Sql & ", codforpanega= " & DBSet(mvarForpaNega, "N") ' forma de pago de facturas de socio negativas
    
    ' almazara
    Sql = Sql & ", ctaretenalmz= " & DBSet(mvarCtaRetenAlmz, "T") ' cuenta contable de retencion almazara
    Sql = Sql & ", ctabancoalmz= " & DBSet(mvarCtaBancoAlmz, "T") ' cuenta contable de banco de almazara
    Sql = Sql & ", ctaventasalmz= " & DBSet(mvarCtaVentasAlmz, "T") ' cuenta contable de ventas almazara
    Sql = Sql & ", ctagastosalmz= " & DBSet(mvarCtaGastosAlmz, "T") ' cuenta contable de gastos almazara
    Sql = Sql & ", codforpaposalmz= " & DBSet(mvarForpaPosiAlmz, "N") ' forma de pago de facturas de almazara positivas
    Sql = Sql & ", codforpanegalmz= " & DBSet(mvarForpaNegaAlmz, "N") ' forma de pago de facturas de almazara negativas
    Sql = Sql & ", letraseriealmz= " & DBSet(mvarLetraSerieAlmz, "T") ' letra de serie de almazara facturas de almazara
    Sql = Sql & ", codgastoalmz= " & DBSet(mvarCodGastoAlmz, "N") ' codigo de gasto para liquidacion de almazara
    Sql = Sql & ", gtomoltura= " & DBSet(mvarGtoMoltura, "N") ' precio por kilo de gastos de molturacion
    Sql = Sql & ", gtoenvasado= " & DBSet(mvarGtoEnvasado, "N") ' precio por litro de gastos de envasado
    
    
    Sql = Sql & ", primfactantalmz= " & DBSet(mvarPrimFactAntAlmz, "N", "S") ' primera factura anticipos almazara
    Sql = Sql & ", ultfactantalmz= " & DBSet(mvarUltFactAntAlmz, "N", "S")   ' ultima factura anticipos almazara
    Sql = Sql & ", primfactliqalmz= " & DBSet(mvarPrimFactLiqAlmz, "N", "S") ' primera factura liquidacion almazara
    Sql = Sql & ", ultfactliqalmz= " & DBSet(mvarUltFactLiqAlmz, "N", "S")   ' ultima factura liquidacion almazara
    
    
    'adv
    Sql = Sql & ", codalmacadv= " & DBSet(mvarAlmacenADV, "N") ' codigo de almacen de adv
    Sql = Sql & ", ctabancoadv= " & DBSet(mvarCtaBancoADV, "T") ' cuenta contable de banco de adv
    Sql = Sql & ", codivaexeadv= " & DBSet(mvarCodIvaExeADV, "N") ' codigo de iva de facturas internas de adv
    
    'suministros
    Sql = Sql & ", bdariges= " & DBSet(mvarBDAriges, "T") ' bd ariges si tiene
    
    'bodega
    Sql = Sql & ", albretiradabodman= " & DBSet(mvarAlbRetiradaManual, "B")
    Sql = Sql & ", ctabancobod= " & DBSet(mvarCtaBancoBOD, "T") ' cuenta contable prevista de banco de bodega
    Sql = Sql & ", ctaventasbod= " & DBSet(mvarCtaVentasBOD, "T") ' cuenta contable de ventas de bodega
    Sql = Sql & ", bodporcenmant= " & DBSet(mvarPorcGtoMantBOD, "T") ' porcentaje de gasto de mantenimiento para liquidacion de bodega
    Sql = Sql & ", codgastobod= " & DBSet(mvarCodGastoBOD, "N") ' codigo de gasto para liquidacion de bodega
   
    Sql = Sql & ", primfactantbod= " & DBSet(mvarPrimFactAntBOD, "N", "S") ' primera factura anticipos bodega
    Sql = Sql & ", ultfactantbod= " & DBSet(mvarUltFactAntBOD, "N", "S")   ' ultima factura anticipos bodega
    Sql = Sql & ", primfactliqbod= " & DBSet(mvarPrimFactLiqBOD, "N", "S") ' primera factura liquidacion bodega
    Sql = Sql & ", ultfactliqbod= " & DBSet(mvarUltFactLiqBOD, "N", "S")   ' ultima factura liquidacion bodega
    Sql = Sql & ", porckilosbod= " & DBSet(mvarPorcKilosBOD, "N", "S") ' porcentaje de incremento de kilos neto sde entradas de bodega (mogente)
    
    'telefonia
    Sql = Sql & ", letraserietel= " & DBSet(mvarLetraSerieTel, "T") ' letra de serie de facturas de telefonia
    Sql = Sql & ", ctaventastel= " & DBSet(mvarCtaVentasTel, "T") ' cuenta contable de ventas de telefonia
   
    'gestion de nominas
    Sql = Sql & ", codalmacnomi= " & DBSet(mvarAlmacenNOMI, "N") ' codigo de almacen de gestion de nominas
    Sql = Sql & ", eurostrabdia= " & DBSet(mvarEurosTrabdiaNOMI, "N") 'euros por trabajador dia nominas
    Sql = Sql & ", porcsegso1= " & DBSet(mvarPorcSegSo1NOMI, "N") ' Porcentaje de seguridad social 1
    Sql = Sql & ", porcsegso2= " & DBSet(mvarPorcSegSo2NOMI, "N") ' Porcentaje de seguridad social 2
    Sql = Sql & ", porcirpf= " & DBSet(mvarPorcIRPFNOMI, "N") ' Porcentaje de irpf
    Sql = Sql & ", porcjornada= " & DBSet(mvarPorcJornadaNOMI, "N") ' Porcentaje de jornadas nomina
    Sql = Sql & ", nromaxjornadas= " & DBSet(mvarMaxJornadasNOMI, "N") ' Porcentaje de jornadas nomina
   
    ' pozos
    Sql = Sql & ", hastametcub1poz= " & DBSet(mvarConsumo1POZ, "N") ' hasta m3 del primer rango
    Sql = Sql & ", hastametcub2poz= " & DBSet(mvarConsumo2POZ, "N") ' hasta m3 del segundo rango
    Sql = Sql & ", precio1poz= " & DBSet(mvarPrecio1POZ, "N") ' precio del rango 1
    Sql = Sql & ", precio2poz= " & DBSet(mvarPrecio2POZ, "N") ' precio del rango 2
    Sql = Sql & ", codivapoz= " & DBSet(mvarCodIvaPOZ, "N") ' codigo de iva de pozos
    Sql = Sql & ", imporcuotapoz= " & DBSet(mvarCuotaPOZ, "N") ' importe de cuota
    Sql = Sql & ", imporderramapoz= " & DBSet(mvarDerramaPOZ, "N") ' importe derrama
    Sql = Sql & ", ctaventasconspoz= " & DBSet(mvarCtaVentasConsPOZ, "T") ' cuenta contable de ventas consumo pozos
    Sql = Sql & ", ctaventascuopoz= " & DBSet(mvarCtaVentasCuoPOZ, "T") ' cuenta contable de ventas cuotas pozos
    Sql = Sql & ", ctaventastalpoz= " & DBSet(mvarCtaVentasTalPOZ, "T") ' cuenta contable de ventas talla pozos
    Sql = Sql & ", ctaventasmtopoz= " & DBSet(mvarCtaVentasMtoPOZ, "T") ' cuenta contable de ventas mantenimiento pozos
    Sql = Sql & ", ctaventasmantapoz= " & DBSet(mvarCtaVentasMantaPOZ, "T") ' cuenta contable de ventas consumo a manta pozos
    '[Monica]21/01/2016: cuenta de recargos (escalona)
    Sql = Sql & ", ctarecargospoz= " & DBSet(mvarCtaRecargosPOZ, "T") ' cuenta contable de recargos de pozos
    
    Sql = Sql & ", codccostpoz= " & DBSet(mvarCodCCostPoz, "T") ' centro de coste pozos
    Sql = Sql & ", forpaconpoz= " & DBSet(mvarForpaConPOZ, "N") ' forma de pago de contado de pozos
    Sql = Sql & ", forparecpoz= " & DBSet(mvarForpaRecPOZ, "N") ' forma de pago de recibos de pozos
   
    Sql = Sql & ", tipolecturapoz= " & DBSet(mvarTipoLecturaPoz, "N") ' tipo de lectura de contador de Escalona (0=lectura_equipo 1=lectura_bd)
    Sql = Sql & ", codcartapoz= " & DBSet(mvarCartaPOZ, "N") ' codigo de carta de reclamacion de cartas de pozos
    Sql = Sql & ", nomzonapoz= " & DBSet(mvarNomZonaPOZ, "T") ' nombre de la zona para escalona y utxera
   
    ' transporte
    Sql = Sql & ", tratipoportes= " & DBSet(mvarTipoPortesTRA, "N") ' tipo de portes de transporte
    Sql = Sql & ", tracodtarif= " & DBSet(mvarCodTarifTRA, "N") ' codigo de tarifa de transporte de tipo 1
    Sql = Sql & ", tracodtarif2= " & DBSet(mvarCodTarifTRA2, "N") ' codigo de tarifa de transporte de tipo 2
    Sql = Sql & ", tracodgasto= " & DBSet(mvarCodGastoTRA, "N") ' codigo de gasto de transporte
    Sql = Sql & ", porcretenfactra= " & DBSet(mvarPorcReteFacTra, "N", "S")
    Sql = Sql & ", preciotra= " & DBSet(mvarPreciokgTra, "N") ' precio por kilo de transporte
    Sql = Sql & ", tratipocontador= " & DBSet(mvarTipoContadorTRA, "N") ' tipo de contador de transporte
    
    '[Monica]09/05/2013:
    Sql = Sql & ", consumominpoz= " & DBSet(mvarConsumoMinPoz, "N") 'consumo minimo de pozos
    '[Monica]11/06/2013:
    Sql = Sql & ", consumomaxpoz= " & DBSet(mvarConsumoMaxPoz, "N") 'consumo maximo de pozos
    
    '[Monica]02/12/2014:
    Sql = Sql & ", claveacceso= " & DBSet(mvarClaveAcceso, "T") ' Clave de Acceso
    
    '[Monica]21/04/2016:
    Sql = Sql & ", sqlserver= " & DBSet(mvarSqlServer, "T") ' SqlServer de Castellduc CASTELLDUCS1-O5
    Sql = Sql & ", sqlserver1= " & DBSet(mvarSqlServer1, "T") ' SqlServer 1 de Castellduc CASTELLDUCS2-O5
    
    '[Monica]17/10/2016: directorio de impresion de entradas
    Sql = Sql & ", directorioentradas= " & DBSet(Replace(mvarPathEntradas, "\", "\\"), "T") ' path ficheros de impresino de entradas
    
    
    
    
'    SQL = SQL & " WHERE codparam =" & Codigo & ";"

    conn.Execute Sql

    Modificar = 1
    Exit Function

eModificar:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Modificar = 0
End Function

