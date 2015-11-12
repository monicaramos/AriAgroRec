Attribute VB_Name = "ModVersiones"
'---------------------------------------------------------------------------------------
'VRS:2.0.0
'---------------------------------------------------------------------------------------
'Versión base de la aplicación.
'---------------------------------------------------------------------------------------
'VRS:2.0.1
'---------------------------------------------------------------------------------------
' (0) En facturacion por cliente o tarjetas si la fecha que me ponen de factura es inferior
'al ultimo periodo de facturacion no debemos dejar realizar la facturacion. Unicamente en el
'caso de que el usuario sea root aritel damos aviso y continuamos.
' (1) En el mantenimiento de articulos al intentar eliminar el articulo 11 da un error sin
'control.
' (2) El articulo no tiene que tener el codigo EAN como requerido.
' (3) Historico de facturas: incluir el boton de imprimir
' (4) Cambio de cliente: al introducir la forma de pago como un numero me trae como descripcion
'el nombre de cliente.
'---------------------------------------------------------------------------------------
'VRS:2.0.2  Modificaciones hechas para el Regaixo
'---------------------------------------------------------------------------------------
' (0) Nuevo mantenimiento de la tabla scagru/sligru utilizado en el Regaixo
' (1) En el colectivo añadimos un nuevo tipo de facturacion que es "facturacion ajena", nueva
'opcion en el combo
' (2) Nuevo mantenimiento de hco de facturas de la tabla schfacr de Regaixo donde se introducen
' las facturas de los socios de la cooperativas. Llevan distinto contador que la factura gorda
' que es la que se hace a la cooperativa
' (3) Mantenimiento de solicitud de tarjetas, se genera un fichero para una empresa de impresión
'---------------------------------------------------------------------------------------
'VRS:4.0.1  Modificaciones y petadas
'---------------------------------------------------------------------------------------
' (0) Traspaso de postes: que updatee el PVP del articulo
' (1) Mantenimiento de Articulos: el % de iva no lo muestra
' (2) Estadisticas de ventas de articulo por cliente: cuando hay desde hasta no funciona no trae
'datos cuando en realidad sí que los hay


