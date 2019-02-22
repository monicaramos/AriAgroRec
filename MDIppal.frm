VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIppal 
   BackColor       =   &H8000000C&
   Caption         =   "AriagroRec - Recolección"
   ClientHeight    =   7860
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   12120
   Icon            =   "MDIppal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cooperativas"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Socios"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Capataces"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Transportistas"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   4
            Object.Width           =   200
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Campos"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Entradas Báscula"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clasificación"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Histórico Entradas"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Precios"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas Socios"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Entrada de Lecturas"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambio Campaña"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   1
      Top             =   7275
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   1032
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3176
            MinWidth        =   3176
            Picture         =   "MDIppal.frx":6852
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2471
            MinWidth        =   2471
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4789
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5997
            MinWidth        =   5997
            Picture         =   "MDIppal.frx":7132
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "10:08"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnParametros 
      Caption         =   "&Datos Básicos"
      Index           =   1
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Datos de Empresa"
         Index           =   1
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Parámetros"
         Index           =   2
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Tipos de Movimiento"
         Index           =   3
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Tipos de Documentos"
         Index           =   4
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Usuarios"
         Index           =   6
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Cambio de Campaña"
         Index           =   7
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Salir"
         Index           =   10
      End
   End
   Begin VB.Menu mnComerGen 
      Caption         =   "Datos &Generales"
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "&Cooperativas"
         Index           =   1
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "&Secciones"
         Index           =   2
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "Situaciones &Especiales"
         Index           =   3
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "&Socios"
         Index           =   4
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "&Zonas de Cultivo"
         Index           =   6
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "&Pueblos"
         Index           =   7
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "&Partidas"
         Index           =   8
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "&Transportistas"
         Index           =   9
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "Tarifas de Transporte"
         Index           =   10
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "Capataces"
         Index           =   11
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "Situación Campos"
         Enabled         =   0   'False
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "&Incidencias"
         Index           =   13
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "Calidades"
         Enabled         =   0   'False
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "&Marco de Plantación"
         Enabled         =   0   'False
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "Tipo de Tierra"
         Enabled         =   0   'False
         Index           =   16
         Visible         =   0   'False
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "&Desarrollo vegetativo"
         Enabled         =   0   'False
         Index           =   17
         Visible         =   0   'False
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "Conceptos de &Gastos"
         Index           =   18
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "Procedencia Riego"
         Enabled         =   0   'False
         Index           =   19
         Visible         =   0   'False
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "Patrón Pie"
         Enabled         =   0   'False
         Index           =   20
         Visible         =   0   'False
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "Seguro Opción"
         Enabled         =   0   'False
         Index           =   21
         Visible         =   0   'False
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "Campos"
         Index           =   22
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "-"
         Index           =   23
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "Portes por Población - Producto"
         Index           =   24
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "Calibradores"
         Enabled         =   0   'False
         Index           =   25
         Visible         =   0   'False
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "Depósitos"
         Index           =   26
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "Coopropiedades"
         Enabled         =   0   'False
         Index           =   27
         Visible         =   0   'False
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "&Bonificaciones"
         Enabled         =   0   'False
         Index           =   28
         Visible         =   0   'False
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "&GlobalGap"
         Enabled         =   0   'False
         Index           =   29
         Visible         =   0   'False
      End
      Begin VB.Menu mnRecG_Admon 
         Caption         =   "&Variedades"
         Index           =   30
      End
   End
   Begin VB.Menu mnInformes 
      Caption         =   "&Informes"
      Begin VB.Menu mnRec_Infor2 
         Caption         =   "Datos Socios "
         Index           =   1
      End
      Begin VB.Menu mnRec_Infor 
         Caption         =   "&Socios por Sección"
         Index           =   1
      End
      Begin VB.Menu mnRec_Infor 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnRec_Infor 
         Caption         =   "Etiquetas Socios"
         Index           =   3
      End
      Begin VB.Menu mnRec_Infor 
         Caption         =   "&Cartas a Socios"
         Index           =   4
      End
      Begin VB.Menu mnRec_Infor 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnRec_Infor 
         Caption         =   "Campos/&Huertos"
         Index           =   6
      End
      Begin VB.Menu mnRec_Infor 
         Caption         =   "Superficies Cultivo / &Edad"
         Index           =   7
      End
      Begin VB.Menu mnRec_Infor 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnRec_Infor 
         Caption         =   "&Grabación Fichero AGRIWEB"
         Index           =   9
      End
      Begin VB.Menu mnRec_Infor 
         Caption         =   "&Kilos por Producto"
         Index           =   10
      End
      Begin VB.Menu mnRec_Infor 
         Caption         =   "&Informe Kilos Socio/Coop"
         Index           =   11
      End
      Begin VB.Menu mnRec_Infor 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnTomaDatos 
         Caption         =   "&Toma de Datos"
         Begin VB.Menu mnRec_Toma 
            Caption         =   "&Informe Toma de Datos"
            Index           =   1
         End
         Begin VB.Menu mnRec_Toma 
            Caption         =   "&Entrada Kilos Estimados"
            Index           =   2
         End
         Begin VB.Menu mnRec_Toma 
            Caption         =   "Informe de &Desviación"
            Index           =   3
         End
         Begin VB.Menu mnRec_Toma 
            Caption         =   "Informe &Clasificación Socio"
            Index           =   4
         End
      End
      Begin VB.Menu mnRec_DifPozos 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnRec_DifPozos 
         Caption         =   "Informe Comprobación de Datos"
         Index           =   2
      End
      Begin VB.Menu mnRec_DifPozos 
         Caption         =   "Cuentas Bancarias Erróneas"
         Index           =   3
      End
      Begin VB.Menu mnRec_InforOfi 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnRec_InforOfi 
         Caption         =   "Miembros ATRIAS"
         Index           =   2
      End
      Begin VB.Menu mnRec_InforOfi 
         Caption         =   "Registro de Fitosanitarios"
         Index           =   3
      End
      Begin VB.Menu mnRec_InforOfi 
         Caption         =   "Informe Diferencia Kilos"
         Index           =   4
      End
      Begin VB.Menu mnRec_InforOfi 
         Caption         =   "Informe Campos sin Entradas"
         Index           =   5
      End
   End
   Begin VB.Menu mnEntradas 
      Caption         =   "&Entradas"
      Begin VB.Menu mnRec_EntradasOrd 
         Caption         =   "&Ordenes Recolección"
         Index           =   1
      End
      Begin VB.Menu mnRec_EntradasOrd 
         Caption         =   "&Incidencias Ordenes"
         Index           =   2
      End
      Begin VB.Menu mnRec_EntradasOrd 
         Caption         =   "Ordenes &Emitidas"
         Index           =   3
      End
      Begin VB.Menu mnRec_EntradasOrd 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnRec_EntradasP 
         Caption         =   "Entradas &Pesada"
         Index           =   1
      End
      Begin VB.Menu mnRec_EntradasP 
         Caption         =   "&Listado Entradas Pesada"
         Index           =   2
      End
      Begin VB.Menu mnRec_EntradasP 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "&Entradas Báscula"
         Index           =   1
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "&Reimpresión Entradas Báscula"
         Index           =   2
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "&Informe de Entradas"
         Index           =   3
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "&Actualizar Entradas"
         Index           =   4
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "Clasificación Automática"
         Index           =   6
         Begin VB.Menu mnRec_Clasifica 
            Caption         =   "&Traspaso desde Calibrador"
            Index           =   1
         End
         Begin VB.Menu mnRec_Clasifica 
            Caption         =   "Traspaso TRA&ZABILIDAD"
            Index           =   2
         End
         Begin VB.Menu mnRec_Clasifica 
            Caption         =   "&Clasificación Automática"
            Index           =   3
         End
         Begin VB.Menu mnRec_Clasifica 
            Caption         =   "Control &Destrio"
            Index           =   4
         End
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "&Preclasificacion"
         Index           =   8
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "&Clasificación"
         Index           =   9
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "Informe de &Entradas"
         Index           =   10
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "Actualizar Clasificación"
         Index           =   11
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "&Histórico de Entradas"
         Index           =   13
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "&Impresión de Albaranes"
         Index           =   14
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "Informe de Clasificacion"
         Index           =   15
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "Informe Entradas"
         Index           =   16
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "&Gastos por Concepto"
         Index           =   17
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "Informe Destrios &Varios"
         Index           =   18
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "Venta &Fruta"
         Index           =   20
      End
      Begin VB.Menu mnRec_Entradas 
         Caption         =   "Histórico Entradas Clasificacion"
         Index           =   21
      End
   End
   Begin VB.Menu mnPagoSocio 
      Caption         =   "&Pago Socios"
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "&Mantenimiento Precios"
         Index           =   1
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "&Cálculo de Precios"
         Index           =   2
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "&Ventas Campo"
         Index           =   4
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "Anticipos &Gastos"
         Index           =   6
         Begin VB.Menu mnRec_AnticiposGastos 
            Caption         =   "&Previsión de Pago"
            Index           =   1
         End
         Begin VB.Menu mnRec_AnticiposGastos 
            Caption         =   "&Factura Gastos"
            Index           =   2
         End
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "Anticipos Genéricos/Retirada"
         Index           =   8
         Begin VB.Menu mnRec_AnticiposGene 
            Caption         =   "Previsión de Pago"
            Index           =   1
         End
         Begin VB.Menu mnRec_AnticiposGene 
            Caption         =   "&Factura Anticipo"
            Index           =   2
         End
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "&Anticipos"
         Index           =   10
         Begin VB.Menu mnRec_Anticipos 
            Caption         =   "&Previsión de Pago"
            Index           =   1
         End
         Begin VB.Menu mnRec_Anticipos 
            Caption         =   "&Informe de Anticipos"
            Index           =   2
         End
         Begin VB.Menu mnRec_Anticipos 
            Caption         =   "&Factura de Anticipos"
            Index           =   3
         End
         Begin VB.Menu mnRec_Anticipos 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnRec_Anticipos 
            Caption         =   "&Deshacer Facturación"
            Index           =   5
         End
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "&Liquidaciones"
         Index           =   12
         Begin VB.Menu mnRec_LiquFO 
            Caption         =   "Cálculo &Aportacion FO"
            Index           =   1
         End
         Begin VB.Menu mnRec_LiquFO 
            Caption         =   "&Conceptos de Incremento"
            Index           =   2
         End
         Begin VB.Menu mnRec_Liqudaciones 
            Caption         =   "&Previsión de Pago"
            Index           =   1
         End
         Begin VB.Menu mnRec_Liqudaciones 
            Caption         =   "&Informe de Liquidaciones"
            Index           =   2
         End
         Begin VB.Menu mnRec_Liqudaciones 
            Caption         =   "&Factura de Liquidaciones"
            Index           =   3
         End
         Begin VB.Menu mnRec_Liqudaciones 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnRec_Liqudaciones 
            Caption         =   "&Deshacer Facturación"
            Index           =   5
         End
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "Liquidación &Industria"
         Index           =   14
         Begin VB.Menu mnRec_LiquIndustria 
            Caption         =   "&Previsión de Pago"
            Index           =   1
         End
         Begin VB.Menu mnRec_LiquIndustria 
            Caption         =   "&Informe de Liquidación"
            Index           =   2
         End
         Begin VB.Menu mnRec_LiquIndustria 
            Caption         =   "&Factura de Industria"
            Index           =   3
         End
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "&Factura Rectificativa"
         Index           =   18
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu mnRec_PagoSocios 
         Caption         =   "Facturación &Contratos"
         Index           =   20
      End
   End
   Begin VB.Menu mnFacturasSocios 
      Caption         =   "&Facturas"
      Begin VB.Menu mnRec_FactSocios 
         Caption         =   "&Facturas Socios"
         Index           =   1
         Begin VB.Menu mnRec_FacturasSocios 
            Caption         =   "&Facturas Socios"
            Index           =   1
         End
         Begin VB.Menu mnRec_FacturasSocios 
            Caption         =   "&Reimpresión de Facturas"
            Index           =   2
         End
         Begin VB.Menu mnRec_FacturasSocios 
            Caption         =   "&Enviar Facturas por email"
            Index           =   3
         End
         Begin VB.Menu mnRec_FacturasSocios 
            Caption         =   "Facturación &Web/Electrónica"
            Index           =   4
         End
         Begin VB.Menu mnRec_FacturasSocios 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnRec_FacturasSocios 
            Caption         =   "Integración &Contable"
            Index           =   6
         End
         Begin VB.Menu mnRec_FacturasSocios 
            Caption         =   "Integración &Aridoc"
            Index           =   7
         End
         Begin VB.Menu mnRec_FacturasSocios 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnRec_FacturasSocios 
            Caption         =   "&Informe de Resultados"
            Index           =   9
         End
         Begin VB.Menu mnRec_FacturasSocios 
            Caption         =   "&Listado Retenciones"
            Index           =   10
         End
         Begin VB.Menu mnRec_FacturasSocios 
            Caption         =   "&Grabacion Modelo 190"
            Index           =   11
         End
         Begin VB.Menu mnRec_FacturasSocios 
            Caption         =   "&Grabacion Modelo 346"
            Index           =   12
         End
         Begin VB.Menu mnRec_FacturasSocios 
            Caption         =   "-"
            Index           =   13
         End
         Begin VB.Menu mnRec_FacturasSocios 
            Caption         =   "Anticipos Pdtes Descontar"
            Index           =   14
         End
      End
      Begin VB.Menu mnRec_FactSocios 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnRec_FactSocios 
         Caption         =   "Facturas &Terceros"
         Index           =   3
         Begin VB.Menu mnRec_FacturasTerceros 
            Caption         =   "&Recepción Facturas"
            Index           =   1
         End
         Begin VB.Menu mnRec_FacturasTerceros 
            Caption         =   "&Histórico Albarán/Factura"
            Index           =   2
         End
         Begin VB.Menu mnRec_FacturasTerceros 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnRec_FacturasTerceros 
            Caption         =   "&Integración Contable"
            Index           =   4
         End
         Begin VB.Menu mnRec_FacturasTerceros 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnRec_FacturasTerceros 
            Caption         =   "Informe de &Albaranes "
            Index           =   6
         End
         Begin VB.Menu mnRec_FacturasTerceros 
            Caption         =   "Informe de &Retenciones"
            Index           =   7
         End
      End
      Begin VB.Menu mnRec_FactSocios 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnRec_FactSocios 
         Caption         =   "Facturas &Varias Cliente"
         Index           =   5
         Begin VB.Menu mnRec_FacturasVarias 
            Caption         =   "&Contadores"
            Index           =   1
         End
         Begin VB.Menu mnRec_FacturasVarias 
            Caption         =   "Concep&tos"
            Index           =   2
         End
         Begin VB.Menu mnRec_FacturasVarias 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnRec_FacturasVarias 
            Caption         =   "Facturas &Varias"
            Index           =   4
         End
         Begin VB.Menu mnRec_FacturasVarias 
            Caption         =   "&Reimpresion de Facturas"
            Index           =   5
         End
         Begin VB.Menu mnRec_FacturasVarias 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnRec_FacturasVarias 
            Caption         =   "&Integración Contable"
            Index           =   7
         End
         Begin VB.Menu mnRec_FacturasVarias 
            Caption         =   "Integración &Aridoc"
            Index           =   8
         End
         Begin VB.Menu mnRec_FacturasVarias 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnRec_FacturasVarias 
            Caption         =   "&Diario de Facturación"
            Index           =   10
         End
      End
      Begin VB.Menu mnRec_FactSocios 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnRec_FactSocios 
         Caption         =   "Facturas Varias &Proveedor"
         Index           =   7
         Begin VB.Menu mnRec_FacturasVariasPr 
            Caption         =   "Facturas Varias"
            Index           =   1
         End
         Begin VB.Menu mnRec_FacturasVariasPr 
            Caption         =   "&Reimpresion de Facturas"
            Index           =   2
         End
         Begin VB.Menu mnRec_FacturasVariasPr 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnRec_FacturasVariasPr 
            Caption         =   "&Integración Contable"
            Index           =   4
         End
         Begin VB.Menu mnRec_FacturasVariasPr 
            Caption         =   "Integración &Aridoc"
            Index           =   5
         End
         Begin VB.Menu mnRec_FacturasVariasPr 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnRec_FacturasVariasPr 
            Caption         =   "&Diario de Facturación"
            Index           =   7
         End
      End
   End
   Begin VB.Menu mnTransporte 
      Caption         =   "&Transporte"
      Begin VB.Menu mnRec_Transporte1 
         Caption         =   "Informe de Transporte"
         Index           =   1
      End
      Begin VB.Menu mnRec_Transporte 
         Caption         =   "&Facturación  Transportista"
         Index           =   1
      End
      Begin VB.Menu mnRec_Transporte 
         Caption         =   "&Reimpresión de Facturas"
         Index           =   2
      End
      Begin VB.Menu mnRec_Transporte 
         Caption         =   "&Histórico de Facturas"
         Index           =   3
      End
      Begin VB.Menu mnRec_Transporte 
         Caption         =   "Facturación &Socio"
         Index           =   4
      End
      Begin VB.Menu mnRec_Transporte 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnRec_Transporte 
         Caption         =   "Integración &Contable"
         Index           =   6
      End
      Begin VB.Menu mnRec_Transporte 
         Caption         =   "&Integración Aridoc"
         Index           =   7
      End
      Begin VB.Menu mnRec_Transporte 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnRec_Transporte 
         Caption         =   "&Envases Retornables"
         Index           =   9
      End
   End
   Begin VB.Menu mnGestPrenom 
      Caption         =   "&Prenómina"
      Begin VB.Menu mnP_PreNomCateg 
         Caption         =   "&Categorias"
         Index           =   1
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "&Salarios"
         Index           =   1
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "&Trabajadores"
         Index           =   2
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "&Cuadrillas"
         Index           =   3
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "Conceptos &Gastos"
         Index           =   4
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "Tarifas &ETT"
         Index           =   5
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "&Partes de Trabajo"
         Index           =   7
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "Pago de Partes"
         Index           =   8
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "&Horas Destajo"
         Index           =   10
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "&Anticipos"
         Index           =   12
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "&Entrada de Horas"
         Index           =   13
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "&Impresión Recibos"
         Index           =   14
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "&Pago Recibos"
         Index           =   16
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "&Integración Aridoc"
         Index           =   17
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "-"
         Index           =   18
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "Informe &Mensual Horas"
         Index           =   19
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "&Entrada Horas ETT"
         Index           =   20
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "&Diferencias Asesoria"
         Index           =   21
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "Informe Mes &Asesoria"
         Index           =   22
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "Rendimiento por Capataz"
         Index           =   23
      End
      Begin VB.Menu mnP_PreNominas 
         Caption         =   "Trabajadores Activos"
         Index           =   24
      End
   End
   Begin VB.Menu mnTrazabilidad 
      Caption         =   "Trazabilidad"
      Begin VB.Menu mnRec_Traza1 
         Caption         =   "Carga &Automática"
         Index           =   1
      End
      Begin VB.Menu mnRec_Trazabilidad 
         Caption         =   "&Tipos"
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnRec_Trazabilidad 
         Caption         =   "&Areas"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnRec_Trazabilidad 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnRec_Trazabilidad 
         Caption         =   "&Detalles Palets en Entradas"
         Index           =   4
      End
      Begin VB.Menu mnRec_Trazabilidad 
         Caption         =   "&Cargas Linea Confección"
         Index           =   5
      End
      Begin VB.Menu mnRec_Trazabilidad 
         Caption         =   "Cargas Lineas &Fecha/Producto"
         Index           =   6
      End
      Begin VB.Menu mnRec_Trazabilidad 
         Caption         =   "&Origenes Palet Confeccionado"
         Index           =   7
      End
      Begin VB.Menu mnRec_Trazabilidad 
         Caption         =   "Destino &Albaranes de Venta"
         Index           =   8
      End
      Begin VB.Menu mnRec_Trazabilidad 
         Caption         =   "&Listado de Stock"
         Index           =   9
      End
      Begin VB.Menu mnRec_Trazabilidad 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnRec_Trazabilidad 
         Caption         =   "&Manejo de Palets"
         Index           =   11
      End
      Begin VB.Menu mnRec_Trazabilidad 
         Caption         =   "Manejo Cargas Confección"
         Index           =   12
      End
      Begin VB.Menu mnRec_Trazabilidad 
         Caption         =   "Manejo Asignacion Albaranes"
         Index           =   13
      End
   End
   Begin VB.Menu mnAlmazara 
      Caption         =   "&Almazara"
      Begin VB.Menu mnRec_AlmTrasCampos 
         Caption         =   "Traspaso &Campos"
         Index           =   1
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "&Traspaso de Almazara"
         Index           =   1
         Begin VB.Menu mnRec_AlmazaraTras 
            Caption         =   "&Traspaso Facturas"
            Index           =   1
         End
         Begin VB.Menu mnRec_AlmazaraTras 
            Caption         =   "&Histórico Facturas"
            Index           =   2
         End
         Begin VB.Menu mnRec_AlmazaraTras 
            Caption         =   "&Reimpresión Facturas"
            Index           =   3
         End
         Begin VB.Menu mnRec_AlmazaraTras 
            Caption         =   "&Integración Contable"
            Index           =   4
         End
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "Traspaso &Báscula"
         Index           =   3
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "&Entradas Almazara"
         Index           =   4
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "Reparto &Gastos Liquidación"
         Index           =   5
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "&Mantenimiento de Precios"
         Index           =   6
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "&Informe de Entradas"
         Index           =   7
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "Extracto Entradas &Socio/Variedad"
         Index           =   8
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "Traspaso &de Rendimiento"
         Index           =   10
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "Histórico de Rendimiento"
         Index           =   11
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "&Retirada Aceite"
         Index           =   13
         Begin VB.Menu mnRec_Almaz 
            Caption         =   "&Albaranes Retirada"
            Index           =   1
         End
         Begin VB.Menu mnRec_Almaz 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnRec_Almaz 
            Caption         =   "Previsión Facturación"
            Index           =   3
         End
         Begin VB.Menu mnRec_Almaz 
            Caption         =   "&Facturación"
            Index           =   4
         End
         Begin VB.Menu mnRec_Almaz 
            Caption         =   "&Reimpresión de Facturas"
            Index           =   5
         End
         Begin VB.Menu mnRec_Almaz 
            Caption         =   "&Histórico Facturas"
            Index           =   6
         End
         Begin VB.Menu mnRec_Almaz 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnRec_Almaz 
            Caption         =   "&Integración Contable"
            Index           =   8
         End
         Begin VB.Menu mnRec_Almaz 
            Caption         =   "&Integración Aridoc"
            Index           =   9
         End
         Begin VB.Menu mnRec_Almaz 
            Caption         =   "-"
            Index           =   10
         End
         Begin VB.Menu mnRec_Almaz 
            Caption         =   "&Diferencia Consumo/Producido"
            Index           =   11
         End
         Begin VB.Menu mnRec_Almaz 
            Caption         =   "Diario de Facturación"
            Index           =   12
         End
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "&Anticipos"
         Index           =   15
         Begin VB.Menu mnRec_AlmzAnticipos 
            Caption         =   "&Previsión de Pago"
            Index           =   1
         End
         Begin VB.Menu mnRec_AlmzAnticipos 
            Caption         =   "&Factura Anticipo"
            Index           =   2
         End
         Begin VB.Menu mnRec_AlmzAnticipos 
            Caption         =   "&Deshacer Facturación"
            Index           =   3
         End
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnRec_Almazara 
         Caption         =   "&Liquidación"
         Index           =   17
         Begin VB.Menu mnRec_AlmzLiquidacion 
            Caption         =   "&Previsión de Pago"
            Index           =   1
         End
         Begin VB.Menu mnRec_AlmzLiquidacion 
            Caption         =   "&Factura Liquidación"
            Index           =   2
         End
         Begin VB.Menu mnRec_AlmzLiquidacion 
            Caption         =   "&Deshacer Facturación"
            Index           =   3
         End
         Begin VB.Menu mnRec_AlmzLiquidacion 
            Caption         =   "Informe Liquidación"
            Index           =   4
         End
      End
   End
   Begin VB.Menu mnADV 
      Caption         =   "A&DV"
      Begin VB.Menu mnRec_ADV 
         Caption         =   "Fami&lias de Artículos"
         Index           =   1
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "&Materias Activas"
         Index           =   2
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "&Artículos"
         Index           =   3
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "&Tratamientos"
         Index           =   4
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "&Partes de Trabajo"
         Index           =   5
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "Reimpresión Partes Trabajo"
         Index           =   6
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "&Previsión Facturación"
         Index           =   8
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "&Facturación"
         Index           =   9
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "Reimpresión Facturas"
         Index           =   10
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "Histórico Parte/Factura"
         Index           =   11
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "&Integración Contable"
         Index           =   13
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "Integración Aridoc"
         Index           =   14
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "Informe de Rendimiento"
         Index           =   16
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnRec_ADV 
         Caption         =   "&Pago de Partes Trabajador"
         Index           =   18
      End
   End
   Begin VB.Menu mnBodega 
      Caption         =   "&Bodega"
      Begin VB.Menu mnRec_Bodega 
         Caption         =   "&Entradas"
         Index           =   1
      End
      Begin VB.Menu mnRec_Bodega 
         Caption         =   "&Bonificaciones"
         Index           =   2
      End
      Begin VB.Menu mnRec_Bodega 
         Caption         =   "Reparto &Gastos Liquidación"
         Index           =   3
      End
      Begin VB.Menu mnRec_Bodega 
         Caption         =   "&Mantenimiento de Precios"
         Index           =   4
      End
      Begin VB.Menu mnRec_Bodega 
         Caption         =   "&Informe de Entradas"
         Index           =   5
      End
      Begin VB.Menu mnRec_Bodega 
         Caption         =   "Extracto Entradas &Socio/Variedad"
         Index           =   6
      End
      Begin VB.Menu mnRec_Bodega 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnRec_Bodega 
         Caption         =   "&Retirada Vino"
         Index           =   8
         Begin VB.Menu mnRec_Bod1 
            Caption         =   "&Albaranes Retirada"
            Index           =   1
         End
         Begin VB.Menu mnRec_Bod1 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnRec_Bod1 
            Caption         =   "Previsión Facturación"
            Index           =   3
         End
         Begin VB.Menu mnRec_Bod1 
            Caption         =   "&Facturación"
            Index           =   4
         End
         Begin VB.Menu mnRec_Bod1 
            Caption         =   "&Reimpresión de Facturas"
            Index           =   5
         End
         Begin VB.Menu mnRec_Bod1 
            Caption         =   "&Histórico Facturas"
            Index           =   6
         End
         Begin VB.Menu mnRec_Bod1 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnRec_Bod1 
            Caption         =   "&Integración Contable"
            Index           =   8
         End
         Begin VB.Menu mnRec_Bod1 
            Caption         =   "&Integración Aridoc"
            Index           =   9
         End
         Begin VB.Menu mnRec_Bod1 
            Caption         =   "-"
            Index           =   10
         End
         Begin VB.Menu mnRec_Bod1 
            Caption         =   "&Listado de Consumo"
            Index           =   11
         End
      End
      Begin VB.Menu mnRec_Bodega 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnRec_Bodega 
         Caption         =   "&Anticipos"
         Index           =   10
         Begin VB.Menu mnRec_BodAnticipos 
            Caption         =   "&Previsión de Pago"
            Index           =   1
         End
         Begin VB.Menu mnRec_BodAnticipos 
            Caption         =   "&Factura Anticipo"
            Index           =   2
         End
         Begin VB.Menu mnRec_BodAnticipos 
            Caption         =   "&Deshacer Facturación"
            Index           =   3
         End
      End
      Begin VB.Menu mnRec_Bodega 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnRec_Bodega 
         Caption         =   "&Liquidación"
         Index           =   12
         Begin VB.Menu mnRec_BodLiquidacion 
            Caption         =   "&Previsión de Pago"
            Index           =   1
         End
         Begin VB.Menu mnRec_BodLiquidacion 
            Caption         =   "&Factura Liquidación"
            Index           =   2
         End
         Begin VB.Menu mnRec_BodLiquidacion 
            Caption         =   "&Deshacer Facturación"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnPozos 
      Caption         =   "&Pozos"
      Begin VB.Menu mnRec_Monast 
         Caption         =   "Pueblos"
         Index           =   1
      End
      Begin VB.Menu mnRec_Monast 
         Caption         =   "Calles"
         Index           =   2
      End
      Begin VB.Menu mnRec_Monast 
         Caption         =   "Socios"
         Index           =   3
      End
      Begin VB.Menu mnRec_Monast 
         Caption         =   "Propiedades"
         Index           =   4
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Tipos de &Pozos"
         Index           =   1
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "&Hidrantes/Contadores"
         Index           =   2
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "&Toma de Lectura Contador"
         Index           =   4
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "&Introducción de Lecturas"
         Index           =   5
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "&Listado de Comprobación"
         Index           =   6
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "&Generación Recibos Consumo"
         Index           =   8
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Generación &Recibos Mantenimiento"
         Index           =   9
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Generación Recibos &Contadores"
         Index           =   10
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Tickets de Consumo Manta"
         Index           =   12
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Informe Recibos Riego a &Manta"
         Index           =   13
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Carta de Tallas a &Socios"
         Index           =   15
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Generación Recibos &Talla"
         Index           =   16
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Cálculo Bonificación Talla"
         Index           =   17
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "&Histórico de Recibos"
         Index           =   18
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "&Reimpresión Recibos"
         Index           =   20
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "&Carta Reclamación Pago"
         Index           =   21
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Generación Recibos con Recargo"
         Index           =   22
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "-"
         Index           =   23
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "&Integración Contable"
         Index           =   24
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Integración &Aridoc"
         Index           =   25
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "-"
         Index           =   26
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Informe &Facturado por Hidrante"
         Index           =   27
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Informe Recibos Pendientes &Cobro"
         Index           =   28
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Recibos Consumo Pdtes Cobro"
         Index           =   29
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Etiquetas de &Contadores"
         Index           =   30
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "-"
         Index           =   31
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "&Rectificación de Lecturas"
         Index           =   32
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Importacion Lecturas"
         Index           =   33
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "-"
         Index           =   34
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Introducción Lecturas"
         Index           =   35
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "-"
         Index           =   36
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Listado de Comprobacion"
         Index           =   37
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "-"
         Index           =   38
      End
      Begin VB.Menu mnRec_Pozos 
         Caption         =   "Exportacion Lecturas"
         Index           =   39
      End
   End
   Begin VB.Menu mnAportaciones 
      Caption         =   "&Aportaciones"
      Begin VB.Menu mnE_Aport 
         Caption         =   "&Tipos"
         Index           =   1
      End
      Begin VB.Menu mnE_Aport 
         Caption         =   "&Aportaciones"
         Index           =   2
      End
      Begin VB.Menu mnE_Aport 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnE_Aport 
         Caption         =   "&Informe de Aportaciones"
         Index           =   4
      End
      Begin VB.Menu mnE_Aport 
         Caption         =   "&Regularización de Aportaciones"
         Index           =   5
      End
      Begin VB.Menu mnE_Aport 
         Caption         =   "&Certificado de Aportaciones"
         Index           =   6
      End
      Begin VB.Menu mnE_Aport 
         Caption         =   "Aportación &Obligatoria"
         Index           =   7
      End
      Begin VB.Menu mnE_Aport 
         Caption         =   "Integr&ación Tesoreria"
         Index           =   8
      End
      Begin VB.Menu mnE_Aport 
         Caption         =   "&Devolución Aportaciones"
         Index           =   9
      End
   End
   Begin VB.Menu mnUtil 
      Caption         =   "&Utilidades"
      Begin VB.Menu mnE_Util 
         Caption         =   "Revisión de caracteres en Multibase"
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "&Traspaso Facturas a Cooperativas"
         Index           =   3
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "Traspaso &ROPAS"
         Index           =   4
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "&Actualizar/Insertar Socios Suministros"
         Index           =   5
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "&Traspaso Datos a Almazara"
         Index           =   6
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "Traspaso &Facturas Telefonía"
         Index           =   8
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "&Facturas Telefonía"
         Index           =   9
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "&Integración Contable"
         Index           =   10
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "Copia Seguridad"
         Index           =   12
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "Reparto Albaranes"
         Index           =   13
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "Traspaso Entradas &Traza"
         Index           =   15
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "Traspaso a Tra&zabilidad"
         Index           =   16
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "Traspaso Albaranes Retirada"
         Index           =   17
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "-"
         Index           =   18
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "Accciones Realizadas"
         Index           =   19
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "Comunicación de Datos"
         Index           =   21
      End
   End
   Begin VB.Menu mnSoporte 
      Caption         =   "&Soporte"
      Begin VB.Menu mnE_Soporte1 
         Caption         =   "&Web Soporte"
      End
      Begin VB.Menu mnp_Barra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnE_Soporte2 
         Caption         =   "&Acerca de"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnE_Soporte3 
         Caption         =   "&Calculo Aportaciones"
      End
      Begin VB.Menu mnE_Soporte4 
         Caption         =   "Ejecucion Prg"
      End
   End
End
Attribute VB_Name = "MDIppal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private PrimeraVez As Boolean
Dim TieneEditorDeMenus As Boolean

Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String, ByVal op As Integer, ByVal tam As Integer)
    Dim i As Integer
    Dim tRes As ResType, iCount As Integer
        
    opcio = op
    tamany = tam
    ghmodule = LoadLibraryEx(sLibraryFilePath, 0, DONT_RESOLVE_DLL_REFERENCES)

    If ghmodule = 0 Then
        MsgBox "Invalid library file.", vbCritical
        Exit Sub
    End If
        
    For tRes = RT_FIRST To RT_LAST
        DoEvents
        EnumResourceNames ghmodule, tRes, AddressOf EnumResNameProc, 0
    Next
    FreeLibrary ghmodule
             
End Sub


Private Sub MDIForm_Activate()
'Dim cad As String

    If PrimeraVez Then
        PrimeraVez = False
'        frmMensaje.pTitulo = "Últimas modificaciones.         14/03/06"
'
''        cad = cad & "-----------------------------------------------------------------------------------------------" & vbCrLf
''        cad = cad & "Para actualizar el estado de un presupuesto desde la pantalla de ventas sin "
''        cad = cad & "entrar en la pantalla de modificación de presupuesto, seleccionar la línea del "
''        cad = cad & "presupuesto a modificar, pulsar botón izquierdo del ratón, se despliega un menu "
''        cad = cad & "con los posibles estados y se selecciona el nuevo estado." & vbCrLf & vbCrLf
'
''        cad = cad & "- Imprimir informes de subcontratación." & vbCrLf
''        cad = cad & "- Ventas pendientes." & vbCrLf
''        cad = cad & "-------------------------------------------------------------------------" & vbCrLf & vbCrLf
'
'        cad = cad & "- Mantenimiento de No Conformidades y lineas de acciones y reclamaciones." & vbCrLf & vbCrLf
'        cad = cad & "- Informes:" & vbCrLf
'        cad = cad & "     Comunicación con cliente." & vbCrLf
'        cad = cad & "     Confirmación de servicios." & vbCrLf
'        cad = cad & "     No conformidad." & vbCrLf
'        cad = cad & "     Reclamación." & vbCrLf & vbCrLf
'
'
'        frmMensaje.pValor = cad
'        frmMensaje.Show vbModal
    End If
End Sub

Private Sub MDIForm_Load()
Dim Cad As String

    PrimeraVez = True
    CargarImagen
    PonerDatosFormulario

    
    If vParam Is Nothing Then
        Caption = "AriAgro - Recolección   " & " v." & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & " FALTA CONFIGURAR"
    Else
        Caption = "AriAgro - Recolección   " & " v." & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & vParam.NombreEmpresa & Cad & _
                  " - Campaña: " & vParam.FecIniCam & " - " & vParam.FecFinCam & "   -  Usuario: " & vUsu.Nombre
    End If

    GetIconsFromLibrary App.Path & "\iconos.dll", 1, 48
    


    ' *** per als iconos XP ***
    GetIconsFromLibrary App.Path & "\iconos.dll", 1, 24
    GetIconsFromLibrary App.Path & "\iconos_BN.dll", 2, 24
    GetIconsFromLibrary App.Path & "\iconos_OM.dll", 3, 24
    
    GetIconsFromLibrary App.Path & "\iconosAriagroRec.dll", 4, 24
  
    'CARGAR LA TOOLBAR DEL FORM PRINCIPAL
    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListPpal

        .Buttons(1).Image = 34  'Cooperativas
        .Buttons(2).Image = 2   'Socios
        .Buttons(3).Image = 32  'Capataces
        .Buttons(4).Image = 7   'Transportistas
        ' el 5 separador
        .Buttons(6).Image = 30  'Campos
        ' el 7 separador
        .Buttons(8).Image = 35  'Entradas Báscula
        .Buttons(9).Image = 5   'Clasificacion de los campos
        .Buttons(10).Image = 25   'Historico de entradas
        ' el 11 separador
        .Buttons(12).Image = 9   'Mantenimiento de precios
        .Buttons(13).Image = 23   'Facturas de socios
        '
        .Buttons(14).Image = 24   'entrada de lecturas
        
        .Buttons(15).Image = 37   'Cambio de campaña
        ' el 16 separador
        .Buttons(17).Image = 1   'Salir
    End With
    
    
    
    GetIconsFromLibrary App.Path & "\iconos.dll", 1, 16
    GetIconsFromLibrary App.Path & "\iconos_BN.dll", 2, 16
    GetIconsFromLibrary App.Path & "\iconos_OM.dll", 3, 16

    GetIconsFromLibrary App.Path & "\iconosAriagroRec.dll", 4, 16

    LeerEditorMenus

    
    PonerDatosFormulario
    
'    Stop

'[Monica]22/02/2019: quitamos lo de indefa
'
'    '[Monica]08/10/2015: solo en el caso de escalona mandamos los datos a indefa
'    If vParamAplic.Cooperativa = 10 Then
'        If Dir("c:\indefa", vbDirectory) <> "" Then
'            LanzaVisorMimeDocumento Me.hWnd, "c:\indefa\ftpINDEFA.bat"
'        Else
'            If MsgBox("No existe el directorio del traspaso. ¿ Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then End
'        End If
'    End If
    
    
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    AccionesCerrar
    End
End Sub

Private Sub mnE_Aport_Click(Index As Integer)
    SubmnC_RecoleccionG_Aport_Click (Index)
End Sub

Private Sub mnE_Soporte3_Click()
    frmListAnticipos.Opcionlistado = 18
    frmListAnticipos.Show vbModal
End Sub

' Me dejo un punto donde poder programar.
Private Sub mnE_Soporte4_Click()
    
'    frmContrRecFactPre.Albaranes = "177,178"
'    frmContrRecFactPre.Show vbModal
    
'    frmVARIOS.Show vbModal
'    AbrirListadoAPOR (13)
    
'     frmPOZMantaTickets.Show vbModal
'     frmPOZMantaAux.Show vbModal
    frmVARIOS.Show vbModal
    
'    MsgBox "Hola" & vbCrLf & "Qué tal?", vbQuestion + vbYesNo + vbDefaultButton2
'    frmLotApport.Show vbModal
    
End Sub

Private Sub mnP_PreNomCateg_Click(Index As Integer)
    SubmnP_PreNominasCateg_click (Index)
End Sub

Private Sub mnP_PreNominas_Click(Index As Integer)
    SubmnP_PreNominas_click (Index)
End Sub

Private Sub mnRec_ADV_Click(Index As Integer)
    SubmnC_RecoleccionG_ADV_Click (Index)
End Sub

Private Sub mnRec_Almaz_Click(Index As Integer)
    SubmnC_RecoleccionG_Almz1_Click (Index)
End Sub

Private Sub mnRec_Almazara_Click(Index As Integer)
    SubmnC_RecoleccionG_Almz_Click (Index)
End Sub

Private Sub mnRec_AlmTrasCampos_Click(Index As Integer)
    SubmnC_RecoleccionG_Almz0_Click (Index)
End Sub

Private Sub mnRec_AlmzAnticipos_Click(Index As Integer)
    SubmnC_RecoleccionG_AlmzAnticipos_Click (Index)
End Sub

Private Sub mnRec_AlmzLiquidacion_Click(Index As Integer)
    SubmnC_RecoleccionG_AlmzLiquidacion_Click (Index)
End Sub

Private Sub mnRec_AnticiposGene_Click(Index As Integer)
    SubmnC_RecoleccionG_AnticiposGene_Click (Index)
End Sub

Private Sub mnRec_Bod1_Click(Index As Integer)
    SubmnC_RecoleccionG_Bod1_Click (Index)
End Sub

Private Sub mnRec_BodAnticipos_Click(Index As Integer)
    SubmnC_RecoleccionG_BodAnticipos_Click (Index)
End Sub

Private Sub mnRec_BodLiquidacion_Click(Index As Integer)
    SubmnC_RecoleccionG_BodLiquidacion_Click (Index)
End Sub


Private Sub mnRec_Bodega_Click(Index As Integer)
    SubmnC_RecoleccionG_Bodega_Click (Index)
End Sub

Private Sub mnRec_AlmazaraTras_Click(Index As Integer)
    SubmnC_RecoleccionG_Almazara_Click (Index)
End Sub

Private Sub mnRec_Anticipos_Click(Index As Integer)
    SubmnC_RecoleccionG_Anticipos_Click (Index)
End Sub

Private Sub mnRec_AnticiposGastos_Click(Index As Integer)
    SubmnC_RecoleccionG_AnticiposGastos_Click (Index)
End Sub

Private Sub mnRec_Clasifica_Click(Index As Integer)
    SubmnC_RecoleccionG_Clasifica_Click (Index)
End Sub

Private Sub mnRec_DifPozos_Click(Index As Integer)
    SubmnC_RecoleccionPOZOS_Click (Index)
End Sub

Private Sub mnRec_Entradas_Click(Index As Integer)
    SubmnC_RecoleccionG_Entradas_Click (Index)
End Sub

Private Sub mnRec_EntradasOrd_Click(Index As Integer)
    SubmnC_RecoleccionG_EntradasOrd_Click (Index)
End Sub

Private Sub mnRec_EntradasP_Click(Index As Integer)
    SubmnC_RecoleccionG_EntradasP_Click (Index)
End Sub

Private Sub mnRec_FacturasSocios_click(Index As Integer)
    SubmnC_RecoleccionG_FacturasSocios_Click (Index)
End Sub

Private Sub mnRec_FacturasTerceros_click(Index As Integer)
    SubmnC_RecoleccionG_FacturasTerceros_Click (Index)
End Sub


Private Sub mnRec_FacturasVarias_Click(Index As Integer)
    SubmnC_RecoleccionG_FacturasVarias_Click (Index)
End Sub

Private Sub mnRec_FacturasVariasPr_Click(Index As Integer)
    SubmnC_RecoleccionG_FacturasVariasProv_Click (Index)
End Sub

Private Sub mnRec_Infor_Click(Index As Integer)
    SubmnC_RecoleccionG_Infor_Click (Index)
End Sub

Private Sub mnRec_Infor2_Click(Index As Integer)
    SubmnC_RecoleccionG_Infor2_Click (Index)
End Sub

Private Sub mnRec_InforOfi_Click(Index As Integer)
    SubmnC_RecoleccionG_InforOfi_Click (Index)
End Sub

Private Sub mnRec_Liqudaciones_Click(Index As Integer)
    SubmnC_RecoleccionG_Liquidaciones_Click (Index)
End Sub

Private Sub mnRec_LiquFO_Click(Index As Integer)
    SubmnC_RecoleccionG_LiquidFO_Click (Index)
End Sub

Private Sub mnRec_LiquIndustria_Click(Index As Integer)
    SubmnC_RecoleccionG_LiquIndustria_Click (Index)
End Sub

Private Sub mnRec_Monast_Click(Index As Integer)
    SubmnC_RecoleccionG_PozosMonast_Click (Index)
End Sub


Private Sub mnRec_PagoSocios_Click(Index As Integer)
    SubmnC_RecoleccionG_PagoSocios_Click (Index)
End Sub

Private Sub mnRec_Pozos_Click(Index As Integer)
    SubmnC_RecoleccionG_Pozos_Click (Index)
End Sub

Private Sub mnRec_Toma_Click(Index As Integer)
    SubmnC_RecoleccionG_TomaDatos_Click (Index)
End Sub

Private Sub mnRec_Transporte_Click(Index As Integer)
    SubmnC_RecoleccionG_Transporte_Click (Index)
End Sub

Private Sub mnRec_Transporte1_Click(Index As Integer)
    SubmnC_RecoleccionG_Transporte1_Click (Index)
End Sub

Private Sub mnRec_Traza1_Click(Index As Integer)
     SubmnC_RecoleccionG_Traza1_Click (Index)
End Sub

Private Sub mnRec_Trazabilidad_Click(Index As Integer)
    SubmnC_RecoleccionG_Trazabilidad_Click (Index)
End Sub

Private Sub mnRecG_Admon_Click(Index As Integer)
    SubmnC_RecoleccionG_Admon_Click (Index)
End Sub

Private Sub mnE_Soporte1_Click()
    Screen.MousePointer = vbHourglass
    LanzaHome "websoporte"
    espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnE_Util_Click(Index As Integer)
    SubmnE_Util_Click (Index)
End Sub


Private Sub mnE_Soporte2_Click()
    frmMensaje.OpcionMensaje = 6
    frmMensaje.Show vbModal
End Sub

Private Sub mnP_Generales_Click(Index As Integer)
    If Index = 7 Then
        mnCambioEmpresa_Click
    Else
        SubmnP_Generales_Click (Index)
    End If
End Sub

Private Sub mnP_Salir1_Click()
'    Unload frmPpal
'    Unload Me
    BotonSalir
End Sub

Private Sub mnP_Salir2_Click()
'    Unload frmPpal
'    Unload Me
    BotonSalir
End Sub

Private Sub BotonSalir()
    Unload frmPpal
    Unload Me
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' cooperativa
            SubmnC_RecoleccionG_Admon_Click (1)
        Case 2 'Socios
            SubmnC_RecoleccionG_Admon_Click (4)
        Case 3 ' Capataces
            SubmnC_RecoleccionG_Admon_Click (11)
        Case 4 'Transportistas
            SubmnC_RecoleccionG_Admon_Click (9)
        Case 6 ' Campos
            SubmnC_RecoleccionG_Admon_Click (22)
        Case 8 ' Entradas
            SubmnC_RecoleccionG_Entradas_Click (1)
        Case 9 ' Clasificacion
            SubmnC_RecoleccionG_Entradas_Click (9)
        Case 10 ' Historico de entradas
            SubmnC_RecoleccionG_Entradas_Click (13)
        Case 12 ' Mantenimiento de precios
            SubmnC_RecoleccionG_PagoSocios_Click (1)
        Case 13 ' Facturas de socios
            SubmnC_RecoleccionG_FacturasSocios_Click (1) '[Monica]15/02/2013: antes 3
        Case 14 ' entrada de lecturas de pozos (solo para Monasterios)
            SubmnC_RecoleccionG_Pozos_Click (5)
        Case 15 ' Cambio de campaña
            mnCambioEmpresa_Click
        Case 17 ' Salir de la aplicacion
            MDIForm_Unload 0
    End Select
End Sub

' ### [Monica] 05/09/2006
Private Sub PonerDatosFormulario()
Dim Config As Boolean

    Config = (vEmpresa Is Nothing) 'Or (vParamAplic Is Nothing)
    
    If Not Config Then HabilitarSoloPrametros_o_Empresas True

    'FijarConerrores
    CadenaDesdeOtroForm = ""

    'Poner datos visible del form
    PonerDatosVisiblesForm
    
    'Habilitar/Deshabilitar entradas del menu segun el nivel de usuario
    PonerMenusNivelUsuario
    
    '[Monica]20/07/2015: solo para el caso de escalona
    If vParamAplic.Cooperativa = 10 Then
        Me.mnRec_Pozos(18).Caption = "Duplicado de Recibos"
    End If

    'Si no hay carpeta interaciones, no habra integraciones
'    Me.mnComprobarPendientes.Enabled = vConfig.Integraciones <> ""


    'Habilitar
    If Config Then HabilitarSoloPrametros_o_Empresas False
    'Panel con el nombre de la empresa
'    If Not vEmpresa Is Nothing Then
'        Me.StatusBar1.Panels(2).Text = "Empresa:   " & vEmpresa.nomempre & "               Código: " & vEmpresa.codempre
'    Else
'        Me.StatusBar1.Panels(2).Text = "Falta configurar"
'    End If


    BloqueoMenusSegunCooperativa
    
    BloqueoMenusSegunNivelUsuario
    
    BloqueoMenusSegunCampanya
    
    ' activamos el menu de traza si la cooperativa lo lleva
    Me.mnTrazabilidad.Enabled = vParamAplic.HayTraza
    Me.mnTrazabilidad.visible = vParamAplic.HayTraza
    
    '[Monica]19/02/2013: Facturas Varias segun parametros
    Me.mnRec_FactSocios(4).visible = vParamAplic.HayFacVarias ' en la barra no se permite enable
    Me.mnRec_FactSocios(5).Enabled = vParamAplic.HayFacVarias
    Me.mnRec_FactSocios(5).visible = vParamAplic.HayFacVarias
    Me.mnRec_FactSocios(6).visible = vParamAplic.HayFacVarias ' en la barra no se permite enable
    Me.mnRec_FactSocios(7).Enabled = vParamAplic.HayFacVarias
    Me.mnRec_FactSocios(7).visible = vParamAplic.HayFacVarias
    
    ' activamos las integraciones al aridoc unicamente si hay aridoc
    Me.mnRec_Transporte(7).Enabled = vParamAplic.HayAridoc
    Me.mnRec_FacturasSocios(7).Enabled = vParamAplic.HayAridoc '[Monica]15/02/2013: antes 8, con el cambio de menu
    
    '[Monica]07/03/2013: Habilitamos si tiene el path de facturacion
    Me.mnRec_FacturasSocios(4).Enabled = vParamAplic.PathFacturaE <> ""
    
    Me.mnP_PreNominas(17).Enabled = vParamAplic.HayAridoc
    Me.mnRec_Almaz(9).Enabled = vParamAplic.HayAridoc
    Me.mnRec_Bod1(9).Enabled = vParamAplic.HayAridoc
    Me.mnRec_ADV(14).Enabled = vParamAplic.HayAridoc
    Me.mnRec_Pozos(25).Enabled = vParamAplic.HayAridoc
    Me.mnRec_FacturasVarias(8).Enabled = vParamAplic.HayAridoc
    Me.mnRec_FacturasVariasPr(5).Enabled = vParamAplic.HayAridoc
    
    
    'Si tiene editor de menus
    If TieneEditorDeMenus Then PoneMenusDelEditor
    
    'Comprobar que los iconos de la barra su correspondiente
    'entrada de menu esta habilitada sino desabilitar
    PoneBarraMenus
    
    '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Or vParamAplic.Cooperativa = 17 Then 'utxera
        MDIppal.mnRec_Pozos.item(2).Caption = "Contadores"
    Else
        MDIppal.mnRec_Pozos.item(2).Caption = "Hidrantes"
    End If
    
    '[Monica]02/09/2013: calculo de digito de control
    Me.mnSoporte.visible = (vUsu.Login = "root")
    Me.mnSoporte.Enabled = (vUsu.Login = "root")
    
    
    
End Sub

' ### [Monica] 05/09/2006
Private Sub HabilitarSoloPrametros_o_Empresas(Habilitar As Boolean)
Dim T As Control
Dim Cad As String

    On Error Resume Next
    For Each T In Me
        Cad = T.Name
        If Mid(T.Name, 1, 2) = "mn" Then
            'If LCase(Mid(T.Name, 1, 8)) <> "mn_b" Then
                T.Enabled = Habilitar
            'End If
        End If
    Next
    
    Me.Toolbar1.Enabled = Habilitar
    Me.Toolbar1.visible = Habilitar
    Me.mnParametros(1).Enabled = True
    Me.mnP_Generales(1).Enabled = True
    Me.mnP_Generales(2).Enabled = True
    Me.mnP_Generales(6).Enabled = True
    Me.mnP_Generales(17).Enabled = True
    
'    Me.mnCambioEmpresa.Enabled = True
End Sub


' ### [Monica] 07/11/2006
' añadida esta parte para la personalizacion de menus

Private Sub LeerEditorMenus()
Dim Sql As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    TieneEditorDeMenus = False
    Sql = "Select count(*) from usuarios.appmenus where aplicacion='Ariagrorec'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            If miRsAux.Fields(0) > 0 Then TieneEditorDeMenus = True
        End If
    End If
    miRsAux.Close
        

ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub PoneMenusDelEditor()
Dim T As Control
Dim Sql As String
Dim C As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    
    Sql = "Select * from usuarios.appmenususuario where aplicacion='AriagroRec' and codusu = " & Val(Right(CStr(vUsu.Codigo - vUsu.DevuelveAumentoPC), 3))
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""

    While Not miRsAux.EOF
        If Not IsNull(miRsAux.Fields(3)) Then
            Sql = Sql & miRsAux.Fields(3) & "·"
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
   
    If Sql <> "" Then
        Sql = "·" & Sql
        For Each T In Me.Controls
            If TypeOf T Is menu Then
                C = DevuelveCadenaMenu(T)
                C = "·" & C & "·"
                If InStr(1, Sql, C) > 0 Then T.visible = False
           
            End If
        Next
    End If
ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function DevuelveCadenaMenu(ByRef T As Control) As String

On Error GoTo EDevuelveCadenaMenu
    DevuelveCadenaMenu = T.Name & "|"
    DevuelveCadenaMenu = DevuelveCadenaMenu & T.Index '& "|"   Monica:con esto no funcionaba
    Exit Function
EDevuelveCadenaMenu:
    Err.Clear
    
End Function

Private Sub LanzaHome(Opcion As String)
    Dim i As Integer
    Dim Cad As String
    On Error GoTo ELanzaHome
    
    'Obtenemos la pagina web de los parametros
    CadenaDesdeOtroForm = DevuelveDesdeBD("websoporte", "sparam", "codparam", 1, "N")
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Falta configurar los datos en parametros.", vbExclamation
        Exit Sub
    End If
        
    i = FreeFile
    Cad = ""
    Open App.Path & "\lanzaexp.dat" For Input As #i
    Line Input #i, Cad
    Close #i
    
    'Lanzamos
    If Cad <> "" Then Shell Cad & " " & CadenaDesdeOtroForm, vbMaximizedFocus
    
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, Cad & vbCrLf & Err.Description
    CadenaDesdeOtroForm = ""
End Sub



Private Sub CargarImagen()

On Error GoTo eCargarImagen
    Me.Picture = LoadPicture(App.Path & "\fondo.dat")
    Exit Sub
eCargarImagen:
    MuestraError Err.Number, "Error cargando imagen. LLame a soporte"
    End
End Sub

Private Sub PonerMenusNivelUsuario()
Dim b As Boolean

'    b = (vSesion.Nivel = 0)    'Administradores y root

'    Me.mnE_Util(11).Enabled = b
'    Me.mnE_Util(11).visible = b
    
End Sub

Public Sub mnCambioEmpresa_Click()
    Dim AntUSU As Usuario

    If Not (Me.ActiveForm Is Nothing) Then
        MsgBox "Cierre todas las ventanas para poder cambiar de usuario", vbExclamation
        Exit Sub
    End If

    'Borramos temporal
    conn.Execute "Delete from zbloqueos where codusu = " & vUsu.Codigo


    CadenaDesdeOtroForm = vUsu.Login & "|" & vUsu.PasswdPROPIO & "|"
    
'    Set AntUSU = vUsu
'    Set vUsu = Nothing
    frmLogin.Show vbModal
'    If vUsu Is Nothing Then
'        Set vUsu = AntUSU
'        Set AntUSU = Nothing
'        Exit Sub
'    End If

'[Monica]29/06/2017: quito lo de la campaña anterior
'    If vParamAplic.ContabilidadNueva And (vUsu.Nivel = 0 Or vUsu.Nivel = 1) Then FrasPendientesContabilizar True


    Screen.MousePointer = vbHourglass
    'Cerramos la conexion
    conn.Close
'--monica: la conexion de contabilidad va por secciones
'    If vParamAplic.NumeroConta <> 0 Then ConnConta.Close


    'Abre la conexión a BDatos:Ariges
    If AbrirConexion() = False Then
        MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
        End
    Else
        Set vParam = Nothing
        Set vParamAplic = Nothing
        'Carga Parametros Generales y Contables de la empresa
        LeerParametros
    End If

'--monica: no se abre la conexion a contabilidad de forma global, se hace por secciones
'    'Abrir conexión a la BDatos de Contabilidad para acceder a
'    'Tablas: Cuentas, Tipos IVA
'    If vParamAplic.NumeroConta <> 0 Then
'        If AbrirConexionConta() = False Then
'            MsgBox "La aplicación no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
'            End
'        End If
'    End If
'
'
    Set vEmpresa = Nothing
    LeerDatosEmpresa




    PonerDatosFormulario
    
    If vParamAplic.ContabilidadNueva And (vUsu.Nivel = 0 Or vUsu.Nivel = 1) Then FrasPendientesContabilizar True
    

    'Ponemos primera vez a false
    PrimeraVez = True
    Me.SetFocus

    Screen.MousePointer = vbDefault
End Sub



Private Sub PonerDatosVisiblesForm()
'Escribe texto de la barra de la aplicación
Dim Cad As String
    Cad = UCase(Mid(Format(Now, "dddd"), 1, 1)) & Mid(Format(Now, "dddd"), 2)
    Cad = Cad & ", " & Format(Now, "d")
    Cad = Cad & " de " & Format(Now, "mmmm")
    Cad = Cad & " de " & Format(Now, "yyyy")
    Cad = "    " & Cad & "    "
    Me.StatusBar1.Panels(7).Text = Cad
    
    '
    Me.StatusBar1.Panels(2).Text = vUsu.CadenaConexion
    If Not EsCampanyaActual(vEmpresa.BDAriagro) Then
        Me.StatusBar1.Panels(4).visible = True
    Else
        Me.StatusBar1.Panels(4).visible = False
    End If
    
    
    Cad = ""
    If vParam Is Nothing Then
        Caption = "AriAgro - Recolección   " & " v." & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & " FALTA CONFIGURAR"
    Else
        Caption = "AriAgro - Recolección   " & " v." & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & vParam.NombreEmpresa & Cad & _
                  "   -   " & vEmpresa.nomresum & "   -   Fechas: " & vParam.FecIniCam & " - " & vParam.FecFinCam & "   -   Usuario: " & vUsu.Nombre
    End If
End Sub



Private Sub PoneBarraMenus()
    'cooperativas
    MDIppal.Toolbar1.Buttons(1).Enabled = MDIppal.mnRecG_Admon(1).visible And MDIppal.mnComerGen.visible
    MDIppal.Toolbar1.Buttons(1).visible = MDIppal.mnRecG_Admon(1).visible And MDIppal.mnComerGen.visible
    'Socios
    MDIppal.Toolbar1.Buttons(2).Enabled = MDIppal.mnRecG_Admon(4).visible And MDIppal.mnComerGen.visible
    MDIppal.Toolbar1.Buttons(2).visible = MDIppal.mnRecG_Admon(4).visible And MDIppal.mnComerGen.visible
    'capataces
    MDIppal.Toolbar1.Buttons(3).Enabled = MDIppal.mnRecG_Admon(11).visible And MDIppal.mnComerGen.visible
    MDIppal.Toolbar1.Buttons(3).visible = MDIppal.mnRecG_Admon(11).visible And MDIppal.mnComerGen.visible
    'transportistas
    MDIppal.Toolbar1.Buttons(4).Enabled = MDIppal.mnRecG_Admon(9).visible And MDIppal.mnComerGen.visible
    MDIppal.Toolbar1.Buttons(4).visible = MDIppal.mnRecG_Admon(9).visible And MDIppal.mnComerGen.visible
    'campos
    MDIppal.Toolbar1.Buttons(6).Enabled = MDIppal.mnRecG_Admon(22).visible And MDIppal.mnComerGen.visible
    MDIppal.Toolbar1.Buttons(6).visible = MDIppal.mnRecG_Admon(22).visible And MDIppal.mnComerGen.visible
    'entradas bascula
    MDIppal.Toolbar1.Buttons(8).Enabled = MDIppal.mnRec_Entradas(1).visible And mnEntradas.visible
    MDIppal.Toolbar1.Buttons(8).visible = MDIppal.mnRec_Entradas(1).visible And mnEntradas.visible
    'clasificacion
    MDIppal.Toolbar1.Buttons(9).Enabled = MDIppal.mnRec_Entradas(9).visible And mnEntradas.visible
    MDIppal.Toolbar1.Buttons(9).visible = MDIppal.mnRec_Entradas(9).visible And mnEntradas.visible
    'historico de entradas
    MDIppal.Toolbar1.Buttons(10).Enabled = MDIppal.mnRec_Entradas(13).visible And mnEntradas.visible
    MDIppal.Toolbar1.Buttons(10).visible = MDIppal.mnRec_Entradas(13).visible And mnEntradas.visible
    'precios
    MDIppal.Toolbar1.Buttons(12).Enabled = MDIppal.mnRec_PagoSocios(1).visible And mnPagoSocio.visible
    MDIppal.Toolbar1.Buttons(12).visible = MDIppal.mnRec_PagoSocios(1).visible And mnPagoSocio.visible
    'facturas socios
    MDIppal.Toolbar1.Buttons(13).Enabled = MDIppal.mnRec_FacturasSocios(1).visible And mnFacturasSocios.visible
    MDIppal.Toolbar1.Buttons(13).visible = MDIppal.mnRec_FacturasSocios(1).visible And mnFacturasSocios.visible
    
    'Entrada de lecturas
    If vParamAplic.Cooperativa = 17 Then
        MDIppal.Toolbar1.Buttons(14).Enabled = MDIppal.mnRec_Pozos(35).visible
        MDIppal.Toolbar1.Buttons(14).visible = MDIppal.mnRec_Pozos(35).visible
    Else
        MDIppal.Toolbar1.Buttons(14).Enabled = MDIppal.mnRec_Pozos(5).visible
        MDIppal.Toolbar1.Buttons(14).visible = MDIppal.mnRec_Pozos(5).visible
    End If
    
    'cambio de campaña
    MDIppal.Toolbar1.Buttons(15).Enabled = MDIppal.mnP_Generales(7).visible And mnParametros(1).visible
    MDIppal.Toolbar1.Buttons(15).visible = MDIppal.mnP_Generales(7).visible And mnParametros(1).visible
    
    '[Monica]18/05/2012
    If vParamAplic.Cooperativa = 3 Then
        MDIppal.mnRec_ADV(4).Caption = "&Tipos de Venta"
        MDIppal.mnRec_ADV(5).Caption = "&Albaranes de Venta"
        MDIppal.mnRec_ADV(6).Caption = "&Reimpresión Albaranes Venta"
        MDIppal.mnRec_ADV(11).Caption = "Histórico Albarán/Factura"
    Else
        MDIppal.mnRec_ADV(4).Caption = "&Tratamientos"
        MDIppal.mnRec_ADV(5).Caption = "&Partes de Trabajo"
        MDIppal.mnRec_ADV(6).Caption = "&Reimpresión Partes Trabajo"
        MDIppal.mnRec_ADV(11).Caption = "Histórico Parte/Factura"
    End If

End Sub
