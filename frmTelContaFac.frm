VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTelContaFac 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integración Contable de Facturas de Telefonia"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6420
   Icon            =   "frmTelContaFac.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCobros 
      Height          =   5760
      Left            =   0
      TabIndex        =   4
      Top             =   45
      Width           =   6375
      Begin VB.Frame Frame2 
         Caption         =   "Datos para la contabilización"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1995
         Left            =   225
         TabIndex        =   33
         Top             =   1845
         Width           =   6000
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   2250
            Locked          =   -1  'True
            TabIndex        =   36
            Text            =   "Text5"
            Top             =   945
            Width           =   3450
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   10
            Left            =   1440
            MaxLength       =   6
            TabIndex        =   26
            Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
            Top             =   945
            Width           =   735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   25
            Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
            Top             =   540
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   35
            Text            =   "Text5"
            Top             =   540
            Width           =   3135
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   11
            Left            =   1440
            MaxLength       =   6
            TabIndex        =   27
            Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
            Top             =   1395
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   2265
            Locked          =   -1  'True
            TabIndex        =   34
            Text            =   "Text5"
            Top             =   1395
            Width           =   3450
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   1170
            MouseIcon       =   "frmTelContaFac.frx":000C
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar cliente"
            Top             =   945
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Iva"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   7
            Left            =   180
            TabIndex        =   39
            Top             =   960
            Width           =   585
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cta. Ventas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   38
            Top             =   555
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   1170
            MouseIcon       =   "frmTelContaFac.frx":015E
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar cliente"
            Top             =   540
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Centro Coste"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   37
            Top             =   1410
            Width           =   960
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   11
            Left            =   1170
            MouseIcon       =   "frmTelContaFac.frx":02B0
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar cliente"
            Top             =   1395
            Width           =   240
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos para Selección"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1230
         Left            =   225
         TabIndex        =   22
         Top             =   495
         Width           =   6000
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   3630
            MaxLength       =   10
            TabIndex        =   24
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   630
            Width           =   1050
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1500
            MaxLength       =   10
            TabIndex        =   23
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   630
            Width           =   1050
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Factura"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   16
            Left            =   180
            TabIndex        =   32
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   15
            Left            =   615
            TabIndex        =   30
            Top             =   675
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   14
            Left            =   2745
            TabIndex        =   28
            Top             =   675
            Width           =   420
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   1215
            Picture         =   "frmTelContaFac.frx":0402
            ToolTipText     =   "Buscar fecha"
            Top             =   630
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   3
            Left            =   3285
            Picture         =   "frmTelContaFac.frx":048D
            ToolTipText     =   "Buscar fecha"
            Top             =   630
            Width           =   240
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3720
         MaxLength       =   7
         TabIndex        =   3
         Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
         Top             =   3150
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1605
         MaxLength       =   7
         TabIndex        =   2
         Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
         Top             =   3150
         Width           =   830
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3570
         TabIndex        =   29
         Top             =   5205
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4755
         TabIndex        =   31
         Top             =   5205
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1605
         MaxLength       =   1
         TabIndex        =   0
         Top             =   2685
         Width           =   345
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   3720
         MaxLength       =   1
         TabIndex        =   1
         Top             =   2700
         Width           =   285
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   315
         TabIndex        =   5
         Top             =   4185
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   13
         Top             =   2910
         Width           =   555
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   2880
         TabIndex        =   12
         Top             =   3195
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   720
         TabIndex        =   11
         Top             =   3150
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   10
         Top             =   2685
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   9
         Top             =   2700
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Letra de Serie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   8
         Top             =   2445
         Width           =   1005
      End
      Begin VB.Label lblProgres 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   4500
         Width           =   5265
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   330
         TabIndex        =   6
         Top             =   4815
         Width           =   5295
      End
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   2490
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text5"
      Top             =   1680
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   1620
      MaxLength       =   6
      TabIndex        =   17
      Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   2805
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text5"
      Top             =   1305
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   1635
      MaxLength       =   10
      TabIndex        =   15
      Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
      Top             =   1305
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   1650
      MaxLength       =   10
      TabIndex        =   14
      Tag             =   "Código Postal|T|S|||clientes|codposta|||"
      Top             =   450
      Visible         =   0   'False
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5940
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   7
      Left            =   1350
      Picture         =   "frmTelContaFac.frx":0518
      ToolTipText     =   "Buscar fecha"
      Top             =   450
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Enabled         =   0   'False
      Height          =   240
      Index           =   9
      Left            =   1335
      MouseIcon       =   "frmTelContaFac.frx":05A3
      MousePointer    =   4  'Icon
      ToolTipText     =   "Buscar cliente"
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Forma Pago"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   195
      Index           =   6
      Left            =   405
      TabIndex        =   21
      Top             =   1695
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgBuscar 
      Enabled         =   0   'False
      Height          =   240
      Index           =   6
      Left            =   1335
      MouseIcon       =   "frmTelContaFac.frx":06F5
      MousePointer    =   4  'Icon
      ToolTipText     =   "Buscar cliente"
      Top             =   1305
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cta. Banco"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   195
      Index           =   3
      Left            =   405
      TabIndex        =   20
      Top             =   1320
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha Vto"
      Enabled         =   0   'False
      ForeColor       =   &H00972E0B&
      Height          =   255
      Index           =   4
      Left            =   405
      TabIndex        =   19
      Top             =   450
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmTelContaFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmCta As frmCtasConta 'Ctas contables
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmFPa As frmForpaConta 'Formas de Pago Contables
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmTIva As frmTipIVAConta 'Tipos de Iva Contables
Attribute frmTIva.VB_VarHelpID = -1
Private WithEvents frmCCos As frmCCosConta 'centros de coste
Attribute frmCCos.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir

Private cadNomRPT As String 'Nombre del informe
Private conSubRPT As Boolean 'Si el informe tiene subreports





Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Dim vSeccion As CSeccion
Dim cContaFra As cContabilizarFacturas


Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim cadMen As String
Dim i As Byte
Dim Sql As String
Dim Tipo As Byte
Dim Nregs As Long
Dim NumError As Long

    If Not DatosOk Then Exit Sub
    
    cadSelect = Tabla & ".intconta=0 "
    
    'D/H Fecha factura
    cDesde = Trim(txtcodigo(2).Text)
    cHasta = Trim(txtcodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'D/H letra de serie
    cDesde = Trim(txtcodigo(4).Text)
    cHasta = Trim(txtcodigo(5).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".numserie}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHColec= """) Then Exit Sub
    End If
    
    'D/H numero de factura
    cDesde = Trim(txtcodigo(0).Text)
    cHasta = Trim(txtcodigo(1).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHColec= """) Then Exit Sub
    End If
    
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    
    ContabilizarFacturas Tabla, cadSelect
     'Eliminar la tabla TMP
    BorrarTMPFacturas
    'Desbloqueamos ya no estamos contabilizando facturas
    DesBloqueoManual ("TELCON") 'VENtas CONtabilizar
    
eError:
    If Err.Number <> 0 Or NumError <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilización. Llame a soporte."
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        ValoresPorDefecto
        PonerFoco txtcodigo(7)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(6).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(8).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(9).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(10).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(11).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     
    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, H, W
    indFrame = 5
    Tabla = "rtelmovil"
    
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
    
    ConexionConta vParamAplic.Seccionhorto

    txtcodigo(8).Text = vParamAplic.CtaVentasTel   ' cuenta contable de ventas de telefonia
    txtNombre(8).Text = PonerNombreCuenta(txtcodigo(8), 0)
    
    txtcodigo(2).Text = Format(Now, "dd/mm/yyyy") ' fecha de factura desde
    txtcodigo(3).Text = Format(Now, "dd/mm/yyyy") ' fecha de factura hasta
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not vSeccion Is Nothing Then
        vSeccion.CerrarConta
        Set vSeccion = Nothing
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtcodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCCos_DatoSeleccionado(CadenaSeleccion As String)
'Centro de Coste de la contabilidad
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Cuentas contables
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'cta contable
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre ctacontable
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de forma de pago
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'forma de pago contable
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre forma de pago
End Sub

Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de tipos de iva
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'tipos de iva contable
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre tipos de iva
End Sub



Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
    Dim obj As Object
    
    Set frmC = New frmCal

    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top

    Set obj = imgFec(Index).Container

    While imgFec(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
       
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + 420 + 30

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(2).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtcodigo(Index).Text <> "" Then frmC.NovaData = txtcodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtcodigo(CByte(imgFec(2).Tag) + 2)
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        
        Case 6, 8 ' ctas contables ventas y banco
            AbrirFrmCtasConta (Index)
        
        Case 9 ' forma de pago
            AbrirFrmForpaConta (Index)
        
        Case 10 ' tipo de iva de contabilidad
            AbrirFrmTipIvaConta (Index)
        
        Case 11 ' centro de coste de contabilidad
            AbrirFrmCCoste (Index)
        
    End Select
    PonerFoco txtcodigo(indCodigo)
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes estaba esto
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 6: KEYBusqueda KeyAscii, 6 'cta contable
            Case 8: KEYBusqueda KeyAscii, 8 'cta contable
            Case 9: KEYBusqueda KeyAscii, 9 'forma de pago
            Case 10: KEYBusqueda KeyAscii, 10 'tipo de iva
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 7: KEYFecha KeyAscii, 7 'fecha de vencimiento
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 6, 8 ' CTAS CONTABLES
            If txtcodigo(Index).Text = "" Then Exit Sub
            txtNombre(Index).Text = PonerNombreCuenta(txtcodigo(Index), 1)
            
        Case 2, 3, 7  'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
        
        Case 9 ' FORMA DE PAGO
            If txtcodigo(Index).Text = "" Then Exit Sub
            If vParamAplic.ContabilidadNueva Then
                txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtcodigo(Index), "N")
            Else
                txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtcodigo(Index), "N")
            End If
        
        Case 10 ' TIPO DE IVA
            If txtcodigo(Index).Text = "" Then Exit Sub
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "tiposiva", "nombriva", "codigiva", "N", cConta)
        
        Case 0, 1 ' NUMERO DE FACTURA
            If txtcodigo(Index).Text <> "" Then PonerFormatoEntero txtcodigo(Index)
        
        Case 4, 5 ' LETRA DE SERIE
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = UCase(txtcodigo(Index).Text)
        
        Case 11 'CENTRO DE COSTE
            txtNombre(11).Text = PonerNombreDeCod(txtcodigo(11), "cabccost", "nomccost", "codccost", "T", cConta)
        
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6015
        Me.FrameCobros.Width = 6555
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub

Private Sub ValoresPorDefecto()
    txtcodigo(7).Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
End Sub

Private Function PonerDesdeHasta(codD As String, codH As String, nomD As String, nomH As String, param As String) As Boolean
'IN: codD,codH --> codigo Desde/Hasta
'    nomD,nomH --> Descripcion Desde/Hasta
'Añade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y añade a cadParam la cadena para mostrar en la cabecera informe:
'       "codigo: Desde codD-nomd Hasta: codH-nomH"
Dim devuelve As String
Dim devuelve2 As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(codD, codH, Codigo, TipCod)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If TipCod <> "F" Then 'Fecha
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadSelect, devuelve2) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            CadParam = CadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Sub AbrirFrmCtasConta(indice As Integer)
    indCodigo = indice
    Set frmCta = New frmCtasConta
    frmCta.DatosADevolverBusqueda = "0|1|"
    frmCta.CodigoActual = txtcodigo(indCodigo)
    frmCta.Show vbModal
    Set frmCta = Nothing
End Sub

Private Sub AbrirFrmForpaConta(indice As Integer)
    indCodigo = indice
    Set frmFPa = New frmForpaConta
    frmFPa.DatosADevolverBusqueda = "0|1|"
    frmFPa.CodigoActual = txtcodigo(indCodigo)
    frmFPa.Show vbModal
    Set frmFPa = Nothing
End Sub

Private Sub AbrirFrmTipIvaConta(indice As Integer)
    indCodigo = indice
    Set frmTIva = New frmTipIVAConta
    frmTIva.DatosADevolverBusqueda = "0|1|"
    frmTIva.CodigoActual = txtcodigo(indCodigo)
    frmTIva.Show vbModal
    Set frmTIva = Nothing
End Sub


Private Sub AbrirFrmCCoste(indice As Integer)
    indCodigo = indice
    Set frmCCos = New frmCCosConta
    frmCCos.DatosADevolverBusqueda = "0|1|"
    frmCCos.CodigoActual = txtcodigo(indice).Text
    frmCCos.Show vbModal
    Set frmCCos = Nothing
    PonerFoco txtcodigo(indice)
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim cadG As String
    b = True

'    If txtCodigo(7).Text = "" And b Then
'        MsgBox "Debe introducir obligatoriamente una Fecha de Vencimiento.", vbExclamation
'        b = False
'        PonerFoco txtCodigo(7)
'    End If
'
'    If txtCodigo(6).Text = "" And b Then
'        MsgBox "Debe introducir obligatoriamente una Cta Contable de Banco.", vbExclamation
'        b = False
'        PonerFoco txtCodigo(6)
'    End If


  If txtcodigo(3).Text = "" Then
        MsgBox "Introduzca la Fecha de Factura a contabilizar.", vbExclamation
        b = False
        PonerFoco txtcodigo(3)
   Else
        ' comprobamos que la contabilizacion se encuentre en los ejercicios contables
         Orden1 = ""
         Orden1 = DevuelveDesdeBDNew(cConta, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")

         Orden2 = ""
         Orden2 = DevuelveDesdeBDNew(cConta, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")
         FIni = CDate(Orden1)
         FFin = CDate(Orden2)
         If Not (CDate(Orden1) <= CDate(txtcodigo(3).Text) And CDate(txtcodigo(3).Text) < CDate(Day(FIni) & "/" & Month(FIni) & "/" & Year(FIni) + 2)) Then
            MsgBox "La Fecha de la contabilización no es del ejercicio actual ni del siguiente. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtcodigo(3)
         End If
   End If


    If b Then
        If txtcodigo(8).Text = "" Then
            MsgBox "Debe introducir obligatoriamente una Cta Contable de Ventas.", vbExclamation
            b = False
            PonerFoco txtcodigo(8)
        Else
            ' comprobamos que la cta contable es del grupo de ventas
            cadG = DevuelveDesdeBDNew(cConta, "parametros", "grupovta", "", "", "")
            If Mid(txtcodigo(8).Text, 1, 1) <> cadG Then
                MsgBox "La Cuenta debe de ser del Grupo de Ventas. Reintroduzca.", vbExclamation
                b = False
                PonerFoco txtcodigo(8)
            End If
        End If
    End If
'--Monica: lo he quitado de momento
'    If txtCodigo(9).Text = "" And b Then
'        MsgBox "Debe introducir obligatoriamente una Forma de Pago.", vbExclamation
'        b = False
'        PonerFoco txtCodigo(9)
'    End If

    If txtcodigo(10).Text = "" And b Then
        MsgBox "Debe introducir obligatoriamente un Tipo de Iva.", vbExclamation
        b = False
        PonerFoco txtcodigo(10)
    End If
     
    If b And txtcodigo(11).Text = "" And vEmpresa.TieneAnalitica Then
        MsgBox "Debe introducir obligatoriamente un Centro de Coste.", vbExclamation
        b = False
        PonerFoco txtcodigo(11)
    End If
     
     
    '07022007 he añadido esto tambien aquí
     If txtcodigo(2).Text = "" Then
        txtcodigo(2).Text = Orden1 'fechaini del ejercicio de la conta
     End If
     
     If txtcodigo(3).Text = "" Then
        txtcodigo(3).Text = Format(Day(CDate(Orden2)), "00") & "/" & Format(Month(CDate(Orden2)), "00") & "/" & Format(Year(CDate(Orden2)) + 1, "0000") 'fecha fin del ejercicio de la conta
     End If


    DatosOk = b
End Function

' copiado del ariges
Private Sub ContabilizarFacturas(cadTabla As String, cadWHERE As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim Sql As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String

    Sql = "TELCON" 'contabilizar facturas de venta

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se pueden Contabilizar Facturas. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    'comprobar que se han rellenado los dos campos de fecha
    'sino rellenar con fechaini o fechafin del ejercicio
    'que guardamos en vbles Orden1,Orden2
    If txtcodigo(2).Text = "" Then
       txtcodigo(2).Text = vEmpresa.FechaIni  'fechaini del ejercicio de la conta
    End If

    If txtcodigo(3).Text = "" Then
       txtcodigo(3).Text = vEmpresa.FechaFin  'fecha fin del ejercicio de la conta
    End If


    'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
    'contabilidad par ello mirar en la BD de la Conta los parámetros
    If Not ComprobarFechasConta(3) Then Exit Sub

    'comprobar si existen en Ariagro (telefonia) facturas anteriores al periodo solicitado
    'sin contabilizar
    If Me.txtcodigo(2).Text <> "" Then
        Sql = "SELECT COUNT(*) FROM " & cadTabla
        Sql = Sql & " WHERE fecfactu <"
        Sql = Sql & DBSet(txtcodigo(2), "F") & " AND intconta=0 "
        If RegistrosAListar(Sql) > 0 Then
            MsgBox "Hay Facturas anteriores sin contabilizar.", vbExclamation
            Exit Sub
        End If
    End If
    

    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================
    'Cargar tabla TEMP con las Facturas que vamos a Trabajar
    b = CrearTMPFacturas(cadTabla, cadWHERE)
    If Not b Then Exit Sub
        
    
    
    'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
'    TerminaBloquear
    Sql = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla
    Sql = Sql & ".numserie=tmpFactu.numserie AND "
    Sql = Sql & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
    
    If Not BloqueaRegistro(Sql, cadWHERE) Then
        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgres Me.Pb1, 100


    'comprobar que las LETRA SERIE de parametros existen en la contabilidad y en Ariagrorec
    '--------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando letras de serie ..."
    b = ComprobarLetraSerie(cadTabla)
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not b Then Exit Sub


    'comprobar que no haya Nº FACTURAS en la contabilidad para esa fecha
    'que ya existan
    '-----------------------------------------------------------------------
    If cadTabla = "rtelmovil" Then
        Me.lblProgres(1).Caption = "Comprobando Nº Facturas en contabilidad ..."
        Sql = "anofaccl>=" & Year(txtcodigo(2).Text) & " AND anofaccl<= " & Year(txtcodigo(3).Text)
        b = ComprobarNumFacturas_new(cadTabla, Sql)
    End If
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then Exit Sub


    'comprobar que todas las CUENTAS de los distintos socios que vamos a
    'contabilizar existen en la Conta: rsocios_seccion.codmaccli IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables socios en contabilidad ..."
    b = ComprobarCtaContable_new("rtelmovil", 1)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then Exit Sub


    'comprobar que todas las CUENTAS de venta de las variedades
    'contabilizar existen en la Conta: vparamaplic.ctaVentasTel IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Ctbles Ventas en contabilidad ..."
    
    b = ComprobarCtaContable_new("rtelmovil", 2)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then Exit Sub



    'comprobar que todos las TIPO IVA de las distintas fecturas que vamos a
    'contabilizar existen en la Conta: rbodfacturas.codiiva1 codiiva2 codiiva3 IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    b = ComprobarIVA(cadTabla, txtcodigo(10).Text)
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not b Then Exit Sub
    
    
    'comprobar si hay contabilidad ANALITICA: conta.parametros.autocoste=1
    'y verificar que las cuentas de rparamaplic.ctaventaalmz rparamaplic.ctagastosalmz
    'empiezan por el digito de conta.parametros.grupogto o conta.parametros.grupovta
    'obtener el centro de coste del usuario para insertarlo en linfact
    If vEmpresa.TieneAnalitica Then  'hay contab. analitica
       Me.lblProgres(1).Caption = "Comprobando Contabilidad Analítica ..."
           
       b = ComprobarCtaContable_new("rtelmovil", 7)
       If b Then
            '(si tiene analítica requiere un centro de coste para insertar en conta.linfact)
            CCoste = txtcodigo(11).Text
            If CCoste <> "" Then
                ' lo buscamos directamente
                b = (DevuelveDesdeBDNew(cConta, "cabccost", "codccost", "codccost", CCoste, "T") <> "")
            Else
                MsgBox "La empresa tiene analítica. Debe introducir un Centro de Coste.", vbExclamation
                b = False
            End If

       End If
       If Not b Then Exit Sub

    End If
    IncrementarProgres Me.Pb1, 10
    Me.Refresh

' -- Monica: lo quito de momento
'    If b Then
'       Me.lblProgres(1).Caption = "Comprobando Forma de Pago ..."
'       b = ComprobarFormadePago(cadTABLA)
'       If Not b Then Exit Sub
'    End If
    IncrementarProgres Me.Pb1, 10
    Me.Refresh


    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Facturas Telefonia: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad..."


    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Contabilizar facturas Telefonia: " & vbCrLf & cadTabla & vbCrLf & cadWHERE
    Set LOG = Nothing
    '-----------------------------------------------------------------------------


    '---- Crear tabla TEMP para los posible errores de facturas
    tmpErrores = CrearTMPErrFact(cadTabla)

    '---- Pasar las Facturas a la Contabilidad
    b = PasarFacturasAContab(cadTabla, "", txtcodigo(8).Text, txtcodigo(11).Text)

    '---- Mostrar ListView de posibles errores (si hay)
    If Not b Then
        If tmpErrores Then
            'Cargar un listview con la tabla TEMP de Errores y mostrar
            'las facturas que fallaron
            frmMensajes.OpcionMensaje = 10
            frmMensajes.Show vbModal
        Else
            MsgBox "No pueden mostrarse los errores.", vbInformation
        End If
    Else
        MsgBox "El proceso ha finalizado correctamente.", vbInformation
    End If

    'Este bien o mal, si son proveedores abriremos el listado
    'Imprimimiremos un listado de contabilizacion de facturas
    '------------------------------------------------------
    If DevuelveValor("Select count(*) from tmpinformes where codusu = " & vUsu.Codigo) > 0 Then
        InicializarVbles
        CadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
        
        CadParam = "|pDHFecha=""" & vUsu.Nombre & "   Hora: " & Format(Now, "hh:mm") & """|"
        numParam = numParam + 1
        cadFormula = "({tmpinformes.codusu} =" & vUsu.Codigo & ")"
        conSubRPT = False
        cadTitulo = "Listado contabilizacion FRAFRE"
        cadNomRPT = "rContabFRE.rpt"
        LlamarImprimir
    End If


    '---- Eliminar tabla TEMP de Errores
    BorrarTMPErrFact

End Sub


Private Function PasarFacturasAContab(cadTabla As String, FecVenci As String, Banpr As String, CCoste As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim i As Integer
Dim numfactu As Integer
Dim Codigo1 As String

    On Error GoTo EPasarFac

    PasarFacturasAContab = False
    
    'Total de Facturas a Insertar en la contabilidad
    Sql = "SELECT count(*) "
    Sql = Sql & " FROM " & cadTabla & " INNER JOIN tmpfactu "
    Codigo1 = "numserie"
    Sql = Sql & " ON " & cadTabla & "." & Codigo1 & "=tmpfactu." & Codigo1
    Sql = Sql & " AND " & cadTabla & ".numfactu=tmpfactu.numfactu AND " & cadTabla & ".fecfactu=tmpfactu.fecfactu "
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        numfactu = Rs.Fields(0)
    Else
        numfactu = 0
    End If
    Rs.Close
    Set Rs = Nothing

    'Modificacion como David
    '-----------------------------------------------------------
    ' Mosrtaremos para cada factura de PROVEEDOR
    ' que numregis le ha asignado
    Sql = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute Sql

    Set cContaFra = New cContabilizarFacturas
    
    If Not cContaFra.EstablecerValoresInciales(ConnConta) Then
        'NO ha establcedio los valores de la conta.  Le dejaremos seguir, avisando que
        ' obviamente, no va a contabilizar las FRAS
        Sql = "Si continua, las facturas se insertaran en el registro, pero no serán contabilizadas" & vbCrLf
        Sql = Sql & "en este momento. Deberán ser contabilizadas desde el ARICONTA" & vbCrLf & vbCrLf
        Sql = Sql & Space(50) & "¿Continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If



    If numfactu > 0 Then
        CargarProgres Me.Pb1, numfactu
        
        Sql = "SELECT * "
        Sql = Sql & " FROM tmpfactu "
            
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
        i = 1

        b = True
        'contabilizar cada una de las facturas seleccionadas
        While Not Rs.EOF
            Sql = cadTabla & "." & Codigo1 & "=" & DBSet(Rs.Fields(0), "T") & " and numfactu=" & DBLet(Rs!numfactu, "N")
            Sql = Sql & " and fecfactu=" & DBSet(Rs!fecfactu, "F")
            If PasarFacturaTel(Sql, txtcodigo(11).Text, txtcodigo(8).Text, txtcodigo(10).Text, cContaFra) = False And b Then b = False
            
            IncrementarProgres Me.Pb1, 1
            Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad...   (" & i & " de " & numfactu & ")"
            Me.Refresh
            i = i + 1
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
    End If
    Set cContaFra = Nothing

EPasarFac:
    If Err.Number <> 0 Then b = False
    
    If b Then
        PasarFacturasAContab = True
    Else
        PasarFacturasAContab = False
    End If
End Function


Private Function ComprobarFechasConta(ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim Cad As String
Dim Rs As ADODB.Recordset
    
Dim vSeccion As CSeccion
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    
    If txtcodigo(ind).Text <> "" Then
        FechaIni = "Select fechaini,fechafin From parametros"
        Set Rs = New ADODB.Recordset
        Rs.Open FechaIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs.EOF Then
            FechaIni = DBLet(Rs!FechaIni, "F")
            FechaFin = DateAdd("yyyy", 1, DBLet(Rs!FechaFin, "F"))
            'nos guardamos los valores
            Orden1 = FechaIni
            Orden2 = FechaFin
        
            If Not EntreFechas(FechaIni, txtcodigo(ind).Text, FechaFin) Then
                 Cad = "El período de contabilización debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 Cad = Cad & "    Desde: " & FechaIni & vbCrLf
                 Cad = Cad & "    Hasta: " & FechaFin
                 MsgBox Cad, vbExclamation
                 txtcodigo(ind).Text = ""
            Else
                ComprobarFechasConta = True
            End If
        End If
        Rs.Close
        Set Rs = Nothing
    Else
        ComprobarFechasConta = True
    End If
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function


Private Sub ConexionConta(Seccion As String)
    
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(Seccion) Then
            vSeccion.AbrirConta
        End If
    Else
        ' si el objeto existia: cerramos la conexion y volvemos crearlo
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(Seccion) Then
            vSeccion.AbrirConta
        End If
    End If
End Sub


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 0
        .Titulo = cadTitulo
        .NombreRPT = cadNomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub


