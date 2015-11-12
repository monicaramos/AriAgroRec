VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmControlDestrio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Destrio"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12585
   Icon            =   "frmControlDestrio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   9390
      TabIndex        =   53
      Text            =   "Text3"
      Top             =   7590
      Width           =   790
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   8580
      TabIndex        =   52
      Text            =   "Text3"
      Top             =   7590
      Width           =   790
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   7770
      TabIndex        =   51
      Text            =   "Text3"
      Top             =   7590
      Width           =   790
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   6960
      TabIndex        =   50
      Text            =   "Text3"
      Top             =   7590
      Width           =   790
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   6150
      TabIndex        =   49
      Text            =   "Text3"
      Top             =   7590
      Width           =   790
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   5340
      TabIndex        =   48
      Text            =   "Text3"
      Top             =   7590
      Width           =   790
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   4560
      TabIndex        =   47
      Text            =   "Text3"
      Top             =   7590
      Width           =   785
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   3750
      TabIndex        =   46
      Text            =   "Text3"
      Top             =   7590
      Width           =   790
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2970
      TabIndex        =   45
      Text            =   "Text3"
      Top             =   7590
      Width           =   785
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   44
      Text            =   "Text3"
      Top             =   7590
      Width           =   790
   End
   Begin VB.Frame Frame2 
      Height          =   1275
      Index           =   0
      Left            =   135
      TabIndex        =   11
      Top             =   495
      Width           =   12225
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   11100
         MaxLength       =   7
         TabIndex        =   6
         Tag             =   "Porc.Destrio|N|S|||rcontrol|porcdest|##0.00||"
         Text            =   "1234567"
         Top             =   720
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   9090
         MaxLength       =   7
         TabIndex        =   5
         Tag             =   "Kilos Man|N|S|||rcontrol|kilosman|###,##0||"
         Text            =   "1234567"
         Top             =   720
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   6540
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Entrada|F|N|||rcontrol|fechacla|dd/mm/yyyy|S|"
         Text            =   "1234567890"
         Top             =   330
         Width           =   1020
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   6540
         MaxLength       =   8
         TabIndex        =   3
         Tag             =   "Nombre|N|N|||rcontrol|codcampo|00000000|S|"
         Text            =   "12345678"
         Top             =   675
         Width           =   825
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1950
         TabIndex        =   28
         Text            =   "12345678901234567890"
         Top             =   720
         Width           =   3390
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1950
         TabIndex        =   22
         Text            =   "12345678901234567890"
         Top             =   330
         Width           =   3390
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Variedad|N|N|0|999999|rcontrol|codvarie|000000|S|"
         Text            =   "123456"
         Top             =   345
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1230
         MaxLength       =   6
         TabIndex        =   1
         Tag             =   "Nombre|N|N|||rcontrol|codsocio|000000|S|"
         Text            =   "123456"
         Top             =   720
         Width           =   690
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   9090
         MaxLength       =   7
         TabIndex        =   4
         Tag             =   "Nro.Clasif|N|S|||rcontrol|nroclasif|0000000|S|"
         Text            =   "1234567"
         Top             =   330
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   9090
         MaxLength       =   7
         TabIndex        =   61
         Tag             =   "Ordinal|N|S|||rcontrol|ordinal|0000|S|"
         Text            =   "1234567"
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "%Destrio"
         Height          =   255
         Index           =   2
         Left            =   10170
         TabIndex        =   60
         Top             =   750
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos Manuales"
         Height          =   255
         Index           =   1
         Left            =   7980
         TabIndex        =   59
         Top             =   750
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   5550
         TabIndex        =   30
         Top             =   360
         Width           =   480
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   6285
         Picture         =   "frmControlDestrio.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   330
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Campo"
         Height          =   255
         Index           =   0
         Left            =   5520
         TabIndex        =   29
         Top             =   705
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   930
         ToolTipText     =   "Buscar Socio"
         Top             =   750
         Width           =   240
      End
      Begin VB.Label Label10 
         Caption         =   "Variedad"
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   375
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   930
         ToolTipText     =   "Buscar Variedad"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Socio"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Top             =   750
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "Nro.Clasif."
         Height          =   255
         Index           =   0
         Left            =   7980
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1350
      TabIndex        =   24
      Text            =   "Text3"
      Top             =   7590
      Width           =   790
   End
   Begin VB.Frame FrameAux0 
      Caption         =   "Kilos Muestreo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6240
      Left            =   150
      TabIndex        =   16
      Top             =   1800
      Width           =   12195
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   19
         Left            =   10680
         MaxLength       =   7
         TabIndex        =   62
         Tag             =   "Ordinal|N|N|||rcontrol_plagas|ordinal|0000|S|"
         Text            =   "ordinal"
         Top             =   3030
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   10830
         TabIndex        =   58
         Text            =   "Text3"
         Top             =   5790
         Width           =   790
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   10020
         TabIndex        =   57
         Text            =   "Text3"
         Top             =   5790
         Width           =   790
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   18
         Left            =   11580
         MaxLength       =   7
         TabIndex        =   56
         Text            =   "Porcen"
         Top             =   2550
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   17
         Left            =   10920
         MaxLength       =   7
         TabIndex        =   55
         Text            =   "Totkilo"
         Top             =   2550
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   3450
         MaxLength       =   10
         TabIndex        =   43
         Text            =   "nomplag"
         Top             =   2550
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   16
         Left            =   2790
         MaxLength       =   7
         TabIndex        =   42
         Tag             =   "Plaga|N|N|||rcontrol_plagas|idplaga|000|S|"
         Text            =   "idplaga"
         Top             =   2550
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   15
         Left            =   10260
         MaxLength       =   7
         TabIndex        =   41
         Tag             =   "Kilos 11|N|S|||rcontrol_plagas|kilosplaga11|###,##0||"
         Text            =   "kilos11"
         Top             =   2550
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   14
         Left            =   9660
         MaxLength       =   7
         TabIndex        =   40
         Tag             =   "Kilos 10|N|S|||rcontrol_plagas|kilosplaga10|###,##0||"
         Text            =   "kilos10"
         Top             =   2550
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   13
         Left            =   9030
         MaxLength       =   7
         TabIndex        =   39
         Tag             =   "Kilos 9|N|S|||rcontrol_plagas|kilosplaga9|###,##0||"
         Text            =   "kilos9"
         Top             =   2550
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   12
         Left            =   8400
         MaxLength       =   7
         TabIndex        =   38
         Tag             =   "Kilos 8|N|S|||rcontrol_plagas|kilosplaga8|###,##0||"
         Text            =   "kilos8"
         Top             =   2550
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   11
         Left            =   7800
         MaxLength       =   7
         TabIndex        =   37
         Tag             =   "Kilos 7|N|S|||rcontrol_plagas|kilosplaga7|###,##0||"
         Text            =   "kilos7"
         Top             =   2550
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   10
         Left            =   7200
         MaxLength       =   7
         TabIndex        =   36
         Tag             =   "Kilos 6|N|S|||rcontrol_plagas|kilosplaga6|###,##0||"
         Text            =   "kilos6"
         Top             =   2550
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   9
         Left            =   6570
         MaxLength       =   7
         TabIndex        =   35
         Tag             =   "Kilos 5|N|S|||rcontrol_plagas|kilosplaga5|###,##0||"
         Text            =   "kilos5"
         Top             =   2550
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   8
         Left            =   5970
         MaxLength       =   7
         TabIndex        =   34
         Tag             =   "Kilos 4|N|S|||rcontrol_plagas|kilosplaga4|###,##0||"
         Text            =   "kilos4"
         Top             =   2550
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   7
         Left            =   5400
         MaxLength       =   7
         TabIndex        =   33
         Tag             =   "Kilos 3|N|S|||rcontrol_plagas|kilosplaga3|###,##0||"
         Text            =   "kilos3"
         Top             =   2550
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   6
         Left            =   4800
         MaxLength       =   7
         TabIndex        =   32
         Tag             =   "Kilos 2|N|S|||rcontrol_plagas|kilosplaga2|###,##0||"
         Text            =   "kilos2"
         Top             =   2550
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   5
         Left            =   4140
         MaxLength       =   7
         TabIndex        =   31
         Tag             =   "Kilos 1|N|S|||rcontrol_plagas|kilosplaga1|###,##0||"
         Text            =   "kilos1"
         Top             =   2550
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   2130
         MaxLength       =   7
         TabIndex        =   27
         Tag             =   "Campo|N|N|||rcontrol_plagas|codcampo|00000000|S|"
         Text            =   "campo"
         Top             =   2550
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   1020
         MaxLength       =   6
         TabIndex        =   25
         Tag             =   "Calidad|N|N|||rcontrol_plagas|codsocio|000000|S|"
         Text            =   "socio"
         Top             =   2565
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   26
         Tag             =   "Fecha Entrada|F|N|||rcontrol_plagas|fechacla|dd/mm/yyyy|S|"
         Text            =   "fecha"
         Top             =   2550
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   1
         Left            =   495
         MaxLength       =   6
         TabIndex        =   18
         Tag             =   "Variedad|N|N|||rcontrol_plagas|codvarie|000000|S|"
         Text            =   "Var"
         Top             =   2565
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   45
         MaxLength       =   16
         TabIndex        =   17
         Tag             =   "Nro.Nota|N|N|||rcontrol_plagas|nroclasif|0000000|S|"
         Text            =   "nroclasif"
         Top             =   2565
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   19
         Top             =   225
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   0
         Left            =   3720
         Top             =   225
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "AdoAux(0)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGridAux 
         Bindings        =   "frmControlDestrio.frx":0097
         Height          =   5010
         Index           =   0
         Left            =   135
         TabIndex        =   20
         Top             =   630
         Width           =   11850
         _ExtentX        =   20902
         _ExtentY        =   8837
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label11 
         Caption         =   "TOTALES: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   54
         Top             =   5790
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   8100
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11325
      TabIndex        =   8
      Top             =   8190
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10110
      TabIndex        =   7
      Top             =   8190
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   1980
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cálculo Gastos"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Exportar/Importar Excel"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Index           =   0
         Left            =   8520
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11310
      TabIndex        =   13
      Top             =   8190
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnGastos 
         Caption         =   "&Cálculo Gastos"
         Enabled         =   0   'False
         Shortcut        =   ^G
         Visible         =   0   'False
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnExporImpor 
         Caption         =   "Exportar/Importar"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmControlDestrio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                   -+-+
' +-+- Menú: CLIENTES                  -+-+
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps numérics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => mínim 1; si no PK => mínim 0; màxim => 99; format => 00)
' (si es DECIMAL; mínim => 0; màxim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit

'Dim T1 As Single

Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes ' campos del socio
Attribute frmMens.VB_VarHelpID = -1
'Private WithEvents frmExp As frmExpImpExcel ' Exportacion o importacion a pagina excel

'Private WithEvents frmArt As frmManArtic 'articulos
Private WithEvents frmVar As frmComVar 'variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmCap As frmManCapataz 'capataces
Attribute frmCap.VB_VarHelpID = -1
Private WithEvents frmTra As frmManTranspor 'tranportistas
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmTar As frmManTarTra 'tarifas de transporte
Attribute frmTar.VB_VarHelpID = -1
Private WithEvents frmInc As frmManInciden 'incidencias
Attribute frmInc.VB_VarHelpID = -1
Private WithEvents frmCal As frmManCalidades 'calidades
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmCam As frmManCampos 'campos
Attribute frmCam.VB_VarHelpID = -1
'
'*****************************************************
Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'+-+-Variables comuns a tots els formularis+-+-+

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Llínies

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient
Dim TituloLinea As String 'Descripció de la llínia que està en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula
Private NomTablaLineas As String 'Nom de la Taula de llínies del Mantenimient en que estem

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de búsqueda posar el valor de població seleccionada i no tornar a recuperar de la Base de Datos

Dim Gastos As Boolean

Dim CodTipoMov As String
Dim NotaExistente As Boolean

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim VarieAnt As String


Dim cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Dim cadparam As String  'Cadena con los parametros para Crystal Report
Dim numParam As Byte  'Numero de parametros que se pasan a Crystal Report
Dim cadSelect As String  'Cadena para comprobar si hay datos antes de abrir Informe
Dim cadTitulo As String  'Titulo para la ventana frmImprimir
Dim cadNombreRPT As String  'Nombre del informe
Dim cadSelect1 As String 'Cadena para comprobar si hay datos antes de abrir Informe


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos



Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
'        Case 0 'Variedades
'            Set frmVar = New frmComVar
'            frmVar.DatosADevolverBusqueda = "0|1|"
'            frmVar.CodigoActual = txtAux(1).Text
'            frmVar.Show vbModal
'            Set frmVar = Nothing
'            PonerFoco txtAux(1)
        Case 1 'Incidencia
            indice = 9
            Set frmInc = New frmManInciden
            frmInc.DatosADevolverBusqueda = "0|1|"
            frmInc.CodigoActual = txtAux(9).Text
            frmInc.Show vbModal
            Set frmInc = Nothing
            PonerFoco txtAux(9)
        Case 2 'calidades
            indice = Index
            Set frmCal = New frmManCalidades
            frmCal.DatosADevolverBusqueda = "2|3|"
            frmCal.CodigoActual = txtAux(2).Text
            frmCal.Show vbModal
            Set frmCal = Nothing
            PonerFoco txtAux(2)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                NotaExistente = False
                InsertarCabecera
            
            
'                If InsertarDesdeForm2(Me, 1) Then
'                    text2(9).Text = PonerNombreCuenta(text1(9), Modo, text1(0).Text)
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
'                If Gastos Then CalcularGastos

                If ModificaCabecera Then
                    espera 0.2
                    TerminaBloquear
                    PosicionarData
                    PonerCampos
                End If
'[Monica]24/06/2010
'                CalcularGastos
'
'                If ModificaDesdeFormulario2(Me, 1) Then
'                    TerminaBloquear
'                    PosicionarData
'                End If
'24/06/2010
            Else
                ModoLineas = 0
            End If
        ' *** si n'hi han llínies ***
        
        
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    
                Case 2 'modificar llínies
                    If ModificarLinea Then
                        PosicionarData
                    Else
                        PonerFoco txtAux(5)
                    End If
            End Select
            'nuevo calculamos los totales de lineas
            CalcularTotales
                    

        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
'    If PrimeraVez Then PrimeraVez = False
    If PrimeraVez Then
        PrimeraVez = False
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                BotonAnyadir
            Else
                PonerModo 1 'búsqueda
                ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
                Text1(0).BackColor = vbYellow 'nro nota
                ' ****************************************************************************
            End If
        End If
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True
 
    ' ICONETS DE LA BARRA
    btnPrimero = 16 'index del botó "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(3).Image = 1   'Buscar
        .Buttons(4).Image = 2   'Totss
        'el 5 i el 6 son separadors
        .Buttons(7).Image = 3   'Insertar
        .Buttons(8).Image = 4   'Modificar
        .Buttons(9).Image = 5   'Borrar
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(14).Image = 11  'Eixir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primer
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Següent
        .Buttons(btnPrimero + 3).Image = 9 'Últim
    End With
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For i = 0 To ToolAux.Count - 1
        With Me.ToolAux(i)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next i
    ' ***********************************
    
    'cargar IMAGES .Image =de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
'    Me.imgBuscar(8).Picture = frmPpal.imgListComun.ListImages(21).Picture
   
    'carga IMAGES de mail
'    For i = 0 To Me.imgMail.Count - 1
'        Me.imgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
'    Next i
    
    
    CodTipoMov = "NOC"
    
    ' *** si n'hi han combos (capçalera o llínies) ***
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(0).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "rcontrol"
    Ordenacion = " ORDER BY codvarie, fechacla, codsocio, codcampo, nroclasif"
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codvarie=-1"
    Data1.Refresh
       
    CargaGrid 0, False
       
    ModoLineas = 0
       
    
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'    Else
'        PonerModo 1 'búsqueda
'        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
'        Text1(0).BackColor = vbYellow 'codforfait
'        ' ****************************************************************************
'    End If
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    ' *** si n'hi han combos a la capçalera ***
    ' *****************************************

    Text3(0).Text = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funció del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim i As Integer, Numreg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de búsqueda, el botó "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
 '   Text1(5).Enabled = True
    
    
    '=======================================
    b = (Modo = 2)
    'Posar Fleches de desplasament visibles
    Numreg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then Numreg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
    '---------------------------------------------
    
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    CmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    
    b = (Modo <> 1) And (Modo <> 3)
    'Campos Nº entrada bloqueado y en azul
    For i = 0 To 4
        BloquearTxt Text1(i), b, False
    Next i
    ' kilos mantenimiento y porcentaje de destrio solo se modifican
    BloquearTxt Text1(5), (Modo <> 1 And Modo <> 4), False
    BloquearTxt Text1(6), (Modo <> 1), False
    
    
    '*** si n'hi han combos a la capçalera ***
'    BloquearCombo Me, Modo
    '**************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
'    BloquearImgZoom Me, Modo, ModoLineas
    BloquearImgFec Me, 0, Modo, ModoLineas
    
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos
    
'    Frame4.Enabled = (Modo = 1)
'  cambiado por esto
'
    For i = 0 To 1
        imgBuscar(i).Enabled = (Modo = 1) Or (Modo = 3)
        imgBuscar(i).visible = (Modo = 1) Or (Modo = 3)
    Next i
    
    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
      
    ' ****************************************************************
    
    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim i As Byte
    
    'Barra de CAPÇALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(4).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(7).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(12).Enabled = True And Not DeConsulta
    Me.mnImprimir.Enabled = True And Not DeConsulta
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    b = (Modo = 2) And Not DeConsulta
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botons de Desplaçament; per a desplaçar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub

Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informació proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enlaça en el data1
'           -> Si no el carreguem sense enllaçar a cap camp
'--------------------------------------------------------------------
Dim Sql As String
Dim Tabla As String
Dim KilosTot As Long

    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 'CLASIFICACION
            Sql = "select sum(kilosplaga1+kilosplaga2+kilosplaga3+kilosplaga4+kilosplaga5+kilosplaga6+kilosplaga7+kilosplaga8+kilosplaga9+kilosplaga10+kilosplaga11) total "
            Sql = Sql & " from rcontrol_plagas "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE rcontrol_plagas.codvarie = -1"
            End If
            Sql = Sql & " and idplaga <> 2 "
            KilosTot = DevuelveValor(Sql)
        
            Sql = "SELECT rcontrol_plagas.codvarie, rcontrol_plagas.fechacla, rcontrol_plagas.codsocio, rcontrol_plagas.codcampo, rcontrol_plagas.nroclasif, "
            Sql = Sql & "rplagasaux.nomplaga, rcontrol_plagas.kilosplaga1, rcontrol_plagas.kilosplaga2, rcontrol_plagas.kilosplaga3, rcontrol_plagas.kilosplaga4, "
            Sql = Sql & "rcontrol_plagas.kilosplaga5, rcontrol_plagas.kilosplaga6, rcontrol_plagas.kilosplaga7, rcontrol_plagas.kilosplaga8, "
            Sql = Sql & "rcontrol_plagas.kilosplaga9, rcontrol_plagas.kilosplaga10, rcontrol_plagas.kilosplaga11, rcontrol_plagas.idplaga,"
            Sql = Sql & "(kilosplaga1+kilosplaga2+kilosplaga3+kilosplaga4+kilosplaga5+kilosplaga6+kilosplaga7+kilosplaga8+kilosplaga9+kilosplaga10+kilosplaga11) total, "
            If KilosTot <> 0 Then
                Sql = Sql & " round((kilosplaga1+kilosplaga2+kilosplaga3+kilosplaga4+kilosplaga5+kilosplaga6+kilosplaga7+kilosplaga8+kilosplaga9+kilosplaga10+kilosplaga11) * 100 / " & DBSet(KilosTot, "N") & ",2) "
            Else
                Sql = Sql & "0 "
            End If
            Sql = Sql & ", ordinal "
            Sql = Sql & " from rcontrol_plagas, rplagasaux "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE rcontrol_plagas.codvarie = -1"
            End If
            Sql = Sql & " and rcontrol_plagas.idplaga = rplagasaux.idplaga "
               
    End Select
    
    MontaSQLCarga = Sql
End Function


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabem quins camps son els que mos torna
        'Creem una cadena consulta i posem els datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 5)
        CadB = Aux
        Aux = " and " & ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 1)
        CadB = CadB & Aux
        Aux = " and " & ValorDevueltoFormGrid(Text1(4), CadenaDevuelta, 2)
        CadB = CadB & Aux
        Aux = " and " & ValorDevueltoFormGrid(Text1(3), CadenaDevuelta, 3)
        CadB = CadB & Aux
        Aux = " and " & ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 4)
        CadB = CadB & Aux
        
        '   Com la clau principal es única, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.imgFec(0).Tag)
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Calidad
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codcalid
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmCam_DatoSeleccionado(CadenaSeleccion As String)
'Campos
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codcampo
End Sub

Private Sub frmcap_DatoSeleccionado(CadenaSeleccion As String)
'Capataces
    Text1(6).Text = RecuperaValor(CadenaSeleccion, 1) 'codcapat
    Text2(6).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
'Incidencias
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codincid
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    PonerDatosCampo Text1(5)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Socios
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    Text2(4).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmTar_DatoSeleccionado(CadenaSeleccion As String)
'Tarifas
    Text1(8).Text = RecuperaValor(CadenaSeleccion, 1) 'codtarifa
    Text2(8).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
'Transportistas
    Text1(7).Text = RecuperaValor(CadenaSeleccion, 1) 'codtranspor
    Text2(7).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Variedades
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(20).Text = vCampo
End Sub

Private Sub imgFec_Click(Index As Integer)
   
   Screen.MousePointer = vbHourglass
   
   Dim esq As Long
   Dim dalt As Long
   Dim menu As Long
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
    
   menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

   frmC.Left = esq + imgFec(Index).Parent.Left + 30
   frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

   
   frmC.NovaData = Now
   Select Case Index
        Case 0
            indice = 1
        Case 1
            indice = 13
   End Select
   
   Me.imgFec(0).Tag = indice
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmC.NovaData = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmC.Show vbModal
   Set frmC = Nothing
   PonerFoco Text1(indice)

End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 20
        frmZ.pTitulo = "Observaciones de la Clasificación"
        frmZ.pValor = Text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(indice)
    End If
End Sub

Private Sub mnBuscar_Click()
Dim i As Integer
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnGastos_Click()
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonGastos
End Sub

Private Sub mnImprimir_Click()

    AbrirListado (32)

'    If Data1.Recordset.EOF Then
'        MsgBox "Debe seleccionar un registro.", vbExclamation
'    Else
'        If CargarTemporalDatosDestrio Then
'            InicializarVbles
'
'            'Añadir el parametro de Empresa
'            cadparam = cadparam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
'            numParam = numParam + 1
'
'           If HayRegParaInforme("tmpexcel", "codusu = " & vUsu.Codigo) Then
'               cadNombreRPT = "rControlDestrio.rpt"
'               cadTitulo = "Resumen Control Destrio"
'
'               cadFormula = "{tmpexcel.codusu} = " & vUsu.Codigo
'
'               LlamarImprimir
'           End If
'
'        End If
'    End If
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 3  'Búscar
           mnBuscar_Click
        Case 4  'Tots
            mnVerTodos_Click
        Case 7  'Nou
            mnNuevo_Click
        Case 8  'Modificar
            mnModificar_Click
        Case 9  'Borrar
            mnEliminar_Click
        Case 12 'Imprimir
            mnImprimir_Click
        Case 14   'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(3) ' <===
        Text1(3).BackColor = vbYellow ' <===
        ' *** si n'hi han combos a la capçalera ***
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()
    
'    If Text1(2).Text <> "" Then
'        Text1(21).Text = Text1(2).Text
'        Text1(21).Tag = Replace(Text1(21).Tag, "FH", "FHH")
'    End If

    CadB = ObtenerBusqueda2(Me, 1, 2, "Frame2")
    
'    Text1(21).Tag = Replace(Text1(21).Tag, "FHH", "FH")
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
    Dim Cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    Cad = ""
    Cad = Cad & "Fecha|fechacla|T||14·"  ' ParaGrid(Text1(1), 10, "Fecha")
    Cad = Cad & "Cod|rcontrol.codsocio|T||10·" 'ParaGrid(Text1(4), 6, "Socio")
    Cad = Cad & "Socio|nomsocio|T||30·"
    Cad = Cad & "Cod|rcontrol.codvarie|T||10·" 'ParaGrid(Text1(3), 6, "Variedad")
    Cad = Cad & "Variedad|nomvarie|T||20·"
    Cad = Cad & "Campo|codcampo|T||10·" 'ParaGrid(Text1(2), 6, "Campo")
    Cad = Cad & "Nro|nroclasif|T||6·" 'ParaGrid(Text1(0), 6, "Nro.Clasif")
    
    If Cad <> "" Then
        
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        Cad = "(" & NombreTabla & " inner join variedades on rcontrol.codvarie = variedades.codvarie) inner join rsocios on rcontrol.codsocio = rsocios.codsocio "
        frmB.vTabla = Cad 'NombreTabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|1|3|5|6|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Control Destrio" ' ***** repasa açò: títol de BuscaGrid *****
        frmB.vSelElem = 1

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha posat valors i tenim que es formulari de búsqueda llavors
        'tindrem que tancar el form llançant l'event
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha retornat datos, es a decir NO ha retornat datos
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            Cad = Cad & Text1(J).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonVerTodos()
'Vore tots
    LimpiarCampos 'Neteja els Text1
    CadB = ""
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
Dim i As Integer

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    text1(0).Text = SugerirCodigoSiguienteStr("forfaits", "codforfait")
'    FormateaCampo text1(0)
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = ""
    End If
    '********************************************************************
    For i = 0 To Text3.Count - 1
        Text3(i).Text = ""
    Next i
       
    PonerFoco Text1(3) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions

End Sub

Private Sub BotonModificar()

    PonerModo 4
    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
'    BloquearTxt Text1(0), True
    CalcularDestrio
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(5)
End Sub

Private Sub BotonEliminar()
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    Cad = "¿Seguro que desea eliminar el control destrio?"
    Cad = Cad & vbCrLf & "Variedad: " & Data1.Recordset.Fields(0)
    
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Proveedor", Err.Description
End Sub

Private Sub BotonGastos()
Dim i As Integer

    Gastos = True

    '[Monica]20/07/2010 inicializada variable variedad anterior pq le damos a modificar
    VarieAnt = Text1(3).Text
    '
    
    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    For i = 0 To 13
        BloquearTxt Text1(i), True
    Next i
    BloquearTxt Text1(20), True
    imgFec(0).Enabled = False
    imgFec(1).Enabled = False
    For i = 0 To 5
        BloquearImage imgBuscar(i), True
    Next i
    
    For i = 14 To 16
        BloquearTxt Text1(i), True
        Text1(i).Enabled = False
    Next i
    
    
    ' desbloqueamos el frame de gastos
    
    For i = 17 To 19
        BloquearTxt Text1(i), False
        Text1(i).Enabled = True
    Next i
    
    
End Sub




Private Sub PonerCampos()
Dim i As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    
    
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For i = 0 To 0
        CargaGrid i, True
        If Not AdoAux(i).Recordset.EOF Then _
            PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
    Next i

    
    ' ************* configurar els camps de les descripcions de la capçalera *************
    Text2(3).Text = PonerNombreDeCod(Text1(3), "variedades", "nomvarie")
    Text2(4).Text = PonerNombreDeCod(Text1(4), "rsocios", "nomsocio")
    ' ********************************************************************************
    
    CalcularTotales
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer
Dim V

    Select Case Modo
        Case 1, 3 'Búsqueda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                ' *** foco al primer camp visible de la capçalera ***
                PonerFoco Text1(0)

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la capçalera ***
                PonerFoco Text1(0)
        
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = AdoAux(NumTabMto).Recordset.Fields(3) 'el 2 es el nº de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(3).Name & " =" & V)
                        ' ***************************************************************
                    End If
            End Select
            
            PosicionarData
            
            ' *** si n'hi han llínies en grids i camps fora d'estos ***
            If Not AdoAux(NumTabMto).Recordset.EOF Then
                DataGridAux_RowColChange NumTabMto, 1, 1
            Else
                LimpiarCamposFrame NumTabMto
            End If
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim NRegs As Integer
Dim Sql As String

    On Error GoTo EDatosOK

    DatosOk = False
    
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    If b Then
        Sql = "select count(*) from rcampos where codvarie=" & DBSet(Text1(3).Text, "N") & " and codsocio=" & DBSet(Text1(4).Text, "N")
        Sql = Sql & " and nrocampo= " & DBSet(Text1(2).Text, "N")
        Sql = Sql & " and fecbajas is null "
        NRegs = TotalRegistros(Sql)
        If NRegs = 0 Then
            MsgBox "No existen campos con variedad, socio y nro.orden dados de alta. Revise.", vbExclamation
            PonerFoco Text1(3)
        Else
            If NRegs > 1 Then
                MsgBox "Existen más de un registro con variedad, socio y nro.orden dados de alta. Revise.", vbExclamation
                PonerFoco Text1(3)
            End If
        End If
    End If
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    Cad = "(nroclasif=" & DBSet(Text1(0).Text, "N")
    Cad = Cad & " and codvarie = " & DBSet(Text1(3).Text, "N")
    Cad = Cad & " and codsocio = " & DBSet(Text1(4).Text, "N")
    Cad = Cad & " and fechacla = " & DBSet(Text1(1).Text, "F")
    Cad = Cad & " and codcampo = " & DBSet(Text1(2).Text, "N")
    Cad = Cad & " and ordinal = " & DBSet(Text1(7).Text, "N") & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    If SituarDataMULTI(Data1, Cad, Indicador) Then
    'If SituarData(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE nroclasif=" & Data1.Recordset!nroclasif
    vWhere = vWhere & " and codvarie = " & Data1.Recordset!codvarie
    vWhere = vWhere & " and codsocio = " & Data1.Recordset!Codsocio
    vWhere = vWhere & " and fechacla = " & DBSet(Data1.Recordset!fechacla, "F")
    vWhere = vWhere & " and codcampo = " & Data1.Recordset!codcampo
    vWhere = vWhere & " and ordinal  = " & Data1.Recordset!Ordinal
    
    
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM rcontrol_plagas " & vWhere
        
    'Eliminar la CAPÇALERA
    conn.Execute "Delete from " & NombreTabla & vWhere
       
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        eliminar = False
    Else
        conn.CommitTrans
        eliminar = True
    End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim Sql As String
Dim NRegs As Integer

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 0 'numero de nota
            PonerFormatoEntero Text1(Index)
        
        Case 3 'Variedad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "variedades", "nomvarie")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmVar = New frmComVar
                        frmVar.DatosADevolverBusqueda = "0|1|"
'                        frmVar.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmVar.Show vbModal
                        Set frmVar = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                Else
                    If (Modo = 3 Or Modo = 4) And EsVariedadGrupo6(Text1(Index).Text) Then
                        MsgBox "Esta variedad es del Grupo de Bodega. Revise.", vbExclamation
                        PonerFoco Text1(Index)
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 4 'Socio
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Socio: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSoc = New frmManSocios
                        frmSoc.DatosADevolverBusqueda = "0|1|"
'                        frmSoc.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmSoc.Show vbModal
                        Set frmSoc = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                Else
                    If EstaSocioDeAlta(Text1(Index)) Then
'                        PonerCamposSocioVariedad
                    Else
                        MsgBox "El socio está dado de baja. Reintroduzca.", vbExclamation
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
                
        Case 1
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            PonerFormatoFecha Text1(1)
        
        Case 2 'campo
            PonerFormatoEntero Text1(Index)
    
        Case 5 'kilos manuales
            If PonerFormatoEntero(Text1(5)) And Modo = 4 Then Me.CmdAceptar.SetFocus
            
        Case 6 ' porcentaje de destrio
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal Text1(6), 4
            
    End Select
    ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 3: KEYBusqueda KeyAscii, 0 'variedad
                Case 4: KEYBusqueda KeyAscii, 1 'socio
                Case 1: KEYFecha KeyAscii, 0 'fecha
            End Select
        End If
    Else
'        If Index <> 3 Or (Index = 3 And Text1(3).Text = "") Then KEYpress KeyAscii
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then Unload Me
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub



'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    Select Case Button.Index
        Case 2
            BotonModificarLinea Index
        Case Else
    End Select
    'End If
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar llínia
       
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
  
    Select Case Index
        Case 0 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, i
                DataGridAux(Index).Refresh
            End If
              
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
    End Select
    
    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 0 ' muestra
        
            txtAux(0).Text = DataGridAux(Index).Columns(4).Text 'nroclasif
            txtAux(1).Text = DataGridAux(Index).Columns(0).Text
            txtAux(2).Text = DataGridAux(Index).Columns(2).Text
            txtAux(3).Text = DataGridAux(Index).Columns(1).Text
            txtAux(4).Text = DataGridAux(Index).Columns(3).Text
            txtAux(16).Text = DataGridAux(Index).Columns(17).Text
            txtAux(17).Text = DataGridAux(Index).Columns(18).Text
            For i = 5 To 15
                txtAux(i).Text = DataGridAux(Index).Columns(i + 1).Text
            Next i
            
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'muestras
            PonerFoco txtAux(5)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'muestras
             For jj = 5 To 15
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
    End Select
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15 ' kilosmuestra
            PonerFormatoEntero txtAux(Index)
            
            If Index = 15 And PonerFormatoEntero(txtAux(Index)) Then
                CmdAceptar.SetFocus
            End If
    End Select
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
   If Not txtAux(Index).MultiLine Then ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
   KEYpress KeyAscii
End Sub

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    ' ******************************************************************************
    DatosOkLlin = b
    
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    ' *** si cal comprovar alguna cosa abans de borrar ***
'    Select Case Index
'        Case 0 'cuentas bancarias
'            If AdoAux(Index).Recordset!ctaprpal = 1 Then
'                MsgBox "No puede borrar una Cuenta Principal. Seleccione antes otra cuenta como Principal", vbExclamation
'                Exit Function
'            End If
'    End Select
    ' ****************************************************
    
    SepuedeBorrar = True
End Function

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    indice = Index + 3
     Select Case Index
        Case 0 'variedades
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = Text1(3).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(3)
        Case 1 'socios
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
'            frmSoc.CodigoActual = Text1(4).Text
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(4)
        Case 2 'campos
            Set frmCam = New frmManCampos
            frmCam.DatosADevolverBusqueda = "0|"
'            frmCam.CodigoActual = Text1(5).Text
            frmCam.Show vbModal
            Set frmCam = Nothing
            PonerFoco Text1(5)
        Case 3 'Capataces
            Set frmCap = New frmManCapataz
            frmCap.DatosADevolverBusqueda = "0|1|"
            frmCap.CodigoActual = Text1(6).Text
            frmCap.Show vbModal
            Set frmCap = Nothing
            PonerFoco Text1(6)
        Case 4 'Transportista
            Set frmTra = New frmManTranspor
            frmTra.DatosADevolverBusqueda = "0|1|"
            frmTra.CodigoActual = Text1(7).Text
            frmTra.Show vbModal
            Set frmTra = Nothing
            PonerFoco Text1(7)
        Case 5 'Tarifa
            Set frmTar = New frmManTarTra
            frmTar.DatosADevolverBusqueda = "0|1|"
            frmTar.CodigoActual = Text1(8).Text
            frmTar.Show vbModal
            Set frmTar = Nothing
            PonerFoco Text1(8)
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If ModoLineas <> 1 Then
        Select Case Index
            Case 0 'cuentas bancarias
                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux(11).Text = DataGridAux(Index).Columns("direccio").Text
'                    txtAux(12).Text = DataGridAux(Index).Columns("observac").Text
                End If
                
            Case 1 'departamentos
                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux(21).Text = DataGridAux(Index).Columns(5).Text
'                    txtAux(22).Text = DataGridAux(Index).Columns(6).Text
'                    txtAux(23).Text = DataGridAux(Index).Columns(8).Text
'                    txtAux(24).Text = DataGridAux(Index).Columns(15).Text
'                    txtAux2(22).Text = DataGridAux(Index).Columns(7).Text
                End If
                
        End Select
        
    Else 'vamos a Insertar
        Select Case Index
            Case 0 'cuentas bancarias
'                txtAux(11).Text = ""
'                txtAux(12).Text = ""
            Case 1 'departamentos
                For i = 21 To 24
'                   txtAux(i).Text = ""
                Next i
'               txtAux2(22).Text = ""
            Case 2 'Tarjetas
'               txtAux(50).Text = ""
'               txtAux(51).Text = ""
        End Select
    End If
End Sub

' ***** si n'hi han varios nivells de tabs *****
'Private Sub SituarTab(numTab As Integer)
'    On Error Resume Next
'
'    SSTab1.Tab = numTab
'
'    If Err.Number <> 0 Then Err.Clear
'End Sub
' **********************************************

Private Sub CargaFrame(Index As Integer, enlaza As Boolean)
Dim tip As Integer
Dim i As Byte

    AdoAux(Index).ConnectionString = conn
    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    AdoAux(Index).Refresh
    
    If Not AdoAux(Index).Recordset.EOF Then
        PonerCamposForma2 Me, AdoAux(Index), 2, "FrameAux" & Index
    Else
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        NetejaFrameAux "FrameAux3" 'neteja només lo que te TAG
    End If
End Sub

' *** si n'hi han tabs sense datagrids ***
Private Sub NetejaFrameAux(nom_frame As String)
Dim Control As Object
    
    For Each Control In Me.Controls
        If (Control.Tag <> "") Then
            If (Control.Container.Name = nom_frame) Then
                If TypeOf Control Is TextBox Then
                    Control.Text = ""
                ElseIf TypeOf Control Is ComboBox Then
                    Control.ListIndex = -1
                End If
            End If
        End If
    Next Control

End Sub

Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'controldestrio_plagas
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;" 'numnotac
            tots = tots & "S|txtAux2(0)|T|Tipo|800|;"
            tots = tots & "S|txtAux(5)|T|Kilos 1|800|;S|txtAux(6)|T|Kilos 2|800|;S|txtAux(7)|T|Kilos 3|800|;"
            tots = tots & "S|txtAux(8)|T|Kilos 4|800|;S|txtAux(9)|T|Kilos 5|800|;S|txtAux(10)|T|Kilos 6|800|;"
            tots = tots & "S|txtAux(11)|T|Kilos 7|800|;S|txtAux(12)|T|Kilos 8|800|;S|txtAux(13)|T|Kilos 9|800|;"
            tots = tots & "S|txtAux(14)|T|Kilos 10|800|;S|txtAux(15)|T|Kilos 11|800|;N||||0|;S|txtAux(17)|T|Total|800|;S|txtAux(18)|T|Porcen|800|;N|txtAux(19)|T|Ordinal|800|;"

            arregla tots, DataGridAux(Index), Me
            
            DataGridAux(0).Columns(6).Alignment = dbgRight
            DataGridAux(0).Columns(7).Alignment = dbgRight
            DataGridAux(0).Columns(8).Alignment = dbgRight
            DataGridAux(0).Columns(9).Alignment = dbgRight
            DataGridAux(0).Columns(10).Alignment = dbgRight
            DataGridAux(0).Columns(11).Alignment = dbgRight
            DataGridAux(0).Columns(12).Alignment = dbgRight
            DataGridAux(0).Columns(13).Alignment = dbgRight
            DataGridAux(0).Columns(14).Alignment = dbgRight
            DataGridAux(0).Columns(15).Alignment = dbgRight
            DataGridAux(0).Columns(16).Alignment = dbgRight
            DataGridAux(0).Columns(18).Alignment = dbgRight
            DataGridAux(0).Columns(19).Alignment = dbgRight
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
            
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
'    If Not AdoAux(Index).Recordset.EOF Then
'        DataGridAux_RowColChange Index, 1, 1
'    Else
''        LimpiarCamposFrame Index
'    End If
    CalcularTotales
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub


Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
    
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'muestras
    End Select
    ModificarLinea = False
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ModoLineas = 0
            Select Case NumTabMto
                Case 0
                    V = AdoAux(NumTabMto).Recordset.Fields(17) 'el 2 es el nº de llinia
            End Select
            CargaGrid NumTabMto, True
            
            ' *** si n'hi han tabs ***
'            SituarTab (NumTabMto + 1)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            PonerFocoGrid Me.DataGridAux(NumTabMto)
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(17).Name & " =" & V)
            
            LLamaLineas NumTabMto, 0
            ModificarLinea = True
        End If
    End If
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codvarie=" & Me.Data1.Recordset!codvarie
    vWhere = vWhere & " and fechacla=" & DBSet(Me.Data1.Recordset!fechacla, "F")
    vWhere = vWhere & " and codsocio=" & Me.Data1.Recordset!Codsocio
    vWhere = vWhere & " and codcampo=" & Me.Data1.Recordset!codcampo
    vWhere = vWhere & " and nroclasif=" & Me.Data1.Recordset!nroclasif
    vWhere = vWhere & " and ordinal=" & Me.Data1.Recordset!Ordinal
    
    ObtenerWhereCab = vWhere
End Function

Private Function ObtenerWhereCab2(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codvarie=" & Me.Data1.Recordset!codvarie
    vWhere = vWhere & " and fechacla=" & DBSet(Me.Data1.Recordset!fechacla, "F")
    vWhere = vWhere & " and codsocio=" & Me.Data1.Recordset!Codsocio
    vWhere = vWhere & " and codcampo=" & Me.Data1.Recordset!codcampo
    vWhere = vWhere & " and nroclasif= 1 " ' es en la primera clasificacion donde vamos a grabar las plagas
    vWhere = vWhere & " and ordinal=" & Me.Data1.Recordset!Ordinal
    
    ObtenerWhereCab2 = vWhere
End Function

'' *** neteja els camps dels tabs de grid que
''estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
    On Error Resume Next
 
'    Select Case Index
'        Case 0 'Cuentas Bancarias
'            txtAux(11).Text = ""
'            txtAux(12).Text = ""
'        Case 1 'Departamentos
'            txtAux(21).Text = ""
'            txtAux(22).Text = ""
'            txtAux2(22).Text = ""
'            txtAux(23).Text = ""
'            txtAux(24).Text = ""
'        Case 2 'Tarjetas
'            txtAux(50).Text = ""
'            txtAux(51).Text = ""
'        Case 4 'comisiones
'            txtAux2(2).Text = ""
'    End Select
'
    If Err.Number <> 0 Then Err.Clear
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub


Private Sub CalcularTotales()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql2 As String
Dim TotalEnvases As String
Dim TotalCostes As String
Dim total As Long
Dim Valor As Currency
Dim i As Integer

    On Error Resume Next

    If Data1.Recordset.EOF Or Modo = 1 Then
        For i = 0 To 10
            Text3(i).Text = ""
        Next i
        Exit Sub
    End If

    Sql = "select sum(kilosplaga1) plaga1, sum(kilosplaga2) plaga2, sum(kilosplaga3) plaga3, sum(kilosplaga4) plaga4, "
    Sql = Sql & " sum(kilosplaga5) plaga5, sum(kilosplaga6) plaga6, sum(kilosplaga7) plaga7, sum(kilosplaga8) plaga8, "
    Sql = Sql & " sum(kilosplaga9) plaga9, sum(kilosplaga10) plaga10, sum(kilosplaga11) plaga11 "
    Sql = Sql & " from rcontrol_plagas "
    Sql = Sql & " where codvarie = " & Data1.Recordset!codvarie
    Sql = Sql & " and codsocio = " & Data1.Recordset!Codsocio
    Sql = Sql & " and codcampo = " & Data1.Recordset!codcampo
    Sql = Sql & " and fechacla = " & DBSet(Data1.Recordset!fechacla, "F")
    Sql = Sql & " and nroclasif = " & Data1.Recordset!nroclasif
    Sql = Sql & " and rcontrol_plagas.idplaga <> 2"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    total = 0
    If Not Rs.EOF Then
        For i = 0 To 10
            Text3(i).Text = DBLet(Rs.Fields(i).Value, "N")
            Text3(i).Text = Format(Text3(i).Text, "###,##0")
            total = total + DBLet(Rs.Fields(i).Value, "N")
        Next i
    End If
    Text3(11).Text = Format(total, "###,##0")
    Text3(12).Text = 100
    If Err.Number <> 0 Then
        Err.Clear
    End If

End Sub

'Private Function HorasDecimal(cantidad As String) As Currency
'Dim Entero As Long
'Dim vCantidad As String
'Dim vDecimal As String
'Dim vEntero As String
'Dim vHoras As Currency
'Dim J As Integer
'    HorasDecimal = 0
'
'    vCantidad = ImporteSinFormato(cantidad)
'
'    J = InStr(1, vCantidad, ",")
'
'    If J > 0 Then
'        vEntero = Mid(vCantidad, 1, J - 1)
'        vDecimal = Mid(vCantidad, J + 1, Len(vCantidad))
'    Else
'        vEntero = vCantidad
'        vDecimal = ""
'    End If
'
'    vHoras = (CLng(vEntero) * 60) + CLng(vDecimal)
'
'    HorasDecimal = Round2(vHoras / 60, 2)
'
'End Function


Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If Me.DataGridAux(0).Bookmark < Me.AdoAux(0).Recordset.RecordCount Then
'        DataGridAux(0).Row = DataGridAux(0).Row + 1
        DataGridAux(0).Bookmark = DataGridAux(0).Bookmark + 1
        BotonModificarLinea 0
    ElseIf DataGridAux(0).Bookmark = AdoAux(0).Recordset.RecordCount Then
'        PonerFocoBtn Me.cmdAceptar
        BotonModificarLinea 0
    End If
End Sub


Private Sub PasarAntReg()
'Nos situamos en el siguiente registro
    If Me.DataGridAux(0).Bookmark > 1 Then
'        DataGridAux(0).Row = DataGridAux(0).Row + 1
        DataGridAux(0).Bookmark = DataGridAux(0).Bookmark - 1
        BotonModificarLinea 0
    ElseIf DataGridAux(0).Bookmark = 1 Then
'        PonerFocoBtn Me.cmdAceptar
        BotonModificarLinea 0
    End If
End Sub



Private Sub VisualizarDatosCampo(Campo As String)
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If Campo = "" Then Exit Sub
    
'    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    Cad = "rcampos.codcampo = " & DBSet(Campo, "N")
     
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rpartida.codzonas, rzonas.nomzonas, "
    Cad1 = Cad1 & " rpueblos.despobla from rcampos, rpartida, rzonas, rpueblos "
    Cad1 = Cad1 & " where " & Cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codzonas = rzonas.codzonas "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla "
     
    Set Rs = New ADODB.Recordset
    Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Text1(5).Text = Campo
        PonerFormatoEntero Text1(5)
        Text2(0).Text = DBLet(Rs.Fields(1).Value, "T") ' nombre de partida
        Text2(1).Text = DBLet(Rs!desPobla, "T")        ' nombre de la poblacion
    End If
    
    Set Rs = Nothing
    
End Sub


Private Sub PonerCamposSocioVariedad()
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If Text1(3).Text = "" Or Text1(4).Text = "" Then Exit Sub
    
    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    Cad = "rcampos.codsocio = " & DBSet(Text1(4).Text, "N") & " and rcampos.fecbajas is null"
    Cad = Cad & " and rcampos.codvarie = " & DBSet(Text1(3).Text, "N")
     
    Cad1 = "select count(*) from rcampos where " & Cad
     
    NumRegis = TotalRegistros(Cad1)
    
    If NumRegis = 0 Then Exit Sub
    If NumRegis = 1 Then
        Cad1 = "select codcampo from rcampos where " & Cad
        Set Rs = New ADODB.Recordset
        Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            Text1(5).Text = DBLet(Rs.Fields(0).Value)
            PonerDatosCampo Text1(5).Text
        End If
    Else
        Set frmMens = New frmMensajes
        frmMens.cadwhere = " and " & Cad '"rcampos.codsocio = " & NumCod & " and rcampos.fecbajas is null"
        frmMens.Campo = Text1(5).Text
        frmMens.OpcionMensaje = 6
        frmMens.Show vbModal
        Set frmMens = Nothing
    End If
    
End Sub

Private Sub PonerDatosCampo(Campo As String)
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If Campo = "" Then Exit Sub
    
'    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    Cad = "rcampos.codcampo = " & DBSet(Campo, "N") & " and rcampos.fecbajas is null"
     
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rpartida.codzonas, rzonas.nomzonas, "
    Cad1 = Cad1 & " rpueblos.despobla from rcampos, rpartida, rzonas, rpueblos "
    Cad1 = Cad1 & " where " & Cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codzonas = rzonas.codzonas "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla "
     
    Set Rs = New ADODB.Recordset
    Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Text1(5).Text = ""
    Text2(0).Text = ""
    Text2(1).Text = ""
    If Not Rs.EOF Then
        Text1(5).Text = Campo
        PonerFormatoEntero Text1(5)
        Text2(0).Text = DBLet(Rs.Fields(1).Value, "T") ' nombre de partida
        Text2(1).Text = DBLet(Rs.Fields(4).Value, "T") ' descripcion de poblacion
    End If
    
    Set Rs = Nothing
    
End Sub

Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String
Dim actualiza As Boolean
Dim NumF As Long

    Sql = "select max(ordinal) from rcontrol where nroclasif = " & DBSet(Text1(0).Text, "N")
    Sql = Sql & " and codvarie = " & DBSet(Text1(3).Text, "N")
    Sql = Sql & " and codsocio = " & DBSet(Text1(4).Text, "N")
    Sql = Sql & " and codcampo = " & DBSet(Text1(2).Text, "N")
    Sql = Sql & " and fechacla = " & DBSet(Text1(1).Text, "F")
    NumF = DevuelveValor(Sql)
    
    Text1(7).Text = NumF + 1
        
    Sql = CadenaInsertarDesdeForm(Me)
    If InsertarOferta(Sql, vTipoMov, actualiza) Then
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
        PonerModo 2
    
'        If Not NotaExistente Then
'            Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
'            PosicionarData
'            BotonModificar
'            cmdAceptar_Click
'        End If
    
    End If
    Text1(0).Text = Format(Text1(0).Text, "0000000")
End Sub


Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov, ActualizarContador As Boolean) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String
Dim Sql2 As String

Dim Rs As ADODB.Recordset
Dim Sql3 As String
Dim cadMen As String

    On Error GoTo EInsertarOferta
    
    bol = True
    
    vSQL = CadenaInsertarDesdeForm(Me)
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Factura (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    cadMen = ""
    Sql3 = "select * from rcontrol where codvarie = " & DBSet(Text1(3).Text, "N")
    Sql3 = Sql3 & " and codsocio = " & DBSet(Text1(4).Text, "N")
    Sql3 = Sql3 & " and fechacla = " & DBSet(Text1(1).Text, "F")
    Sql3 = Sql3 & " and codcampo = " & DBSet(Text1(2).Text, "N")
    Sql3 = Sql3 & " and nroclasif = " & DBSet(Text1(0).Text, "N")
    Sql3 = Sql3 & " and ordinal = " & DBSet(Text1(7).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        bol = InsertarLineasPlagas(Rs, cadMen)
        cadMen = "Insertando Lineas: " & cadMen
    End If
    
    Set Rs = Nothing
    
    MenError = MenError & cadMen
    
EInsertarOferta:
    If Err.Number <> 0 Or Not bol Then
        MenError = "Insertando Control Destrio." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        InsertarOferta = True
    Else
        conn.RollbackTrans
        InsertarOferta = False
    End If
End Function

Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = " codvarie= " & Text1(3).Text
    Sql = Sql & " and codsocio= " & Text1(4).Text
    Sql = Sql & " and fechacla = " & DBSet(Text1(1).Text, "F")
    Sql = Sql & " and codcampo = " & DBSet(Text1(2).Text, "N")
    Sql = Sql & " and nroclasif = " & DBSet(Text1(0).Text, "N")
    Sql = Sql & " and ordinal = " & DBSet(Text1(7).Text, "N")
    
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function

Private Function InsertarLineasPlagas(ByRef Rs As ADODB.Recordset, cadErr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Sql1 As String
Dim RS1 As ADODB.Recordset
Dim Cad As String
Dim KilosMuestra As Currency
Dim TotalKilos As Currency
Dim Calidad As Currency
Dim Diferencia As Currency
Dim HayReg As Byte
Dim TipoClasif As Byte
Dim vTipoClasif As String
Dim vCalidDest As String
Dim CalidadClasif As String
Dim CalidadVC As String
Dim i As Integer

    On Error GoTo EInsertar
    
    Sql1 = "insert into rcontrol_plagas (codvarie, fechacla, codsocio, codcampo, "
    Sql1 = Sql1 & " nroclasif, idplaga, kilosplaga1, kilosplaga2, kilosplaga3, kilosplaga4, kilosplaga5,"
    Sql1 = Sql1 & "kilosplaga6, kilosplaga7, kilosplaga8, kilosplaga9, kilosplaga10, kilosplaga11, ordinal) values "
    
                    '[Monica]30/01/2014: antes no estaba clareta (16)
    For i = 1 To 16 '15 '14 [Monica]25/11/2011: antes no estaba la palga pixat (15)
        Sql1 = Sql1 & "(" & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!fechacla, "F") & "," & DBSet(Rs!Codsocio, "N") & ","
        Sql1 = Sql1 & DBSet(Rs!codcampo, "N") & "," & DBSet(Rs!nroclasif, "N") & "," & DBSet(i, "N") & ","
        Sql1 = Sql1 & "0,0,0,0,0,0,0,0,0,0,0, " & DBSet(Text1(7).Text, "N") & "),"
    Next i
    
    ' quitamos la ultima coma
    Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
    
    conn.Execute Sql1
    InsertarLineasPlagas = True
    Exit Function
        
EInsertar:
    If Err.Number <> 0 Then
        InsertarLineasPlagas = False
        cadErr = Err.Description
    Else
        InsertarLineasPlagas = True
    End If
End Function

Private Function ModificaCabecera() As Boolean
Dim b As Boolean
Dim MenError As String
Dim Sql As String

    On Error GoTo EModificarCab

    conn.BeginTrans
    
    b = True
    
    If b Then b = ModificaDesdeFormulario1(Me, 1) 'ModificaDesdeFormulario2(Me, 2, "Frame2")

    If b Then
        MenError = "Actualizar Destrion Clasificación Automática"
        b = ActualizarDestrioClasAuto(MenError)
    End If
    
    If b Then
        MenError = "Insertar Plagas en Clasificación"
        b = InsertaPlagasClasAuto(MenError)
    End If
    
EModificarCab:
    If Err.Number <> 0 Or Not b Then
        MenError = "Modificando Control Destrio." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        b = False
    End If
    If b Then
        ModificaCabecera = True
        conn.CommitTrans
    Else
        ModificaCabecera = False
        conn.RollbackTrans
    End If
End Function



Private Sub CalcularDestrio()
Dim Sql As String
Dim PorcDest As Currency
Dim KilosTot As Long


    Sql = "select sum(kilosplaga1+kilosplaga2+kilosplaga3+kilosplaga4+kilosplaga5+kilosplaga6+kilosplaga7+kilosplaga8+kilosplaga9+kilosplaga10+kilosplaga11) total "
    Sql = Sql & " from rcontrol_plagas "
    Sql = Sql & ObtenerWhereCab(True)
    Sql = Sql & " and idplaga <> 2 "
    KilosTot = DevuelveValor(Sql)

    Sql = "SELECT "
    If KilosTot <> 0 Then
        Sql = Sql & " round((kilosplaga1+kilosplaga2+kilosplaga3+kilosplaga4+kilosplaga5+kilosplaga6+kilosplaga7+kilosplaga8+kilosplaga9+kilosplaga10+kilosplaga11) * 100 / " & DBSet(KilosTot, "N") & ",2) "
    Else
        Sql = Sql & "0 "
    End If
    Sql = Sql & " from rcontrol_plagas "
    Sql = Sql & ObtenerWhereCab(True)
    Sql = Sql & " and rcontrol_plagas.idplaga = 1 "

    PorcDest = Round2(100 - DevuelveValor(Sql), 2)
    
    Text1(6).Text = PorcDest
    Text1(6).Text = Format(Text1(6).Text, "##0.00")
    

End Sub

Private Function ActualizarDestrioClasAuto(Mens As String) As Boolean
Dim Sql As String
Dim KilosTot As Long
Dim i As Integer
Dim Porcen As Currency

    On Error GoTo eActualizarDestrioClasAuto
    
    ActualizarDestrioClasAuto = False

    Sql = "update rclasifauto set kilospeq = " & DBSet(Text1(5).Text, "N")
    Sql = Sql & ", porcdest = " & DBSet(Text1(6).Text, "N")
    Sql = Sql & " where numnotac = " & DBSet(Data1.Recordset!nroclasif, "N")
    Sql = Sql & " and codvarie = " & Data1.Recordset!codvarie
    Sql = Sql & " and codsocio = " & Data1.Recordset!Codsocio
    Sql = Sql & " and codcampo = " & Data1.Recordset!codcampo
    Sql = Sql & " and fechacla = " & DBSet(Data1.Recordset!fechacla, "F")
    Sql = Sql & " and ordinal = " & DBSet(Data1.Recordset!Ordinal, "N")
    
    conn.Execute Sql

    ActualizarDestrioClasAuto = True
    Exit Function

eActualizarDestrioClasAuto:
    Mens = Mens & vbCrLf & Err.Description
End Function




Private Function InsertaPlagasClasAuto(Mens As String) As Boolean
Dim Sql As String
Dim KilosTot As Long
Dim i As Integer
Dim Porcen As Currency
    On Error GoTo eInsertaPlagasClasAuto
    
    InsertaPlagasClasAuto = False

'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
    Sql = "delete from rclasifauto_plagas where " 'numnotac = " & DBSet(Data1.Recordset!nroclasif, "N")
    'SQL = SQL & " and codvarie = " & Data1.Recordset!CodVarie
    Sql = Sql & " codVarie = " & Data1.Recordset!codvarie
    Sql = Sql & " and codsocio = " & Data1.Recordset!Codsocio
    Sql = Sql & " and codcampo = " & Data1.Recordset!codcampo
    Sql = Sql & " and fechacla = " & DBSet(Data1.Recordset!fechacla, "F")
    Sql = Sql & " and ordinal = " & DBSet(Data1.Recordset!Ordinal, "N")
    
    conn.Execute Sql


    Sql = "select sum(kilosplaga1+kilosplaga2+kilosplaga3+kilosplaga4+kilosplaga5+kilosplaga6+kilosplaga7+kilosplaga8+kilosplaga9+kilosplaga10+kilosplaga11) total "
    Sql = Sql & " from rcontrol_plagas "
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'    SQL = SQL & ObtenerWhereCab(True)
    Sql = Sql & ObtenerWhereCab2(True)
    Sql = Sql & " and idplaga <> 2 "
    KilosTot = DevuelveValor(Sql)


    For i = 3 To 13
        Sql = "SELECT "
        If KilosTot <> 0 Then
            Sql = Sql & " round((kilosplaga1+kilosplaga2+kilosplaga3+kilosplaga4+kilosplaga5+kilosplaga6+kilosplaga7+kilosplaga8+kilosplaga9+kilosplaga10+kilosplaga11) * 100 / " & DBSet(KilosTot, "N") & ",2) "
        Else
            Sql = Sql & "0 "
        End If
        Sql = Sql & " from rcontrol_plagas "
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'        SQL = SQL & ObtenerWhereCab(True)
        Sql = Sql & ObtenerWhereCab2(True)
        Sql = Sql & " and rcontrol_plagas.idplaga = " & DBSet(i, "N")
    
        Porcen = DevuelveValor(Sql)
        
        Select Case i
            Case 3 ' piojo gris
                Select Case Porcen
                    Case 1 To 5
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 1"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "1)"
                            
                            conn.Execute Sql
                        End If
                        
                    Case 5.01 To 15
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 2"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "2)"
                            
                            conn.Execute Sql
                        End If
        
                    Case Is > 15
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 3"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "3)"
                            
                            conn.Execute Sql
                        End If
                End Select
            
            
            Case 4 ' piojo rojo
                Select Case Porcen
                    Case 1 To 5
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 4"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "4)"
                            
                            conn.Execute Sql
                        End If
                        
                    Case 5.01 To 15
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 5"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "5)"
                            
                            conn.Execute Sql
                        End If
        
                    Case Is > 15
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 6"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "6)"
                            
                            conn.Execute Sql
                        End If
                End Select
            
            
            Case 5 ' serpeta
                Select Case Porcen
                    Case 1 To 5
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 7"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "7)"
                            
                            conn.Execute Sql
                        End If
                        
                    Case 5.01 To 15
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 8"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "8)"
                            
                            conn.Execute Sql
                        End If
        
                    Case Is > 15
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 9"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "9)"
                            
                            conn.Execute Sql
                        End If
                End Select
            
            
            Case 6 ' araña
                Select Case Porcen
                    Case 1 To 5
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 16"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "16)"
                            
                            conn.Execute Sql
                        End If
                        
                    Case 5.01 To 15
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 17"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "17)"
                            
                            conn.Execute Sql
                        End If
        
                    Case Is > 15
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 18"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "18)"
                            
                            conn.Execute Sql
                        End If
                End Select
            
            
            Case 7 ' %piedra
                If Porcen > 1 Then
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                    SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                    Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                    Sql = Sql & " and codplaga = 22"
                    If TotalRegistros(Sql) = 0 Then
                        Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = SQL & Data1.Recordset!nroclasif & ","
                        Sql = Sql & "1,"
                        Sql = Sql & Data1.Recordset!codvarie & ","
                        Sql = Sql & Data1.Recordset!Codsocio & ","
                        Sql = Sql & Data1.Recordset!codcampo & ","
                        Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                        Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                        Sql = Sql & "22)"
                        
                        conn.Execute Sql
                    End If
                End If
                
            Case 8 ' negrita
                Select Case Porcen
                    Case 1 To 5
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 19"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "19)"
                            
                            conn.Execute Sql
                        End If
                        
                    Case 5.01 To 15
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 20"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "20)"
                            
                            conn.Execute Sql
                        End If
        
                    Case Is > 15
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 21"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "21)"
                            
                            conn.Execute Sql
                        End If
                End Select
            
            
            Case 13 ' mosca
                Select Case Porcen
                    Case 1 To 5
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 10"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "10)"
                            
                            conn.Execute Sql
                        End If
                        
                    Case 5.01 To 15
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 11"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "11)"
                            
                            conn.Execute Sql
                        End If
        
                    Case Is > 15
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                        SQL = "select count(*) from rclasifauto_plagas " & Replace(ObtenerWhereCab(True), "nroclasif", "numnotac")
                        Sql = "select count(*) from rclasifauto_plagas " & ObtenerWhereCab2(True)
                        Sql = Sql & " and codplaga = 12"
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifauto_plagas (numnotac, codvarie, codsocio, codcampo, fechacla, ordinal, codplaga) values ("
'[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
'                            SQL = SQL & Data1.Recordset!nroclasif & ","
                            Sql = Sql & "1,"
                            Sql = Sql & Data1.Recordset!codvarie & ","
                            Sql = Sql & Data1.Recordset!Codsocio & ","
                            Sql = Sql & Data1.Recordset!codcampo & ","
                            Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
                            Sql = Sql & DBSet(Data1.Recordset!Ordinal, "N") & ","
                            Sql = Sql & "12)"
                            
                            conn.Execute Sql
                        End If
                
                End Select
        End Select
    Next i

    InsertaPlagasClasAuto = True
    Exit Function


eInsertaPlagasClasAuto:
    Mens = Mens & vbCrLf & Err.Description
End Function

'Private Function CargarTemporalDatosDestrio() As Boolean
'Dim Sql As String
'Dim KilosTot As Currency
'Dim KilosMan As Currency
'Dim Rs As ADODB.Recordset
'
'    On Error GoTo eCargarTemporalDatosDestrio
'
'    CargarTemporalDatosDestrio = True
'
'    Sql = "delete from tmpexcel where codusu = " & vUsu.Codigo
'    conn.Execute Sql
'
'    Sql = "select idplaga, (sum(kilosplaga1) + sum(kilosplaga2) + sum(kilosplaga3) + sum(kilosplaga4) + "
'    Sql = Sql & " sum(kilosplaga5) + sum(kilosplaga6) + sum(kilosplaga7) + sum(kilosplaga8) + "
'    Sql = Sql & " sum(kilosplaga9) + sum(kilosplaga10) + sum(kilosplaga11)) as Total  "
'    Sql = Sql & " from rcontrol_plagas "
'    Sql = Sql & " where codvarie = " & Data1.Recordset!CodVarie
'    Sql = Sql & " and codsocio = " & Data1.Recordset!CodSocio
'    Sql = Sql & " and codcampo = " & Data1.Recordset!codcampo
'    Sql = Sql & " and fechacla = " & DBSet(Data1.Recordset!fechacla, "F")
'    Sql = Sql & " group by 1 "
'    Sql = Sql & " order by 1 "
''    Sql = Sql & " and rcontrol_plagas.idplaga <> 2"
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    Sql = "select sum(kilosman) from rcontrol "
'    Sql = Sql & " where codvarie = " & Data1.Recordset!CodVarie
'    Sql = Sql & " and codsocio = " & Data1.Recordset!CodSocio
'    Sql = Sql & " and codcampo = " & Data1.Recordset!codcampo
'    Sql = Sql & " and fechacla = " & DBSet(Data1.Recordset!fechacla, "F")
'
'    KilosMan = DevuelveValor(Sql)
'
'    KilosTot = 0
'    While Not Rs.EOF
'        Sql = "insert into tmpexcel (codusu,numalbar,fecalbar,codvarie,codsocio,codcampo,calidad1,calidad2) values ( "
'        Sql = Sql & vUsu.Codigo & ","
'        Sql = Sql & DBLet(Rs!idplaga, "N") & ","
'        Sql = Sql & DBSet(Data1.Recordset!fechacla, "F") & ","
'        Sql = Sql & Data1.Recordset!CodVarie & ","
'        Sql = Sql & Data1.Recordset!CodSocio & ","
'        Sql = Sql & Data1.Recordset!codcampo & ","
'        Sql = Sql & DBSet(Rs!total, "N") & ","
'        Sql = Sql & DBSet(KilosMan, "N") & ")"
'
'        conn.Execute Sql
'        If DBLet(Rs!idplaga, "N") <> 2 Then KilosTot = KilosTot + DBLet(Rs!total, "N")
'        Rs.MoveNext
'    Wend
'    Set Rs = Nothing
'
'    Sql = "update tmpexcel set kilosnet = " & DBSet(KilosTot, "N") & " where codusu = " & vUsu.Codigo
'    conn.Execute Sql
'
'    CargarTemporalDatosDestrio = True
'    Exit Function
'
'eCargarTemporalDatosDestrio:
'    MuestraError Err.Number, "Cargar Datos Temporal Destrio", Err.Description
'End Function
'
'Private Sub LlamarImprimir()
'    With frmImprimir
'        .FormulaSeleccion = cadFormula
'        .OtrosParametros = cadparam
'        .NumeroParametros = numParam
'        .SoloImprimir = False
'        .EnvioEMail = False
'        .Titulo = cadTitulo
'        .NombreRPT = cadNombreRPT
'        .ConSubInforme = False
'        .Opcion = 0
'        .Show vbModal
'    End With
'End Sub
'
'Private Sub InicializarVbles()
'    cadFormula = ""
'    cadSelect = ""
'    cadSelect1 = ""
'    cadparam = ""
'    numParam = 0
'End Sub
'

