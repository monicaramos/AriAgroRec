VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#17.2#0"; "CODEJO~3.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#17.2#0"; "COC9F8~1.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#17.2#0"; "COA2AE~1.OCX"
Begin VB.Form frmppal 
   Caption         =   "Ariagrorec"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11580
   FillStyle       =   0  'Solid
   Icon            =   "frmPpalN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   8880
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7532
            Key             =   "New"
            Object.Tag             =   "100"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7590
            Key             =   "Open"
            Object.Tag             =   "101"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":75EE
            Key             =   "Save"
            Object.Tag             =   "103"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":764C
            Key             =   "Print"
            Object.Tag             =   "113"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":76AA
            Key             =   "Cut"
            Object.Tag             =   "108"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7708
            Key             =   "Copy"
            Object.Tag             =   "106"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7766
            Key             =   "Paste"
            Object.Tag             =   "107"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":77C4
            Key             =   "Bold"
            Object.Tag             =   "120"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7822
            Key             =   "Italic"
            Object.Tag             =   "121"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7880
            Key             =   "Underline"
            Object.Tag             =   "122"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":78DE
            Key             =   "Align Left"
            Object.Tag             =   "123"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":793C
            Key             =   "Center"
            Object.Tag             =   "124"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":799A
            Key             =   "Align Right"
            Object.Tag             =   "125"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":79F8
            Key             =   "About"
            Object.Tag             =   "112"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7A56
            Key             =   ""
            Object.Tag             =   "166"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7AB4
            Key             =   ""
            Object.Tag             =   "168"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7B12
            Key             =   ""
            Object.Tag             =   "165"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListPPal48 
      Left            =   5280
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_OM 
      Left            =   3240
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_BN 
      Left            =   2280
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_BN16 
      Left            =   2040
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_OM16 
      Left            =   2400
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageListPpal16 
      Left            =   720
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgListComun 
      Left            =   1560
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   360
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImaListBotoneras32 
      Left            =   2400
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":E3D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":14C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":1B496
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":21CF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":2855A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":2EDBC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImaListBotoneras 
      Left            =   2880
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":3561E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":3BE80
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":426E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":48F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":4F7A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":56008
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":5C86A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":630CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImaListBotoneras_BN 
      Left            =   2760
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483626
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":63ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":6A340
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":70BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":77404
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7DC66
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":844C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":8AD2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":9158C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":97DEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":9E650
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2880
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":9F062
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":A58C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":A8076
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListDocumentos 
      Left            =   2400
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":AE8D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":AFB5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":B230C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":B4446
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":B4760
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":B7B52
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":B9764
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":BA541
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":BB4B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":BC428
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImaListBotoneras32_BN 
      Left            =   3120
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":BD3C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":C3C27
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":CA489
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":D0CEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":D754D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":DDDAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":E4611
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListviews 
      Left            =   5640
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":EAE73
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":F16D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":F3E87
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":F9AA9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcoForms 
      Left            =   1320
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":10030B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":100D1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":100DB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListComun16 
      Left            =   1200
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   5640
      Top             =   1080
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   4800
      Top             =   1920
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPpalN.frx":1017CA
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   3840
      Top             =   600
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane DockingPaneManager 
      Left            =   4320
      Top             =   1320
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeCommandBars.ImageManager ImageManagerGalleryStyles 
      Left            =   3360
      Top             =   120
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPpalN.frx":1017E4
   End
End
Attribute VB_Name = "frmppal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Dim ContextEvent As CalendarEvent


Dim MRUShortcutBarWidth


Const IMAGEBASE = 10000
Const MinimizedShortcutBarWidth = 32 + 8

Dim WithEvents statusBar  As XtremeCommandBars.statusBar
Attribute statusBar.VB_VarHelpID = -1
Dim FontSizes(4) As Integer
Dim RibbonSeHaCreado As Boolean
Dim Pane As Pane
Dim Cad As String

'Variables comunes para todos los procedimientos de carga menus en el ribbon
'Codejock
Dim TabNuevo As RibbonTab
Dim GroupNew As RibbonGroup, GroupGoTo As RibbonGroup, GroupArrange As RibbonGroup
Dim GroupManageCalendars As RibbonGroup, GroupShare As RibbonGroup, GroupFind As RibbonGroup

Dim Control As CommandBarControl
Dim ControlNew_NewItems As CommandBarPopup
Dim Rn2 As ADODB.Recordset
Dim Habilitado As Boolean


Dim PrimeraVez As Boolean

Dim EmpresasQueYaHaComunicadoAsientosDescuadrados As String






Public Function RibbonBar() As RibbonBar
    Set RibbonBar = CommandBars.ActiveMenuBar
    
End Function

Sub LoadResources(DllName As String, IniFileName As String)
Dim elpath As String
    
    elpath = App.Path & "\Styles\"
    CommandBarsGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    ShortcutBarGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    SuiteControlsGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    CalendarGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    ReportControlGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    DockingPaneGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
End Sub

Public Sub CheckButton(nButton As Integer)
    CommandBars.Actions(ID_OPTIONS_STYLEBLUE2010).Checked = False
    CommandBars.Actions(ID_OPTIONS_STYLESILVER2010).Checked = False
    CommandBars.Actions(ID_OPTIONS_STYLEBLACK2010).Checked = False
    
    CommandBars.Actions(nButton).Checked = True
End Sub

Sub OnThemeChanged(Id As Integer)
Dim N_Skin As Integer
    CheckButton Id
    
    Dim FlatStyle As Boolean
    FlatStyle = Id >= ID_OPTIONS_STYLESCENIC7 And Id <= ID_OPTIONS_STYLEBLACK2010
        
        
    Me.BackColor = frmShortBar.wndShortcutBar.PaintManager.SplitterBackgroundColor
   
    
    CommandBars.EnableOffice2007Frame False

    Select Case CommandBars.VisualTheme
        Case xtpThemeResource, xtpThemeRibbon
            CommandBars.AllowFrameTransparency False 'True
            CommandBars.EnableOffice2007Frame True
            CommandBars.SetAllCaps False
            CommandBars.statusBar.SetAllCaps False
        Case Else
            CommandBars.AllowFrameTransparency True
            CommandBars.EnableOffice2007Frame False
            CommandBars.SetAllCaps False
            CommandBars.statusBar.SetAllCaps False
    End Select
    
    Dim ToolTipContext As ToolTipContext
    Set ToolTipContext = CommandBars.ToolTipContext
    ToolTipContext.Style = xtpToolTipResource
    ToolTipContext.ShowTitleAndDescription True, xtpToolTipIconNone
    ToolTipContext.ShowImage True, IMAGEBASE
    ToolTipContext.SetMargin 2, 2, 2, 2
    ToolTipContext.MaxTipWidth = 180
    
    statusBar.ToolTipContext.Style = ToolTipContext.Style
    frmShortBar.wndShortcutBar.ToolTipContext.Style = ToolTipContext.Style
    
       
    'CreateBackstage
    'SetBackstageTheme
    
    'CommandBars.PaintManager.LoadFrameIcon App.hInstance, App.Path + "\styles\Ariconta.ico", 16, 16
            
    'Set Captions VisualTheme
    On Error Resume Next
    Dim CtrlCaption As ShortcutCaption
    Dim Form As Form, Ctrl As Object
            
    For Each Form In Forms
        For Each Ctrl In Form.Controls
                    
            Set CtrlCaption = Ctrl
            If Not CtrlCaption Is Nothing Then
                CtrlCaption.VisualTheme = frmShortBar.wndShortcutBar.VisualTheme
            End If
                    
        Next
    Next
       
    DockingPaneManager.PaintManager.SplitterSize = 5
    DockingPaneManager.PaintManager.SplitterColor = frmShortBar.wndShortcutBar.PaintManager.SplitterBackgroundColor
    
    DockingPaneManager.PaintManager.ShowCaption = False
    DockingPaneManager.RedrawPanes
        
    frmShortBar.SetColor Id
    frmInbox.SetColor Id
        

    frmPaneCalendar.SetFlatStyle FlatStyle
    frmPaneContacts.SetFlatStyle FlatStyle
    'frmPaneInformacion.SetFlatStyle FlatStyle
    'frmPaneAcercaDe.SetFlatStyle FlatStyle
    
    
    
    
    
    
    LoadIcons
    N_Skin = Id - 2895
    EstablecerSkin N_Skin
    
    'Updatear SKIN usuario
    If CStr(N_Skin) <> vUsu.Skin Then
        vUsu.Skin = N_Skin
        vUsu.ActualizarSkin
    End If
    
End Sub

Public Sub SetBackstageTheme()
Dim i As Integer
    Dim nTheme As XtremeCommandBars.XTPBackstageButtonControlAppearanceStyle
    nTheme = xtpAppearanceResource

   ' If Not (pageBackstageInfo Is Nothing) Then
        'pageBackstageInfo.btnProtectDocument.Appearance = nTheme
        'pageBackstageInfo.btnProtectDocument.Appearance = nTheme
        'pageBackstageInfo.btnCheckForIssues.Appearance = nTheme
        'pageBackstageInfo.btnManageVersions.Appearance = nTheme
   ' End If
    
    If Not (pageBackstageHelp Is Nothing) Then
        For i = 0 To 4
            pageBackstageHelp.btnAcciones(i).Appearance = nTheme
        Next
        
    End If
    
    'If Not (pageBackstageSend Is Nothing) Then
        'pageBackstageSend.btnTab(0).Appearance = nTheme
        'pageBackstageSend.btnTab(1).Appearance = nTheme
        'pageBackstageSend.btnTab(2).Appearance = nTheme
        'pageBackstageSend.btnTab(3).Appearance = nTheme
    'End If

End Sub

Private Sub CreateStatusBar()
Dim Pane As StatusBarPane

    If RibbonSeHaCreado Then
        'StatusBar.Pane(0).Value = vEmpresa.nomempre & "    " & vUsu.Login
        statusBar.Pane(0).Text = "Nº " & vEmpresa.codempre
        statusBar.Pane(1).Text = vEmpresa.nomempre
    
    Else
    
         
         Set statusBar = Nothing
         
         Set statusBar = CommandBars.statusBar
         statusBar.visible = True
         
         
         Set Pane = statusBar.AddPane(ID_INDICATOR_PAGENUMBER)
         Pane.Text = "Nº " & vEmpresa.codempre
         Pane.Caption = "&C"
         Pane.Value = vEmpresa.nomempre & "    " & vUsu.Login
         Pane.Button = True
         Pane.SetPadding 8, 0, 8, 0
         
         Set Pane = statusBar.AddPane(ID_INDICATOR_WORDCOUNT)
         Pane.Text = vEmpresa.nomempre
         Pane.Caption = ""
         Pane.Value = vEmpresa.codempre
         Pane.Button = True
         Pane.SetPadding 8, 0, 8, 0
         
         
         Set Pane = statusBar.AddPane(0)
         Pane.Style = SBPS_STRETCH Or SBPS_NOBORDERS
         Pane.BeginGroup = True
                 
        '
         statusBar.RibbonDividerIndex = 3
         statusBar.EnableCustomization True
         
         CommandBars.Options.KeyboardCuesShow = xtpKeyboardCuesShowNever
         CommandBars.Options.ShowKeyboardTips = True
         CommandBars.Options.ToolBarAccelTips = True
    End If
End Sub

Private Sub DockBarRightOf(BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    CommandBars.RecalcLayout
    BarOnLeft.GetWindowRect Left, top, Right, Bottom
    
    CommandBars.DockToolBar BarToDock, Right, (Bottom + top) / 2, BarOnLeft.Position

End Sub

Private Sub CommandBars_CommandBarKeyDown(CommandBar As XtremeCommandBars.ICommandBar, KeyCode As Long, Shift As Integer)
    Debug.Print CommandBar.BarID
End Sub

Public Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim AbiertoFormulario  As Boolean
    AbiertoFormulario = False
    

    Select Case Control.Id
        Case XTPCommandBarsSpecialCommands.XTP_ID_RIBBONCONTROLTAB:
            
        
          
        Case XTP_ID_RIBBONCUSTOMIZE:
            CommandBars.ShowCustomizeDialog 3
            
        Case ID_APP_ABOUT:
          
           LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & "AriCONTA-6.html?"
   
        
        Case ID_FILE_NEW:
            'frmEmail.Show 0, Me
        
        
        
        Case ID_Licencia_Usuario_Final_txt, ID_Licencia_Usuario_Final_web, ID_Ver_Version_operativa_web
            OpcionesMenuInformacion Control.Id
        
        
        
        Case ID_VIEW_STATUSBAR:
            CommandBars.statusBar.visible = Not CommandBars.statusBar.visible
            CommandBars.RecalcLayout
            
        Case ID_RIBBON_EXPAND:
            RibbonBar.Minimized = Not RibbonBar.Minimized
            
        Case ID_RIBBON_MINIMIZE:
            RibbonBar.Minimized = Not RibbonBar.Minimized
            
        Case ID_OPTIONS_FONT_SYSTEM, ID_OPTIONS_FONT_NORMAL, ID_OPTIONS_FONT_LARGE, ID_OPTIONS_FONT_EXTRALARGE
            Dim newFontHeight As Integer
            newFontHeight = FontSizes(Control.Id - ID_OPTIONS_FONT_SYSTEM)
            RibbonBar.FontHeight = newFontHeight
            
        Case ID_OPTIONS_FONT_AUTORESIZEICONS
            CommandBars.PaintManager.AutoResizeIcons = Not CommandBars.PaintManager.AutoResizeIcons
            CommandBars.RecalcLayout
            RibbonBar.RedrawBar
            
        Case ID_OPTIONS_STYLEBLUE2010:
            LoadResources "Office2010.dll", "Office2010Blue.ini"
            CommandBars.VisualTheme = xtpThemeRibbon
            DockingPaneManager.VisualTheme = ThemeResource
            frmShortBar.wndShortcutBar.VisualTheme = xtpShortcutThemeResource
            frmInbox.CalendarControl.VisualTheme = xtpCalendarThemeResource
            frmInbox.ScrollBarCalendar.Appearance = xtpAppearanceResource
            
            OnThemeChanged ID_OPTIONS_STYLEBLUE2010
            
            
            
       Case ID_OPTIONS_STYLESILVER2010:
            LoadResources "Office2010.dll", "Office2010Silver.ini"
            CommandBars.VisualTheme = xtpThemeRibbon
            DockingPaneManager.VisualTheme = ThemeResource
            frmShortBar.wndShortcutBar.VisualTheme = xtpShortcutThemeResource
            frmInbox.CalendarControl.VisualTheme = xtpCalendarThemeResource
            frmInbox.ScrollBarCalendar.Appearance = xtpAppearanceResource
            
            OnThemeChanged ID_OPTIONS_STYLESILVER2010
        
       Case ID_OPTIONS_STYLEBLACK2010:
            LoadResources "Office2010.dll", "Office2010Black.ini"
            CommandBars.VisualTheme = xtpThemeRibbon
            DockingPaneManager.VisualTheme = ThemeResource
            frmShortBar.wndShortcutBar.VisualTheme = xtpShortcutThemeResource
            frmInbox.CalendarControl.VisualTheme = xtpCalendarThemeResource
            frmInbox.ScrollBarCalendar.Appearance = xtpAppearanceResource
            
            OnThemeChanged ID_OPTIONS_STYLEBLACK2010
        
        Case ID_APP_EXIT:
            Unload Me
        
    
            
        Case ID_GROUP_GOTO_TODAY:
            Select Case frmInbox.CalendarControl.ViewType
                Case xtpCalendarDayView:
                    frmInbox.CalendarControl.DayView.ShowDay DateTime.Now, True
            
                Case xtpCalendarWorkWeekView:
                    frmInbox.CalendarControl.DayView.SetSelection DateTime.Now, DateTime.Now, True
                    frmInbox.CalendarControl.RedrawControl
            
                Case xtpCalendarWeekView:
                    frmInbox.CalendarControl.WeekView.SetSelection DateTime.Now, DateTime.Now, True
            
                Case xtpCalendarMonthView:
                    frmInbox.CalendarControl.MonthView.SetSelection DateTime.Now, DateTime.Now, True
            End Select
            
        Case ID_GROUP_GOTO_NEXT7DAYS:
            Dim lastDate As Date
            lastDate = frmInbox.CalendarControl.DayView.Days(frmInbox.CalendarControl.DayView.DaysCount - 1).Date
            frmInbox.CalendarControl.ViewType = xtpCalendarDayView
            frmInbox.CalendarControl.DayView.ShowDays lastDate + 1, lastDate + 7
            
        Case ID_GROUP_ARRANGE_DAY:
            frmInbox.CalendarControl.ViewType = xtpCalendarDayView
            
        Case ID_GROUP_ARRANGE_WORK_WEEK:
            frmInbox.CalendarControl.ViewType = xtpCalendarWorkWeekView
            
        Case ID_GROUP_ARRANGE_WEEK:
            frmInbox.CalendarControl.UseMultiColumnWeekMode = True
            frmInbox.CalendarControl.ViewType = xtpCalendarWeekView

        Case ID_GROUP_ARRANGE_MONTH, ID_GROUP_ARRANGE_MONTH_LOW, _
             ID_GROUP_ARRANGE_MONTH_MEDIUM, ID_GROUP_ARRANGE_MONTH_HIGH:
            frmInbox.CalendarControl.ViewType = xtpCalendarMonthView
            
        Case ID_CALENDAREVENT_OPEN:
            frmInbox.mnuOpenEvent
            
        Case ID_CALENDAREVENT_DELETE:
            frmInbox.mnuDeleteEvent
            
        Case ID_CALENDAREVENT_NEW, ID_GROUP_NEW_APPOINTMENT:
            'falta### frmEditEvent.AllDayOverride = False
            frmInbox.mnuNewEvent
            frmInbox.CalendarControl.Options.DayViewCurrentTimeMarkVisible = True
            
        Case ID_GROUP_NEW_MEETING:
            'falta### frmEditEvent.AllDayOverride = False
            'falta### frmEditEvent.chkMeeting.Value = 1
            frmInbox.mnuNewEvent
            frmInbox.CalendarControl.Options.DayViewCurrentTimeMarkVisible = True
            
        Case ID_GROUP_NEW_ALLDAY:
            'falta### frmEditEvent.AllDayOverride = True
            frmInbox.mnuNewEvent
            frmInbox.CalendarControl.Options.DayViewCurrentTimeMarkVisible = True
        Case ID_GROUP_SHARE_SHARE
            frmInbox.CalendarControl.PrintOptions.Footer.TextCenter = vEmpresa.nomempre
            frmInbox.CalendarControl.PrintOptions.Footer.TextLeft = "Ariconta6. Ariadna SW"
            frmInbox.CalendarControl.PrintOptions.Footer.TextRight = Format(Now, "dd/mm/yyyy hh:mm")
            frmInbox.CalendarControl.PrintPreviewOptions.Title = "Ariconta6 " & vEmpresa.nomempre
            frmInbox.CalendarControl.PrintPreview True
            
        Case ID_CALENDAREVENT_CHANGE_TIMEZONE:
            frmInbox.mnuChangeTimeZone
            
        Case ID_CALENDAREVENT_60:
            frmInbox.mnuTimeScale 60
            
        Case ID_CALENDAREVENT_30:
            frmInbox.mnuTimeScale 30
            
        Case ID_CALENDAREVENT_15:
            frmInbox.mnuTimeScale 15
            
        Case ID_CALENDAREVENT_10:
            frmInbox.mnuTimeScale 10
            
        Case ID_CALENDAREVENT_5:
            frmInbox.mnuTimeScale 5
            
            
            
     
        Case Else
            AbiertoFormulario = True
            AbrirFormularios Control.Id
            
            
    End Select
    
    
    If AbiertoFormulario Then
        AbiertoFormulario = False
        'mOTIVO... no lo se
        'Pero si lo vamos cambiando funciona
        If Me.DockingPaneManager.Panes(1).Enabled = 3 Then
            Me.DockingPaneManager.Panes(1).Enabled = 3
            Me.DockingPaneManager.Panes(2).Enabled = 3

            frmPaneCalendar.DatePicker.Enabled = True
            
            DockingPaneManager.RedrawPanes
            
            
        Else
            Me.DockingPaneManager.Panes(1).Enabled = 3
            Me.DockingPaneManager.Panes(2).Enabled = 3
             
        End If
        DockingPaneManager.NormalizeSplitters

    End If
End Sub



Private Sub CommandBars_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
        Dim Control As CommandBarControl, ControlItem As CommandBarControl
        
        If TypeOf CommandBar Is RibbonBackstageView Then
            Debug.Print "RibbonBackstageView"
        End If
        
        Set Control = CommandBar.FindControl(, IDS_ARRANGE_BY)
        If Not Control Is Nothing Then
            Dim Index As Long
            Index = Control.Index
            Control.visible = False
            
            Do While Index + 1 <= CommandBar.Controls.Count
                Set ControlItem = CommandBar.Controls.Item(Index + 1)
                If ControlItem.Id = IDS_ARRANGE_BY Then
                    ControlItem.Delete
                Else
                    Exit Do
                End If
            Loop
            
'            Dim CurrentColumn As ReportColumn
'            For Each CurrentColumn In frmInbox. wndReportControl.Columns
'                Set ControlItem = CommandBar.Controls.Add(xtpControlButton, ID_REPORTCONTROL_COLUMN_ARRANGE_BY, CurrentColumn.Caption)
'                ControlItem.Parameter = CurrentColumn.ItemIndex
'                If Not frmInbox. wndReportControl.SortOrder.IndexOf(CurrentColumn) = -1 Then
'                    ControlItem.Checked = True
'                End If
'                If Not CurrentColumn.Visible Then
'                    ControlItem.Visible = False
'                End If
'            Next
        
        End If
End Sub

Private Sub CommandBars_SpecialColorChanged()
    Me.BackColor = CommandBars.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
End Sub

Private Sub CommandBars_ToolBarVisibleChanged(ByVal ToolBar As XtremeCommandBars.ICommandBar)
     Debug.Print ToolBar.BarID
End Sub

Private Sub CommandBars_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
        
    On Error Resume Next
    
    
    
    Select Case Control.Id
        Case ID_VIEW_STATUSBAR:
            'Control.Checked = CommandBars.StatusBar.Visible
        
        
            
        Case ID_GROUP_ARRANGE_WORK_WEEK:
            'Control.Checked = IIf(frmInbox.CalendarControl.ViewType = xtpCalendarWorkWeekView, True, False)
            
        Case ID_GROUP_ARRANGE_WEEK:
            'Control.Checked = IIf(frmInbox.CalendarControl.ViewType = xtpCalendarWeekView, True, False)
            
        Case ID_GROUP_ARRANGE_MONTH:
            'Control.Checked = IIf(frmInbox.CalendarControl.ViewType = xtpCalendarMonthView, True, False)
        
        Case ID_OPTIONS_ANIMATION:
            'Control.Checked = CommandBars.ActiveMenuBar.EnableAnimation
            
        Case ID_OPTIONS_FONT_SYSTEM, ID_OPTIONS_FONT_NORMAL, ID_OPTIONS_FONT_LARGE, ID_OPTIONS_FONT_EXTRALARGE
             '   Dim newFontHeight As Integer
             '   newFontHeight = FontSizes(Control.Id - ID_OPTIONS_FONT_SYSTEM)
             '   Control.Checked = IIf(RibbonBar.FontHeight = newFontHeight, True, False)
                
        Case ID_OPTIONS_FONT_AUTORESIZEICONS
              '  Control.Checked = CommandBars.PaintManager.AutoResizeIcons

        Case ID_RIBBON_EXPAND:
            'Control.Visible = RibbonBar.Minimized
            
        Case ID_RIBBON_MINIMIZE:
            'Control.Visible = Not RibbonBar.Minimized
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub DockingPaneManager_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, ByVal Container As XtremeDockingPane.IPaneActionContainer, Cancel As Boolean)
    If (Action = PaneActionSplitterResized) Then
        DockingPaneManager.RecalcLayout
        
        ' Save MRUShortcutBarWidth
        If (frmShortBar.ScaleWidth > MinimizedShortcutBarWidth And Container.Container.Type = PaneTypeSplitterContainer) Then
            Debug.Print frmShortBar.ScaleWidth
            MRUShortcutBarWidth = frmShortBar.ScaleWidth
        End If
    Else
        If (Action = PaneActionSplitterResized) Then Debug.Print "Resizing "
    End If
End Sub

Private Sub DockingPaneManager_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.Tag = PANE_SHORTCUTBAR Then
        Item.Handle = frmShortBar.hWnd
    ElseIf Item.Tag = PANE_REPORT_CONTROL Then
        Item.Handle = frmInbox.hWnd
    End If
End Sub

Private Sub Form_Activate()
Dim res

    If PrimeraVez Then
        PrimeraVez = False
        DoEvents
        

        
        AccionesIncioAbrirProgramaEmpresa
        
        


    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaDatosMenusDemas(DesdeLoad As Boolean)
Dim AntiguoTab As Integer
    
    
    Screen.MousePointer = vbHourglass
    AntiguoTab = -1
    If RibbonSeHaCreado Then
        If Not RibbonBar.SelectedTab Is Nothing Then AntiguoTab = RibbonBar.SelectedTab.Id
    End If
    CreateRibbon
    Screen.MousePointer = vbHourglass
    CreateBackstage
    Screen.MousePointer = vbHourglass
    CreateRibbonOptions
    
    If vEmpresa.TieneTesoreria And vUsu.SoloTesoreria = 1 Then
        vEmpresa.SoloTesoreria
    Else
        If Not DesdeLoad Then vEmpresa.Leer vEmpresa.codempre
    End If
    'vEmpresa.TieneContabilidad = False
    '??????
    '0=solo contabilidad / 1=todo / 2=solo tesoreria
    Screen.MousePointer = vbHourglass
    CargaMenu AntiguoTab
    CreateStatusBar
    Screen.MousePointer = vbHourglass
    PonerCaption
    CreateCalendarTabOriginal
    RibbonSeHaCreado = True
End Sub






Public Sub CambiarEmpresa(QueEmpresa As Integer)
Dim cur As Integer

    Screen.MousePointer = vbHourglass
    Me.Hide
    CambiarEmpresa2 QueEmpresa
    Me.Show
       DoEvents
       Screen.MousePointer = vbHourglass
    AccionesIncioAbrirProgramaEmpresa
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub CambiarEmpresa2(QueEmpresa As Integer)
Dim RB As RibbonBar
    CadenaDesdeOtroForm = vUsu.Login & "|" & vEmpresa.codempre & "|"
        
    
        
    Set vUsu = New Usuario
    vUsu.Leer RecuperaValor(CadenaDesdeOtroForm, 1)
    
    vUsu.CadenaConexion = "ariconta" & QueEmpresa
    
    vUsu.LeerFiltros "ariconta", 301 ' asientos
    vUsu.LeerFiltros "ariconta", 401 ' facturas de cliente
    
    AbrirConexion vUsu.CadenaConexion
    
    Set vEmpresa = New Cempresa
    Set vParam = New Cparametros
    Set vParamT = New CparametrosT
    'NO DEBERIAN DAR ERROR
    vEmpresa.Leer
    vParam.Leer
    If vEmpresa.TieneTesoreria Then vParamT.Leer
    
    PonerCaption
    
    Screen.MousePointer = vbHourglass
   CargaDatosMenusDemas True
   frmPaneContacts.SeleccionarNodoEmpresa vEmpresa.codempre
   pageBackstageHelp.Label9.Caption = vEmpresa.nomempre
   pageBackstageHelp.tabPage(0).visible = False
   pageBackstageHelp.tabPage(1).visible = False
   frmInbox.OpenProvider
   Set RB = RibbonBar
   RB.Minimized = False
   RB.RedrawBar
   
   
  
   
    vControl.UltEmpre = vUsu.CadenaConexion
    
    vControl.Grabar
    
    
    
    
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Load()
   
    'Cargamos librerias de icinos de los forms
    frmIdentifica.pLabel "Carga DLL"
    CargaIconosDlls
   
    
   
    CommandBarsGlobalSettings.App = App
            
    frmIdentifica.pLabel "Leyendo menus usuario"
    CargaDatosMenusDemas True
    
    ShowEventInPane = False
       
    FontSizes(0) = 0
    FontSizes(1) = 11
    FontSizes(2) = 13
    FontSizes(3) = 16
               
    DockingPaneManager.SetCommandBars Me.CommandBars
              
    Set frmShortBar = New frmShortcutBar2
    Set frmInbox = New frmInbox
        
    Dim A As Pane, B As Pane, C As Pane, d As Pane
    
    frmIdentifica.pLabel "Creando paneles"
    Set A = DockingPaneManager.CreatePane(PANE_SHORTCUTBAR, 170, 120, DockLeftOf, Nothing)
    A.Tag = PANE_SHORTCUTBAR
    A.MinTrackSize.Width = MinimizedShortcutBarWidth
    
    Set B = DockingPaneManager.CreatePane(PANE_REPORT_CONTROL, 700, 400, DockRightOf, A)
    B.Tag = PANE_REPORT_CONTROL
   
    DockingPaneManager.Options.HideClient = True
    PonerTabPorDefecto -1
    
    Set CommandBars.Icons = CommandBarsGlobalSettings.Icons
    LoadIcons
    
    DockingPaneManager.RecalcLayout
    MRUShortcutBarWidth = frmShortBar.ScaleWidth
   
   
    'En funcion
    ' ID_OPTIONS_STYLEBLUE2010  ID_OPTIONS_STYLESILVER2010    ID_OPTIONS_STYLEBLACK2010
    frmIdentifica.pLabel "Carga skin"
    Screen.MousePointer = vbHourglass
    If vUsu.Skin = 3 Then
        Cad = ID_OPTIONS_STYLEBLACK2010
    Else
        If vUsu.Skin = 2 Then
            Cad = ID_OPTIONS_STYLESILVER2010
        Else
            Cad = ID_OPTIONS_STYLEBLUE2010
        End If
    End If
    CommandBars.FindControl(, Cad, , True).Execute
    
    PrimeraVez = True

    
End Sub


Private Sub CargaIconosDlls()
Dim TamanyoImgComun As Integer

    ImageList1.ImageHeight = 48
    ImageList1.ImageWidth = 48
    GetIconsFromLibrary App.Path & "\styles\icoconppal.dll", 1, 48


    ImageList2.ImageHeight = 16
    ImageList2.ImageWidth = 16
    GetIconsFromLibrary App.Path & "\styles\icoconppal.dll", 1, 16

    ImageListPPal48.ImageHeight = 48
    ImageListPPal48.ImageWidth = 48
    GetIconsFromLibrary App.Path & "\styles\icoconppal2.dll", 8, 48


    ImageListPpal16.ImageHeight = 16
    ImageListPpal16.ImageWidth = 16
    GetIconsFromLibrary App.Path & "\styles\icoconppal2.dll", 9, 16

    
    
    imgListComun.ListImages.Clear
    imgListComun_BN.ListImages.Clear
    imgListComun_OM.ListImages.Clear
    
        TamanyoImgComun = 24
        
        imgListComun.ImageHeight = TamanyoImgComun
        imgListComun.ImageWidth = TamanyoImgComun
        GetIconsFromLibrary App.Path & "\styles\iconosconta.dll", 2, TamanyoImgComun  'antes icolistcon
    
        
        
        '++
        imgListComun_BN.ImageHeight = TamanyoImgComun
        imgListComun_BN.ImageWidth = TamanyoImgComun
        GetIconsFromLibrary App.Path & "\styles\iconosconta_BN.dll", 3, TamanyoImgComun
      
        imgListComun_OM.ImageHeight = TamanyoImgComun
        imgListComun_OM.ImageWidth = TamanyoImgComun
        GetIconsFromLibrary App.Path & "\styles\iconosconta_OM.dll", 4, TamanyoImgComun
        
    
    imgListComun16.ImageHeight = 16
    imgListComun16.ImageWidth = 16
    GetIconsFromLibrary App.Path & "\styles\iconosconta.dll", 5, 16
    
    GetIconsFromLibrary App.Path & "\styles\iconosconta_BN.dll", 6, 16
    GetIconsFromLibrary App.Path & "\styles\iconosconta_OM.dll", 7, 16


End Sub

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



Public Sub ExpandButtonClicked()
   
    
    
    Dim A As Pane
    Set A = DockingPaneManager.FindPane(PANE_SHORTCUTBAR)
    
    Dim ShortcutBarMinimized As Boolean
    ShortcutBarMinimized = frmShortBar.ScaleWidth <= MinimizedShortcutBarWidth
    
    Dim NewWidth As Long
    If (ShortcutBarMinimized) Then
        NewWidth = MRUShortcutBarWidth
    Else
        NewWidth = MinimizedShortcutBarWidth
        frmShortBar.wndShortcutBar.PopupWidth = MRUShortcutBarWidth
    End If
        
    
    ' Set Size of Pane
    A.MinTrackSize.Width = NewWidth
    A.MaxTrackSize.Width = NewWidth
        
    DockingPaneManager.RecalcLayout
    DockingPaneManager.NormalizeSplitters
    DockingPaneManager.RedrawPanes
    
    ' Restore Constraints
    A.MinTrackSize.Width = MinimizedShortcutBarWidth
    A.MaxTrackSize.Width = 32000
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If Not (pageBackstageInfo Is Nothing) Then Unload pageBackstageInfo
    If Not (pageBackstageHelp Is Nothing) Then Unload pageBackstageHelp
    'If Not (pageBackstageSend Is Nothing) Then Unload pageBackstageSend
    
    'close all sub forms
    On Error Resume Next
    Dim i As Long
    For i = Forms.Count - 1 To 1 Step -1
        
        Unload Forms(i)
    Next
    
    
    GuardarDatosUltimaTab
  
End Sub



Private Sub GuardarDatosUltimaTab()
    i = RibbonBar.SelectedTab.Id
    If i = ID_TAB_CALENDAR_HOME Then Exit Sub 'no guardo este tab
    If i <> vUsu.TabPorDefecto Then
        vUsu.TabPorDefecto = i
        vUsu.GuardarTabPorDefecto
    End If
End Sub


Public Function AddButton(Controls As CommandBarControls, ControlType As XTPControlType, Id As Long, Caption As String, Optional BeginGroup As Boolean = False, Optional DescriptionText As String = "", Optional ButtonStyle As XTPButtonStyle = xtpButtonAutomatic, Optional Category As String = "Controls") As CommandBarControl
    Dim Control As CommandBarControl
    Set Control = Controls.Add(ControlType, Id, Caption)
    
    Control.BeginGroup = BeginGroup
    Control.DescriptionText = DescriptionText
    Control.Style = ButtonStyle
    Control.Category = Category
    
    Set AddButton = Control
    
End Function

Private Sub CommandBars_Resize()
    
    On Error Resume Next
    
    Dim Left As Long
    Dim top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    CommandBars.GetClientRect Left, top, Right, Bottom
    
End Sub

Private Sub LoadIcons()
    CommandBars.Icons.RemoveAll
    SuiteControlsGlobalSettings.Icons.RemoveAll
    ReportControlGlobalSettings.Icons.RemoveAll

    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\help.png", ID_APP_ABOUT, xtpImageNormal
        
        
        
   
    'Para que no carge imagen de ratios y graficas y punteo, no lo pongo aqui ya que los cargo "pequeños"
    '
  
      
    'ICONOS PEQUEÑOS
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
            Array(ID_RatiosyGráficas, ID_EvolucióndeSaldos, ID_Totalesporconcepto, 1, 1, ID_AseguClientes), xtpImageNormal
        
    
    
    
    'Pequeños
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", _
            Array(1, 1, 1, 1, ID_EstadísticaInmovilizado, ID_SimulaciónAmortización, ID_DeshacerAmortización, 1, 1, ID_VentaBajainmovilizado), xtpImageNormal
        
    'Pequeños diario
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
            Array(ID_TiposdeDiario, ID_TiposdePago, ID_ModelosdeCartas, ID_BicSwift, 1, ID_Agentes), xtpImageNormal
      
         '      ID_AsientosPredefinidos
     
    'Deberiamos cargar un array con unos(1) de longitud 143
    ' y en funcion del valor del campo imagen en el punto de menu correspondiente
    ' lo pondremos en el array.
    ' Ejemplo    303 Extractos  Campo imagen: 87
    ' quiere decir que en el campo 87 del array sustituieremos el 1 por el 303


'
    Dim T() As Variant
    'Cad linea son 15
    T = Array(1, ID_Conceptos, ID_TiposdeIVA, ID_Bancos, ID_FormasdePago, ID_FacturasRecibidas, 1, ID_FacturasEmitidas, ID_LibroFacturasRecibidas, 1, 1, 1, 1, 1, 1, _
        ID_RelaciónClientesporcuenta, ID_RelacionProveedoresporcuenta, 1, 1, 1, ID_RealizarCobro, ID_RealizarPago, 1, ID_Elementos, 1, 1, 1, 1, 1, ID_Punteoextractobancario, _
        1, ID_InformePagospendientes, 1, 1, ID_Empresa, ID_ParametrosContabilidad, 1, ID_Contadores, ID_Extractos, ID_CarteradePagos, 1, 1, 1, ID_Punteo, 1, _
        1, ID_PlanContable, 1, ID_ConsoPyG, 1, ID_Informes, 1, ID_Usuarios, 1, 1, 1, 1, ID_Nuevaempresa, ID_ConfigurarBalances, 1, _
        ID_ConsoSitu, ID_Compensaciones, 1, 1, 1, 1, ID_ConceptosInm, 1, 1, ID_GenerarAmortización, 1, 1, 1, 1, 1, _
        ID_ImportarFacturasCliente, 1, 1, ID_Compensarcliente, ID_SumasySaldos, ID_CuentadeExplotación, ID_BalancedeSituación, ID_PérdidasyGanancias, 1, 1, 1, 1, 1, ID_CarteradeCobros, ID_InformeCobrosPendientes, _
        ID_Renumeracióndeasientos, ID_CierredeEjercicio, ID_Deshacercierre, ID_DiarioOficial, ID_PresentaciónTelemáticadeLibros, ID_Traspasodecuentasenapuntes, ID_Renumerarregistrosproveedor, ID_TraspasocodigosdeIVA, 1, 1, 1, 1, 1, 1, 1, _
        ID_Traspasodecuentasenapuntes, ID_Aumentardígitoscontables, 1, 1, 1, 1, 1, ID_LibroFacturasEmitidas, 1, 1, ID_Remesas, 1, ID_Consolidado, 1, ID_GraficosChart, _
        ID_RecepcionTalónPagare, ID_RemesasTalenPagare, ID_Accionesrealizadas, 1, ID_LiquidacionIVA, 1, 1, 1, ID_AsientosPredefinidos, 1, 1, 1, 1, ID_FrasConso, 1, _
        ID_Renumerarregistrosproveedor, ID_GROUP_SHARE_SHARE, 1, ID_ConsoSumasSaldos, ID_ConsoCtaExplota, ID_Asientos, ID_TraspasocodigosdeIVA, 1)
    
     
    
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlook2013L_32x32.bmp", T, xtpImageNormal
    
           

    'Este de abjo funciona correctamente.
    'NO tocar. Es por si falla volver a empezar
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlook2013L_32x32.bmp", _
'            Array(ID_CarteradeCobros, ID_InformeCobrosPendientes, ID_RealizarCobro, ID_Compensarcliente, 1, ID_BalancePresupuestario, 1, _
'            ID_CentrosdeCoste, 1, 1, ID_Presupuestos, ID_Remesas, ID_Detalledeexplotación, ID_CarteradePagos, ID_CuentadeExplotaciónAnalítica, ID_ExtractosporCentrodeCoste, _
'            ID_Asientos, ID_Extractos, ID_Punteo, 1, ID_CuentadeExplotación, ID_Totalesporconcepto, ID_BalancedeSituación, ID_PérdidasyGanancias, _
'            ID_SumasySaldos, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
'            1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
'            ID_Empresa, ID_ParametrosContabilidad, ID_Contadores, ID_Usuarios, 1, ID_Informes, ID_Nuevaempresa, ID_ConfigurarBalances, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
'            ID_FacturasEmitidas, ID_LibroFacturasEmitidas, ID_FacturasRecibidas, ID_LibroFacturasRecibidas, 1, 1, 1, 1, 1, ID_Elementos, ID_GenerarAmortización, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
'            1, ID_PlanContable, ID_TiposdeDiario, ID_Conceptos, ID_TiposdeIVA, ID_TiposdePago, ID_Bancos, ID_FormasdePago, _
'            ID_BicSwift, ID_Agentes, ID_AsientosPredefinidos, ID_ModelosdeCartas, _
'            ID_Renumeracióndeasientos, ID_CierredeEjercicio, ID_Deshacercierre, 1, 1, 1, 1, 1, 1, ID_DiarioOficial, _
'            ID_PresentaciónTelemáticadeLibros, ID_Traspasodecuentasenapuntes, ID_Renumerarregistrosproveedor, 1, ID_TraspasocodigosdeIVA), xtpImageNormal
'
    
    'Presupuiestaria y analitaica cargadas arriba en pequeño
    '---------------------------------------------------------
    '
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
            Array(ID_CentrosdeCoste, ID_ExtractosporCentrodeCoste, ID_Detalledeexplotación, ID_CuentadeExplotaciónAnalítica, ID_Presupuestos, ID_BalancePresupuestario), xtpImageNormal
    

    

    'Pequeños
    ' ID_Compensaciones ID_Reclamaciones  ID_InformeImpagados ID_RemesasTalenPagare ID_Norma57Pagosventanilla  ID_TransferenciasAbonos
    ' ID_InformePagosbancos ID_Transferencias ID_Pagosdomiciliados ID_GastosFijos ID_Compensarproveedor ID_Confirming
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", _
            Array(ID_AnticipoFacturas, ID_Reclamaciones, ID_InformeImpagados, ID_RemesasTalenPagare, ID_Norma57Pagosventanilla, ID_TransferenciasAbonos, ID_Confirming, _
            ID_Pagosdomiciliados, ID_GastosFijos, ID_Compensarproveedor), xtpImageNormal
    
    
    
    
    
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
            Array(ID_InformePagosbancos, ID_Transferencias, ID_MemoriaPlazosdepago, ID_Informeporcuenta, ID_SituaciónTesoreria, ID_InformeporNIF), xtpImageNormal
    
     
 
        
        
    '------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlookcalicons.png", _
            Array(ID_GROUP_NEW_APPOINTMENT, ID_GROUP_NEW_MEETING, ID_GROUP_NEW_ITEMS, ID_GROUP_GOTO_TODAY, _
            ID_GROUP_GOTO_NEXT7DAYS, ID_GROUP_ARRANGE_DAY, ID_GROUP_ARRANGE_WORK_WEEK, ID_GROUP_ARRANGE_WEEK, _
            ID_GROUP_ARRANGE_MONTH, ID_GROUP_ARRANGE_SCHEDULE_VIEW, ID_GROUP_MANAGE_CALENDARS_OPEN, ID_GROUP_MANAGE_CALENDARS_GROUPS, _
            1, ID_ConfigurarBalances1, ID_ConfigurarBalances2, ID_ConfigurarBalances3), xtpImageNormal
            
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\RibbonMinimize.png", _
            Array(ID_RIBBON_MINIMIZE, ID_RIBBON_EXPAND), xtpImageNormal
            
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\Search.png", _
            ID_SEARCH_ICON, xtpImageNormal
            
     CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\reporticonslarge.png", _
            Array(ID_GROUP_MAIL_NEW_NEW, ID_GROUP_MAIL_NEW_NEW_ITEMS, ID_GROUP_MAIL_DELETE_DELETE, ID_GROUP_MAIL_RESPOND_REPLY, _
            ID_GROUP_MAIL_RESPOND_REPLY_ALL, ID_GROUP_MAIL_RESPOND_FORWARD, ID_GROUP_MAIL_MOVE_MOVE, ID_GROUP_MAIL_MOVE_ONENOTE), xtpImageNormal
            
     CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\reporticonssmall.png", _
            Array(ID_GROUP_MAIL_DELETE_CLEANUP, ID_GROUP_MAIL_DELETE_JUNK, ID_GROUP_MAIL_RESPOND_MEETING, ID_GROUP_MAIL_RESPOND_IM, _
            ID_GROUP_MAIL_RESPOND_MORE, ID_GROUP_MAIL_TAGS_UNREAD, ID_GROUP_MAIL_TAGS_CATEGORIZE, ID_GROUP_MAIL_TAGS_FOLLOWUP, ID_GROUP_MAIL_FIND_ADDRESSBOOK, _
            ID_GROUP_MAIL_FIND_FILTER, ID_GROUP_MAIL_MOVE_MOVE, ID_GROUP_MAIL_MOVE_ONENOTE), xtpImageNormal
    
        CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlookpane.png", _
            Array(ID_SWITCH_NORMAL, ID_SWITCH_CALENAR_AND_TASK, ID_SWITCH_CALENDAR, ID_SWITCH_CLASSIC, ID_SWITCH_READING), xtpImageNormal
            
        CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", _
            Array(SHORTCUT_INBOX, SHORTCUT_CALENDAR, SHORTCUT_CONTACTS, SHORTCUT_TASKS, SHORTCUT_NOTES, _
            SHORTCUT_FOLDER_LIST, SHORTCUT_SHORTCUTS, SHORTCUT_JOURNAL, SHORTCUT_SHOW_MORE, SHORTCUT_SHOW_FEWER), xtpImageNormal
        CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_24x24.bmp", _
            Array(SHORTCUT_INBOX, SHORTCUT_CALENDAR, SHORTCUT_CONTACTS, SHORTCUT_TASKS, SHORTCUT_NOTES, _
            SHORTCUT_FOLDER_LIST, SHORTCUT_SHORTCUTS, SHORTCUT_JOURNAL, SHORTCUT_SHOW_MORE, SHORTCUT_SHOW_FEWER), xtpImageNormal
            
        CommandBars.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
            Array(ID_QUICKSTEP_REPLAY_DELETE, ID_QUICKSTEP_TO_MANAGER, ID_QUICKSTEP_MOVE_TO, ID_QUICKSTEP_CREATE_NEW, ID_QUICKSTEP_TEAM_EMAIL, ID_QUICKSTEP_DONE), xtpImageNormal
            
        ReportControlGlobalSettings.Icons.LoadBitmap App.Path & "\styles\bmreport.bmp", _
        Array(COLUMN_MAIL_ICON, COLUMN_IMPORTANCE_ICON, COLUMN_CHECK_ICON, RECORD_UNREAD_MAIL_ICON, RECORD_READ_MAIL_ICON, _
            RECORD_REPLIED_ICON, RECORD_IMPORTANCE_HIGH_ICON, COLUMN_ATTACHMENT_ICON, COLUMN_ATTACHMENT_NORMAL_ICON, _
            RECORD_IMPORTANCE_LOW_ICON), xtpImageNormal
            
            
        CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\suministro-inmediato-informacion.bmp", ID_SII, xtpImageNormal
            
            
        Dim i As Integer
        For i = 1 To 17
            SuiteControlsGlobalSettings.Icons.LoadIcon App.Path & "\styles\TreeView\icon" & i & ".ico", i, xtpImageNormal
        Next i
End Sub

Private Sub SaveRibbonBarToXML()
    Dim Px As PropExchange
    Set Px = XtremeCommandBars.CreatePropExchange()
    
    Px.CreateAsXML False, "Settings"
        
    Dim Options As StateOptions
    Set Options = CommandBars.CreateStateOptions()
    Options.SerializeControls = True
        
    CommandBars.DoPropExchange Px.GetSection("CommandBars"), Options
    
    Px.SaveToFile "C:\Layout.xml"
    
End Sub



Private Function CreateQuickStepGallery() As CommandBarGalleryItems

    Dim GalleryItems As CommandBarGalleryItems
    Set GalleryItems = CommandBars.CreateGalleryItems(ID_GALLERY_QUICKSTEP)
        
    GalleryItems.ItemWidth = 120
    GalleryItems.ItemHeight = 20
            
    GalleryItems.AddItem ID_QUICKSTEP_MOVE_TO, "Move To: ?"
    GalleryItems.AddItem ID_QUICKSTEP_TO_MANAGER, "To Manager"
    GalleryItems.AddItem ID_QUICKSTEP_TEAM_EMAIL, "Team E-mail"
    GalleryItems.AddItem ID_QUICKSTEP_DONE, "Done"
    GalleryItems.AddItem ID_QUICKSTEP_REPLAY_DELETE, "Reply & Delete"
    GalleryItems.AddItem ID_QUICKSTEP_CREATE_NEW, "Create New"
        
    GalleryItems.Icons = CommandBarsGlobalSettings.Icons

    Set CreateQuickStepGallery = GalleryItems

End Function

Private Sub CommandBars_ControlNotify(ByVal Control As XtremeCommandBars.ICommandBarControl, ByVal Code As Long, ByVal NotifyData As Variant, Handled As Variant)
   
    If (Code = XTP_BS_TABCHANGED) Then

        
    End If
End Sub


Private Sub CreateBackstage()

    
    Dim RibbonBar As RibbonBar
    Set RibbonBar = CommandBars.ActiveMenuBar
    
    Dim BackstageView As RibbonBackstageView
    Set BackstageView = CommandBars.CreateCommandBar("CXTPRibbonBackstageView")
    
    BackstageView.SetTheme xtpThemeRibbon


    CommandBars.Icons.LoadBitmap App.Path & "\styles\BackstageIcons.png", _
    Array(1, 1, 1002, 1, 1, ID_APP_EXIT), xtpImageNormal

    Set RibbonBar.AddSystemButton.CommandBar = BackstageView
    
    'BackstageView.AddCommand ID_FILE_SAVE, "Cambiar empresa"
    'BackstageView.AddCommand ID_FILE_SAVE_AS, "Personalizar"
    'BackstageView.AddCommand ID_FILE_OPEN, "Open"
    'BackstageView.AddCommand ID_FILE_CLOSE, "Close"
    
    'If (pageBackstageInfo Is Nothing) Then Set pageBackstageInfo = New pageBackstageInfo
    'If (pageBackstageSend Is Nothing) Then Set pageBackstageSend = New pageBackstageSend
    If (pageBackstageHelp Is Nothing) Then Set pageBackstageHelp = New pageBackstageHelp
    
    Dim ControlInfo As RibbonBackstageTab
    Set ControlInfo = BackstageView.AddTab(1000, "Info", pageBackstageHelp.hWnd)
    
    'BackstageView.AddTab 1002, "Empresas", pageBackstageSend.hwnd

    ' Los menus de informacion...
    'BackstageView.AddTab 1001, "Acerca de", pageBackstageInfo.hwnd
    
    
    
    
    
    
    
    
    
    
    'BackstageView.AddCommand ID_FILE_OPTIONS, "Options"
    BackstageView.AddCommand ID_APP_EXIT, "Salir"
    
    ControlInfo.DefaultItem = True
    

End Sub




Private Sub CreateCalendarTabOriginal()

    Dim TabCalendarHome As RibbonTab
    Dim GroupNew As RibbonGroup, GroupGoTo As RibbonGroup, GroupArrange As RibbonGroup

    
    Dim Control As CommandBarControl
    Dim ControlNew_NewItems As CommandBarPopup
    Dim ControlArrange_Month As CommandBarPopup
    Dim ControlManage_Open As CommandBarPopup
    Dim ControlManage_Groups As CommandBarPopup
    Dim ControlShare_Publish As CommandBarPopup
           
    Dim PopupBar As CommandBar
    
    Set TabCalendarHome = RibbonBar.InsertTab(14, "Agenda")
    TabCalendarHome.Id = ID_TAB_CALENDAR_HOME
 
    Set GroupNew = TabCalendarHome.Groups.AddGroup("&Nueva", ID_GROUP_NEW)
        
    Set Control = GroupNew.Add(xtpControlButton, ID_GROUP_NEW_APPOINTMENT, "&Evento")
    Control.Enabled = False
    Control.visible = False
    Set Control = GroupNew.Add(xtpControlButton, ID_GROUP_NEW_MEETING, "&Cita")
    Control.Enabled = True
    Set Control = GroupNew.Add(xtpControlButton, ID_GROUP_SHARE_SHARE, "&Imprimir")
    Control.Enabled = True
    
    
    
    '------------------------------------
    'Set ControlNew_NewItems = GroupNew.Add(xtpControlButtonPopup, ID_GROUP_NEW_ITEMS, "New &Items")
    '    Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_APPOINTMENT, "Evento")
    '    Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_ALLDAY, "E&vento todo el dia")
    '    Control.BeginGroup = True
    'ControlNew_NewItems.KeyboardTip = "V"
    
    Set GroupGoTo = TabCalendarHome.Groups.AddGroup("I&r a", ID_GROUP_GOTO)
    Set Control = GroupGoTo.Add(xtpControlButton, ID_GROUP_GOTO_TODAY, "&Hoy")
    Set Control = GroupGoTo.Add(xtpControlButton, ID_GROUP_GOTO_NEXT7DAYS, "Próximos &7 dias ")
    GroupGoTo.ShowOptionButton = True
    GroupGoTo.ControlGroupOption.Caption = "Ir a (Ctrl+G)"
    GroupGoTo.ControlGroupOption.ToolTipText = "Ir a (Ctrl+G)"
    GroupGoTo.ControlGroupOption.DescriptionText = "Ir a fecha especificada."
    
    Set GroupArrange = TabCalendarHome.Groups.AddGroup("Vista", ID_GROUP_ARRANGE2)
    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_DAY, "&Dia vista")
    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_WORK_WEEK, "Samana &trabajo")
    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_WEEK, "Sema&na vista")
    Set ControlArrange_Month = GroupArrange.Add(xtpControlSplitButtonPopup, ID_GROUP_ARRANGE_MONTH, "Mes")
            Set Control = ControlArrange_Month.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_ARRANGE_MONTH_LOW, "Ver detalle")
            Control.ToolTipText = "Muestra solo eventos todo el dia."
            Control.DescriptionText = Control.ToolTipText
            Set Control = ControlArrange_Month.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_ARRANGE_MONTH_MEDIUM, "Detalle &Medio")
            Control.ToolTipText = "Eventos todo el dia y si esta libre el dia o tiene eventos."
            Control.DescriptionText = Control.ToolTipText
            Set Control = ControlArrange_Month.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_ARRANGE_MONTH_HIGH, "Detalle &Alto")
            Control.ToolTipText = "Muestra todo."
            Control.DescriptionText = Control.ToolTipText

'    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_SCHEDULE_VIEW, "Schedule View")
'    GroupArrange.ShowOptionButton = True
'    GroupArrange.ControlGroupOption.Caption = "Calendar Options"
'    GroupArrange.ControlGroupOption.ToolTipText = "Calendar Options"
'    GroupArrange.ControlGroupOption.DescriptionText = "Change the settings for calendars, meetings and time zones."
'
'
  
    
End Sub





Private Sub CreateRibbon()
    Dim RibbonBar As RibbonBar
    
    If RibbonSeHaCreado Then Exit Sub
        
    
    
    Set RibbonBar = CommandBars.AddRibbonBar("The Ribbon")
    RibbonBar.EnableDocking xtpFlagStretched
    
    RibbonBar.AllowQuickAccessCustomization = False
    RibbonBar.ShowQuickAccessBelowRibbon = False
    RibbonBar.ShowGripper = False
    
    RibbonBar.AllowMinimize = False
    RibbonBar.AddSystemButton
    
    RibbonBar.SystemButton.IconId = ID_SYSTEM_ICON
    RibbonBar.SystemButton.Caption = "&Menu"
    RibbonBar.SystemButton.Style = xtpButtonCaption
End Sub

Private Sub CreateRibbonOptions()

    CommandBars.EnableActions
    If RibbonSeHaCreado Then Exit Sub
    
    CommandBars.Actions.Add ID_OPTIONS_STYLEBLUE2010, "Office 2010 Blue", "Office 2010 Blue", "Office 2010 Blue", "Themes"
    CommandBars.Actions.Add ID_OPTIONS_STYLESILVER2010, "Office 2010 Silver", "Office 2010 Silver", "Office 2010 Silver", "Themes"
    CommandBars.Actions.Add ID_OPTIONS_STYLEBLACK2010, "Office 2010 Black", "Office 2010 Black", "Office 2010 Black", "Themes"

    Dim Control As CommandBarControl, ControlAbout As CommandBarControl
    Dim ControlPopup As CommandBarPopup, ControlOptions As CommandBarPopup
         
    Set ControlOptions = RibbonBar.Controls.Add(xtpControlPopup, 0, "Opciones")
    ControlOptions.Flags = xtpFlagRightAlign
    
    Set Control = ControlOptions.CommandBar.Controls.Add(xtpControlPopup, 0, "Styles")
    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLEBLUE2010, "Office 2010 Blue"
    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLESILVER2010, "Office 2010 Silver"
    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLEBLACK2010, "Office 2010 Black"
    
    Set ControlPopup = ControlOptions.CommandBar.Controls.Add(xtpControlPopup, 0, "Tamaño fuente", -1, False)
    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_SYSTEM, "Sistema", -1, False
    Set Control = ControlPopup.CommandBar.Controls.Add(xtpControlRadioButton, ID_OPTIONS_FONT_NORMAL, "Normal", -1, False)
    Control.BeginGroup = True
    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_LARGE, "Grande", -1, False
    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_EXTRALARGE, "Extra grande", -1, False
    Set Control = ControlPopup.CommandBar.Controls.Add(xtpControlButton, ID_OPTIONS_FONT_AUTORESIZEICONS, "Ajustar Icons", -1, False)
    Control.BeginGroup = True
    
    'ControlOptions.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_RTL, "Right To Left"
    ControlOptions.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_ANIMATION, "Animation   "
    
    Set Control = AddButton(RibbonBar.Controls, xtpControlButton, ID_RIBBON_MINIMIZE, "Minimizar la barra", False, "Muestra solo los titulos del menu principal.")
    Control.Flags = xtpFlagRightAlign
    
    Set Control = AddButton(RibbonBar.Controls, xtpControlButton, ID_RIBBON_EXPAND, "Expandir la barra", False, "Muestra todos los elementos del menu.")
    Control.Flags = xtpFlagRightAlign
        
    Set ControlAbout = RibbonBar.Controls.Add(xtpControlButton, ID_APP_ABOUT, "&Acerca de")
    ControlAbout.Flags = xtpFlagRightAlign Or xtpFlagManualUpdate
    

        
End Sub








'*************************************************************************
'*************************************************************************
'*************************************************************************
'
'       CARGA menus en Ribbon
'
'




Public Sub CargaMenu(AntiguoTab As Integer)
Dim RN As ADODB.Recordset




    Set RN = New ADODB.Recordset
    Set Rn2 = New ADODB.Recordset
    On Error GoTo eCargaMenu
    

    If RibbonSeHaCreado Then RibbonBar.RemoveAllTabs
    
    Cad = "Select * from menus where aplicacion = 'ariconta' and padre =0 ORDER BY padre,orden "
    RN.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
    
        
        If Not BloqueaPuntoMenu(RN!Codigo, "ariconta") Then
             Habilitado = True
             
             If Not MenuVisibleUsuario(DBLet(RN!Codigo), "ariconta") Then
                 Habilitado = False
             Else
         
                 If (MenuVisibleUsuario(DBLet(RN!padre), "ariconta") And DBLet(RN!padre) <> 0) Or DBLet(RN!padre) = 0 Then
                     'OK todo habilitado
                 Else
                     Habilitado = False
                 End If
             End If
             
            
                
            If Habilitado Then
                
                Select Case RN!Codigo
                Case 1
                    '1   "CONFIGURACION"
                    CargaMenuConfiguracion RN!Codigo
                Case 2
                    '2   "DATOS GENERALES"
                    CargaMenuDatosGenerales RN!Codigo
                Case 3
                    '3   "DIARIO"
                    CargaMenuDiarios RN!Codigo
                Case 4
                    '4   "FACTURAS"
                    CargaMenuFacturas RN!Codigo
                Case 5
                    '5   "INMOVILIZADO"
                    CargaMenuInmovilizado RN!Codigo
                Case 6
                    '6   "CARTERA DE COBROS"
                    CargaMenuTesoreriaCobros RN!Codigo
                Case 7
                    
                Case 8
                    '8   "CARTERA DE PAGOS"
                    CargaMenuTesoreriaPagos RN!Codigo
                Case 9
                    '9   "INFORMES TESORERIA"
                     CargaMenuTesoreriaInformes RN!Codigo
                Case 10
                    '10  "ANALÍTICA"
                    'Va dentro de diario
                    'UNa solapa para el
                    CargaMenuAnaliticaPResupuestaria RN!Codigo
                Case 11
                    '11  "PRESUPUESTARIA"
                    CargaMenuAnaliticaPResupuestaria RN!Codigo
                Case 12
                
                Case 13
                    '13  "CIERRE EJERCICIO"
                    CargaMenuCierreEjercicio RN!Codigo
                     
                Case 14
                    '14  "UTILIDADES"
                    CargaMenuUtilidades RN!Codigo
                Case Else
                    MsgBox "Menu no tratado"
                    End
                End Select
                
            End If
                                                 
        End If  'de habilitado el padre
    
        RN.MoveNext
    Wend
    RN.Close
                        
    PonerTabPorDefecto AntiguoTab
    
eCargaMenu:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
    
    Set TabNuevo = Nothing
    Set GroupNew = Nothing
    Set Control = Nothing
    Set RN = Nothing
    Set Rn2 = Nothing
End Sub

Private Sub PonerTabPorDefecto(AntiguoTabSeleccionado As Integer)
Dim Anterior As Integer

    On Error Resume Next
    
    If AntiguoTabSeleccionado < 0 Then
        Anterior = vUsu.TabPorDefecto
    Else
        Anterior = AntiguoTabSeleccionado
    End If
    
    Cad = ""
    For i = 0 To RibbonBar.TabCount - 1
        J = RibbonBar.Tab(i).Id
        'Debug.Print J & " " & RibbonBar.Tab(i).Caption
        If J = Anterior Then
            
            RibbonBar.Tab(i).visible = True
            RibbonBar.Tab(i).Selected = True
            Set RibbonBar.SelectedTab = RibbonBar.Tab(i)
            Cad = "OK"
            Exit For
        End If
    Next
    If Cad = "" Then
        
        For J = RibbonBar.TabCount To 1 Step -1
            RibbonBar.Tab(J - 1).visible = True
            RibbonBar.Tab(J - 1).Selected = True
        Next J
    End If

    Err.Clear
End Sub

Private Sub CargaMenuConfiguracion(IdMenu As Integer)

        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Configuracion")
        TabNuevo.Id = CLng(IdMenu)
        Set GroupNew = TabNuevo.Groups.AddGroup("", 1000000)
        
       
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
         
           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!padre), "ariconta") Then Habilitado = False
                End If
           
           
                If Rn2!Codigo = ID_ConfigurarBalances Then
                    Set ControlNew_NewItems = GroupNew.Add(xtpControlButtonPopup, Rn2!Codigo, Rn2!Descripcion)
                    Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_ConfigurarBalances1, "Balances")
                    Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_ConfigurarBalances2, "Ratios")
                    If vUsu.Login = "root" Then Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_ConfigurarBalances3, "Personalizables")
                    
                    'Personalizan
                    ControlNew_NewItems.Enabled = Habilitado
                Else
                    Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Control.Enabled = Habilitado
                End If
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close
        
        'color Categorias  eventos
        If Not GroupNew Is Nothing Then
            Set Control = GroupNew.Add(xtpControlButton, 199, "Categorias calendario")
        End If
        Set GroupNew = Nothing
End Sub






Private Sub CargaMenuDatosGenerales(IdMenu As Integer)
Dim SegundoGrupo As RibbonGroup
Dim B As Boolean


        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Datos generales")
        TabNuevo.Id = CLng(IdMenu)
        
        
        'En este llevaremos dos solapas, tesoreria y contabilidad (no le ponemos nombres)
        Cad = CStr(IdMenu * 100000)
        'If vEmpresa.TieneContabilidad Then Set GroupNew = TabNuevo.Groups.AddGroup("", Cad & "0")
        Set GroupNew = TabNuevo.Groups.AddGroup("", Cad & "0")
        If vEmpresa.TieneTesoreria Then Set SegundoGrupo = TabNuevo.Groups.AddGroup("", Cad & "1")
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu
        If Not vEmpresa.TieneTesoreria Then
            'SOLO CONTA
            Cad = Cad & " AND tipo <> 1"  '=0
        Else
                                                            'solo tesoreria
            If Not vEmpresa.TieneContabilidad Then
                Cad = Cad & " AND  tipo <> 0 "                  '=1
            End If
        End If
        Cad = Cad & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
         
           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
           
'                Cargamos = True
'                If Not vEmpresa.TieneContabilidad Then
'                    If Rn2!Tipo = 0 Then Cargamos = False
'                End If
'                If Not vEmpresa.TieneTesoreria Then
'                    If Rn2!Tipo = 1 Then Cargamos = False
'                End If
                
'                If Cargamos Then
                     Habilitado = True
                     If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
                         Habilitado = False
                     Else
                         If Not MenuVisibleUsuario(DBLet(Rn2!padre), "ariconta") Then Habilitado = False
                     End If
                
                    
                         
                     If Rn2!Tipo = 1 Then
                         Set Control = SegundoGrupo.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                     Else
                         Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                     End If
                      
                     Control.Enabled = Habilitado
                     ' Set ControlNew_NewItems = GroupNew.Add(xtpControlButtonPopup, ID_GROUP_NEW_ITEMS, "New &Items")
                     '     Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_APPOINTMENT, "&Appointment")
                     '     Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_ALLDAY, "All Day E&vent")
                     '     Control.BeginGroup = True
                     ' ControlNew_NewItems.KeyboardTip = "V"
'                End If
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close

         Set GroupNew = Nothing
End Sub


Private Sub CargaMenuDiarios(IdMenu As Integer)
Dim GrupSald As RibbonGroup
Dim GrOtro As RibbonGroup
Dim GrConsoli As RibbonGroup
Dim OtroCon

        If Not vEmpresa.TieneContabilidad Then Exit Sub

        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Diario")
        TabNuevo.Id = CLng(IdMenu)
        
        Cad = CStr(IdMenu * 100000)
        Set GroupNew = TabNuevo.Groups.AddGroup("ASIENTOS", Cad & "0")
        Set GrupSald = TabNuevo.Groups.AddGroup("BALANCES", Cad & "1")
        Set GrOtro = TabNuevo.Groups.AddGroup("", Cad & "2")
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
        
           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!padre), "ariconta") Then Habilitado = False
                End If
                

                
                Select Case Rn2!Codigo
                Case 301, 303, 304, 314, 211
                    Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                Case 306, 307, 308, 309
                    Set Control = GrupSald.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    
                'Consolidado
                Case 315
                    Set GrConsoli = TabNuevo.Groups.AddGroup("CONSOLIDADO", Cad & "4")
                    
                    Set ControlNew_NewItems = GrConsoli.Add(xtpControlButtonPopup, Rn2!Codigo, "Informes") 'Rn2!Descripcion
                    'Set Control = GrConsoli.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    
                    Set OtroCon = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_ConsoSumasSaldos, "Balance sumas y saldos")
                    Set OtroCon = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_ConsoCtaExplota, "Cuenta explotacion")
                    Set OtroCon = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_ConsoPyG, "Pérdidas y ganancias")
                    Set OtroCon = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_ConsoSitu, "Balance de situacion")
                    'Dos consolidados
                    
                    ControlNew_NewItems.Enabled = Habilitado
                    
                Case Else
                    Set Control = GrOtro.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                                        
                End Select
                
                
                Control.Enabled = Habilitado
                
              
              
              
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close
    Set GrupSald = Nothing
    Set GrOtro = Nothing
     Set GrConsoli = Nothing
End Sub


Private Sub CargaMenuFacturas(IdMenu As Integer)
Dim GropCli As RibbonGroup
Dim GrupPag As RibbonGroup
Dim Consoli As RibbonGroup
Dim OpsAseg As RibbonGroup
Dim Insertado As Boolean
Dim B As Boolean

        If Not vEmpresa.TieneContabilidad Then Exit Sub
        
        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Facturas")
        TabNuevo.Id = CLng(IdMenu)
        
        Cad = CStr(IdMenu * 100000)
        Set GropCli = TabNuevo.Groups.AddGroup("EMITIDAS", Cad & "0")
        Set GrupPag = TabNuevo.Groups.AddGroup("RECIBIDAS", Cad & "1")
        Set GroupNew = TabNuevo.Groups.AddGroup("I.V.A.", Cad & "2")
            

'
'        401 "Facturas Emitidas" 14
'        402 "Libro Facturas Emitidas"   16
'        403 "Relación Clientes por cuenta"  0
'        404 "Facturas Recibidas"    15
'        405 "Libro Facturas Recibidas"  17
'        406 "Relacion Proveedores por cuenta"   0
'        408 "Modelo 303"    0
'        409 "Modelo 340"    0
'        410 "Modelo 347"    0
'        411 "Modelo 349"    0
'        412 "Liquidacion I.V.A."    18
'        413 consolidad
'        414 ID_AseguClientes
'        415 AseguComunicaSeguro
'        416 SII
'        417 ID_AseguComunicaSeguroAvisos
'        418 ID_AseguComprobarVtos
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
        
        
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
        
           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!padre), "ariconta") Then Habilitado = False
                End If
            End If
            
            Insertado = True
            Select Case Rn2!Codigo
            Case 401, 402, 403
                Set Control = GropCli.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
            Case 404, 405, 406
                Set Control = GrupPag.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
            
            Case 413
                Set Consoli = TabNuevo.Groups.AddGroup("CONSOLIDADO", CStr(IdMenu * 100000) & "2")
                Set Control = Consoli.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                 
                 
            Case ID_AseguClientes, ID_AseguComunicaSeguro, ID_AseguComunicaSeguroAvisos, ID_AseguComprobarVtos
                 If Not vEmpresa.TieneTesoreria Then
                    Insertado = False
                 Else
                    If vParamT.TieneOperacionesAseguradas Then
                           If OpsAseg Is Nothing Then Set OpsAseg = TabNuevo.Groups.AddGroup("OP. ASEGURADAS", CStr(IdMenu * 100000) & "4")
                           Set Control = OpsAseg.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Else
                       Insertado = False
                    End If
                End If
            Case Else
                B = True
                If Rn2!Codigo = ID_SII Then
                    
                    If vParam.SIITiene Then
                        If vUsu.Nivel > 0 Then B = False
                    Else
                        B = False
                        
                    End If
                End If
                If Not B Then Habilitado = False
                
                Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                
            End Select
            
            
            Cad = "NO"
            If Insertado Then Control.Enabled = Habilitado
           
            Rn2.MoveNext
        Wend
        Rn2.Close

    
        
        
        
End Sub



Private Sub CargaMenuInmovilizado(IdMenu As Integer)

        If Not vEmpresa.TieneContabilidad Then Exit Sub

        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Inmovilizado")
        TabNuevo.Id = CLng(IdMenu)
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
        
           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!padre), "ariconta") Then Habilitado = False
                End If
            End If
            
            If Cad = "" Then Set GroupNew = TabNuevo.Groups.AddGroup("", CStr(IdMenu * 100000) & "0")
            Cad = "NO"
            'Set Control = GroupNew.Add(xtpControlButton, ID_GROUP_NEW_APPOINTMENT, "&New Appointment")
            Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
            Control.Enabled = Habilitado
            
           ' Set ControlNew_NewItems = GroupNew.Add(xtpControlButtonPopup, ID_GROUP_NEW_ITEMS, "New &Items")
           '     Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_APPOINTMENT, "&Appointment")
           '     Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_ALLDAY, "All Day E&vent")
           '     Control.BeginGroup = True
           ' ControlNew_NewItems.KeyboardTip = "V"
        
            
            Rn2.MoveNext
        Wend
        Rn2.Close


End Sub




Private Sub CargaMenuTesoreriaCobros(IdMenu As Integer)
Dim GrupCob As RibbonGroup
Dim GrupRem As RibbonGroup
        
'    601 "Cartera de Cobros"
'    602 "Informe Cobros Pendientes"
'    604 "Realizar Cobro"
'    606 "Compensaciones"
'    607 "Compensar cliente"
'    608 "Reclamaciones"
'    609 "Remesas"
'    610 "Informe Impagados"
'    611 "Recepción Talón-Pagaré"
'    612 "Remesas Talón-Pagaré"
'    613 "Norma 57 - Pagos ventanilla"
'    614 "Transferencias Abonos"
'615 Anticipos facturas
        
        If Not vEmpresa.TieneTesoreria Then Exit Sub
        
        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Tesoreria")
        TabNuevo.Id = CLng(IdMenu)
        NumRegElim = TabNuevo.Index
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        
        
        'Creamos los tres grupos
        Cad = CStr(IdMenu * 100000)
        Set GrupCob = TabNuevo.Groups.AddGroup("COBROS", Cad & "0")
        Set GrupRem = TabNuevo.Groups.AddGroup("REMESAS", Cad & "1")
        Set GroupNew = TabNuevo.Groups.AddGroup("", Cad & "2")
        
        
        While Not Rn2.EOF
        
           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!padre), "ariconta") Then Habilitado = False
                End If
            End If
            


            Select Case Rn2!Codigo
            Case 601, 602, 604, 607, 608, 610, 613, 614, ID_AnticipoFacturas
                'Solapa cobros
                Set Control = GrupCob.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
            
            Case 609, 611, 612
                'Solapa remesas
                Set Control = GrupRem.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                
            Case Else
                Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                
            
            
            End Select
            
            
            Control.Enabled = Habilitado
            
           ' ControlNew_NewItems.KeyboardTip = "V"
        
            
            Rn2.MoveNext
        Wend
        Rn2.Close


End Sub



Private Sub CargaMenuTesoreriaPagos(IdMenu As Integer)

        
        If Not vEmpresa.TieneTesoreria Then Exit Sub
        
        
    
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        
        
        'Los pagos se gargan sobre la solapa de TESORERIA
        
        Set GroupNew = RibbonBar.Tab(NumRegElim).Groups.AddGroup("PAGOS", CStr(IdMenu * 100000) & "0")
        
        
        While Not Rn2.EOF
        
'            801 "Cartera de Pagos"  5
'            802 "Informe Pagos pendientes"  19
'            803 "Informe Pagos bancos"  0
'            804 "Realizar Pago" 24
'            805 "Transferencias"    0
'            806 "Pagos domiciliados"    0
'            807 "Gastos Fijos"  0
'            809 "Compensar proveedor"   0
'            810 "Confirming"    0
        
           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!padre), "ariconta") Then Habilitado = False
                End If
            End If
            
            Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
            
            Control.Enabled = Habilitado
            
 
            
            Rn2.MoveNext
        Wend
        Rn2.Close


End Sub


Private Sub CargaMenuTesoreriaInformes(IdMenu As Integer)

        
        If Not vEmpresa.TieneTesoreria Then Exit Sub
        
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu
        'De momento NO cargamos el 904
        Cad = Cad & " AND codigo <>904  ORDER BY padre,orden"
            
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        
        
        'Los informes se gargan sobre la solapa de TESORERIA
        
        Set GroupNew = RibbonBar.Tab(NumRegElim).Groups.AddGroup("INFORMES", CStr(IdMenu * 100000) & "0")
        
        
        While Not Rn2.EOF
'       901 "ariconta"  9   "Informe por NIF *" 1   1   0
'       902 "ariconta"  9   "Informe por cuenta *"  2   1   0
'       903 "ariconta"  9   "Situación Tesoreria *" 3   1   29
'       904 "ariconta"  9   "Memoria Plazos de pago *"  4   1   0
'     ID_InformeporNIF ID_Informeporcuenta ID_SituaciónTesoreria
           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!padre), "ariconta") Then Habilitado = False
                End If
            End If
            
            Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
            
            Control.Enabled = Habilitado
            
 
            
            Rn2.MoveNext
        Wend
        Rn2.Close


End Sub




Private Sub CargaMenuAnaliticaPResupuestaria(IdMenu As Integer)

        
        
        If Not vEmpresa.TieneContabilidad Then Exit Sub
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        
        
        'Los pagos se gargan sobre la solapa de diario
        Set TabNuevo = RibbonBar.FindTab(3)
        If TabNuevo Is Nothing Then
            Set TabNuevo = RibbonBar.FindTab(15)
            If TabNuevo Is Nothing Then
                
                Set TabNuevo = RibbonBar.InsertTab(15, "ANALITICA-PRESUPUESTO")
                TabNuevo.Id = 15
                
            End If
        End If
        Cad = CStr(IdMenu * 100000) & "0"
        Set GroupNew = TabNuevo.Groups.AddGroup(IIf(IdMenu = 10, "ANALITICA", "PRESUPUESTOS"), Cad)
        
        
        While Not Rn2.EOF
        
'            801 "Cartera de Pagos"  5
'            802 "Informe Pagos pendientes"  19
'            803 "Informe Pagos bancos"  0
'            804 "Realizar Pago" 24
'            805 "Transferencias"    0
'            806 "Pagos domiciliados"    0
'            807 "Gastos Fijos"  0
'            809 "Compensar proveedor"   0
'            810 "Confirming"    0
        
           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!padre), "ariconta") Then Habilitado = False
                End If
            End If
            

            
            
            Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
            
            Control.Enabled = Habilitado
            
           ' ControlNew_NewItems.KeyboardTip = "V"
        
            
            Rn2.MoveNext
        Wend
        Rn2.Close


End Sub



Private Sub CargaMenuCierreEjercicio(IdMenu As Integer)
Dim GropCli As RibbonGroup
Dim GrupPag As RibbonGroup
        
        If Not vEmpresa.TieneContabilidad Then Exit Sub
        
        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Cierre ejercicio")
        TabNuevo.Id = CLng(IdMenu)
        
        Set GroupNew = TabNuevo.Groups.AddGroup("", 13000001)
    

        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
        
           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!padre), "ariconta") Then Habilitado = False
                End If
            End If
'        1301    "Renumeración de asientos"  0
'        1303    "Cierre de Ejercicio"   0
'        1304    "Deshacer cierre"   0
'        1306    "Diario Oficial"    0
'        1308    "Presentación Telemática de Libros" 0
'        1309    "Memoria Plazos de pago"    0
            
            
           
            Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
            Control.Enabled = Habilitado
            ' ControlNew_NewItems.KeyboardTip = "V"
         
            Rn2.MoveNext
        Wend
        Rn2.Close


End Sub

Private Function DevulevePosicionUtilidades(Id As Integer) As Integer
    Select Case Id
    Case ID_Traspasodecuentasenapuntes
        DevulevePosicionUtilidades = 1
    Case ID_Renumerarregistrosproveedor
        DevulevePosicionUtilidades = 2
    Case ID_Aumentardígitoscontables
        DevulevePosicionUtilidades = 3
    Case ID_TraspasocodigosdeIVA
        DevulevePosicionUtilidades = 4
    Case Else
        'ID_Accionesrealizadas
        DevulevePosicionUtilidades = 5
    End Select
End Function

Private Sub CargaMenuUtilidades(IdMenu As Integer)
Dim Col As Collection

        
        
        
        'Este veremos si tiene alguna utilidad activa. Si es asi, crearemos la solapa, si no nada
        '.......................................................................
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        Set Col = New Collection
        While Not Rn2.EOF
           i = i + 1
           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!padre), "ariconta") Then Habilitado = False
                End If
            End If
            
            If Rn2!Codigo = 1414 Then
                If Not vEmpresa.TieneTesoreria Then
                    Habilitado = False
                Else
                    If vParamT.FormaPagoInterTarjeta < 0 Then Habilitado = False
                End If
            End If
            
            Col.Add Abs(Habilitado) & "|" & Rn2!Codigo & "|" & Rn2!Descripcion & "|"
            If Habilitado Then Cad = "S"
            
            Rn2.MoveNext
        Wend
        Rn2.Close
        
            '1408    "Traspaso de cuentas en apuntes"
            '1409    "Renumerar registros proveedor"
            '1410    "Aumentar dígitos contables"
            '1411    "Traspaso códigos de I.V.A."
            '1412    "Acciones realizadas"
            '1413    Importar fras cliente
            '1414    importacon facturas (de momento consum)
            
        'Ya puedo utilizar numregelim
        If Cad <> "" Then
            'OK creamos solapa y demas
            'Creamos la TAB
            Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Utilidades")
            TabNuevo.Id = CLng(IdMenu)
            Set GroupNew = TabNuevo.Groups.AddGroup("", 14000001)
            For NumRegElim = 1 To Col.Count
                Habilitado = CStr(RecuperaValor(Col.Item(NumRegElim), 1)) = "1"
                Set Control = GroupNew.Add(xtpControlButton, CLng(RecuperaValor(Col.Item(NumRegElim), 2)), CStr(RecuperaValor(Col.Item(NumRegElim), 3)))
                Control.Enabled = Habilitado
            Next
                
            
        End If
        

Set Col = Nothing
End Sub






'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**********************************************************f****************************************************
Private Sub AbrirFormularios(Accion As Long)
  
   
   ' If Accion <> ID_SII Then AbrirFormSII_2 False
  
    
    Select Case Accion
        Case 101 ' empresa
            frmempresa.Show vbModal
        Case 102 ' parametros contabilidad
            If Not (vEmpresa Is Nothing) Then
                frmparametros.Show vbModal
            End If
        Case 103 ' parametros tesoreria
        Case 104 ' contadores
            Screen.MousePointer = vbHourglass
            If vUsu.Nivel = 0 Then frmContadores.Show vbModal
            Screen.MousePointer = vbDefault
        Case 105 ' usuarios
            frmMantenusu.Show vbModal
        Case 106 ' informes
            frmCrystal.Show vbModal
        Case 107 ' crear nueva empresa
            If vUsu.Nivel > 1 Then Exit Sub
            CadenaDesdeOtroForm = ""
            frmCentroControl.Opcion = 2
            frmCentroControl.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                If Val(CadenaDesdeOtroForm) > 0 Then
                    Cad = "update ariconta" & CadenaDesdeOtroForm & ".tiposdiario set numdiari=numdiari where numdiari<0"
                    If EjecutaSQL(Cad) Then
                        Cad = ""
                        CambiarEmpresa CInt(CadenaDesdeOtroForm)
                    End If
                    Cad = ""
                End If
            End If
            
            
        'Case 108 'Configurar Balances
        'ID_ConfigurarBalances1 = 120   'balances
        'ID_ConfigurarBalances2 = 121   'ratios
        'ID_ConfigurarBalances2 = 122   'personalizables
        Case ID_ConfigurarBalances1, ID_ConfigurarBalances2, ID_ConfigurarBalances3
            Screen.MousePointer = vbHourglass
            'frmColBalan2.TipoVista = 0 'Pyg   y situacion
            If Accion = 120 Then
                frmColBalan2.TipoVista = 0 'Pyg   y situacion
            ElseIf Accion = 121 Then
                frmColBalan2.TipoVista = 1 'ratios
            Else
                frmColBalan2.TipoVista = 2 'personalizales
            End If
            frmColBalan2.Show vbModal, Me
            
        Case 199
             frmCalendarCategorias.Show vbModal
            
            
        Case 201 ' plan contable
            Screen.MousePointer = vbHourglass
            frmColCtas.ConfigurarBalances = 0
            frmColCtas.DatosADevolverBusqueda = ""
            frmColCtas.Show vbModal, Me
        Case 202 ' tipos de diario
            Screen.MousePointer = vbHourglass
            frmTiposDiario.Show vbModal
        Case 203 ' conceptos
            Screen.MousePointer = vbHourglass
            frmConceptos.Show vbModal
        Case 204 ' tipos de iva
            Screen.MousePointer = vbHourglass
            frmIva.Show vbModal
        Case 205 ' tipos de pago
            Screen.MousePointer = vbHourglass
            frmTipoPago.Show vbModal
        Case 206 ' formas de pago
            Screen.MousePointer = vbHourglass
            frmFormaPago.Show vbModal
        Case 207 ' bancos
            Screen.MousePointer = vbHourglass
            frmBanco.Show vbModal
        Case 208 ' bic
            Screen.MousePointer = vbHourglass
            frmBic.Show vbModal
        Case 209 ' agentes
            Screen.MousePointer = vbHourglass
            frmAgentes.Show vbModal
        Case 210 ' departamentos
        Case 211 ' asientos predefinidos
            Screen.MousePointer = vbHourglass
            frmAsiPre.Show vbModal
        Case 212 ' cartas de reclamacion
            Screen.MousePointer = vbHourglass
            frmCartas.Show vbModal
        
        Case 301 ' asientos
            Screen.MousePointer = vbHourglass
            frmAsientosHco.Asiento = ""
            frmAsientosHco.DesdeNorma43 = 0
            frmAsientosHco.Show vbModal
        Case 303 ' extractos
            Screen.MousePointer = vbHourglass
            frmConExtr.EjerciciosCerrados = False
            frmConExtr.cuenta = ""
            frmConExtr.Show vbModal
        Case 304 ' punteo
            Screen.MousePointer = vbHourglass
            frmPuntear.EjerciciosCerrados = False
            frmPuntear.Show vbModal
        Case 305 ' reemision de diarios
'            AbrirListado 6, False
        Case 306 ' sumas y saldos
            frmInfBalSumSal.Show vbModal
            
        Case 307 ' cuenta de explotacion
            frmInfCtaExplo.Show vbModal
            
        Case 308 ' balance de situacion
            frmInfBalances.Opcion = 0
            frmInfBalances.Show vbModal
            
        Case 309 ' perdidas y ganancias
            frmInfBalances.Opcion = 1
            frmInfBalances.Show vbModal

        Case 310 ' totales por concepto
            frmInfTotCtaCon.Show vbModal
        Case 311 ' evolucion de saldos
            frmInfEvolSal.Show vbModal
        Case 312 ' ratios y graficas
            frmInfRatios.Show vbModal
        Case 314 ' puntero extracto bancario
            frmPunteoBanco.Show vbModal
        
        Case 315
            'frmInfBalSumSalConso.Show vbModal
        Case ID_ConsoSumasSaldos, ID_ConsoCtaExplota
            frmInfBalSumSalConso.Opcion = IIf(Accion = ID_ConsoSumasSaldos, 0, 1)
            frmInfBalSumSalConso.Show vbModal
        
        Case ID_ConsoPyG, ID_ConsoSitu
            frmInfConsBalan.Opcion = IIf(Accion = ID_ConsoPyG, 1, 0)
            frmInfConsBalan.Show vbModal
        
        'Gradficas
        Case ID_GraficosChart
            frmGraficos.Show vbModal
        
        
        
        
        Case 401 ' emitidas
            Screen.MousePointer = vbHourglass
            frmFacturasCli.Factura = ""
            frmFacturasCli.Show vbModal
        Case 402 ' libro emitidas
            frmFacturasCliListado.Show vbModal
        Case 403 ' relacion clientes por cuenta
            frmFacturasCliCtaVtas.Show vbModal
        Case 404 ' recibidas
            Screen.MousePointer = vbHourglass
            frmFacturasPro.Factura = ""
            frmFacturasPro.Show vbModal
        Case 405 ' libro recibidas
            frmFacturasProListado.Show vbModal
        Case 406 ' relacion proveedores por cuenta
            frmFacturasProCtaGastos.Show vbModal
        Case 407 ' liquidacion iva
'            AbrirListado 12, False
        Case 408 ' certificado iva
            frmModelo303.Opcionlistado = 0
            frmModelo303.Show vbModal
        Case 409 ' modelo 340
            frmModelo340.Show vbModal
        Case 410 ' modelo 347
            frmModelo347.Show vbModal
        Case 411 ' modelo 349
            frmModelo349.Show vbModal
        Case 412 ' liquidacion de iva
            frmHcoLiqIVA.Show vbModal
            
        Case ID_FrasConso   '413
            frmConsolidadoFras.Show vbModal
            
        Case ID_AseguClientes   '414
            frmSegurosListClientes.Show vbModal
            
        Case ID_AseguComunicaSeguro, ID_AseguComunicaSeguroAvisos '415,417
            frmSegurosListComunicacion.numero = IIf(Accion = ID_AseguComunicaSeguroAvisos, 1, 0)
            frmSegurosListComunicacion.Show vbModal
        
        Case ID_AseguComprobarVtos
            ComprobarOperacionesAseguradas False
            
        Case ID_SII
            AbrirFormSII True
        
        
        
        Case 502 ' conceptos
            Screen.MousePointer = vbHourglass
            frmInmoConceptos.Show vbModal
        Case 503 ' elementos
            frmInmoElto.DatosADevolverBusqueda = ""
            frmInmoElto.Show vbModal
        Case 505 ' estadistica
            frmInmoInfEst.Show vbModal
        Case 507 ' historico inmovilizado
            Screen.MousePointer = vbHourglass
            frmInmoHco.Show vbModal
        Case 508 ' simulacion
            frmInmoSimu.Show vbModal
        Case 509 ' calculo y contabilizacion
            frmInmoGenerar.Opcion = 2
            frmInmoGenerar.Show vbModal
        Case 510 ' deshacer amortizacion
            frmInmoDeshacer.Show vbModal
        Case 511 ' venta-baja inmmovilizado
            frmInmoVenta.Opcion = 3
            frmInmoVenta.Show vbModal
        Case 601 ' cartera de cobros
            frmTESCobros.Show vbModal
        Case 602 ' informe de cobros pendientes
            frmTESCobrosPdtesList.Show vbModal
        
        Case 603 ' impresion de recibos
            frmTESImpRecibo.documentoDePago = ""
            frmTESImpRecibo.Show vbModal
        Case 604 ' realizar cobro
            With frmTESRealizarCobros
               
                '--.vSQL = SQL
                .Regresar = False
                .Cobros = True
                .ContabTransfer = False
                .SegundoParametro = ""
                'Los textos
'                .vTextos = Text1(2).Text & "|" & Me.txtCta(0).Text & " - " & Me.txtDescCta(0).Text & "|" & SubTipo & "|"
                
                'Marzo2013   Cobramos un solo cliente
                'Aparecera un boton para traer todos los cobros
                '.CodmactaUnica = "4300000001" 'Trim(txtCtaNormal(9).Text)
                .Show vbModal
            End With
        
        Case 606 ' compensaciones
            frmTESCompensaciones.Show vbModal
        Case 607 ' compensar cliente
            CadenaDesdeOtroForm = ""
            frmTESCompensaAboCli.Show vbModal
        Case 608 ' reclamaciones
            frmTESReclamaCli.Show vbModal
        Case 609 ' remesas
            frmTESRemesas.Tipo = 1 ' efectos
            frmTESRemesas.Show vbModal
        Case 610 ' Informe Impagados
            frmTESCobrosDevList.Show vbModal
        Case 611 ' Recepción Talón-Pagaré
            frmTESRecepcionDoc.Show vbModal
        Case 612 ' Remesas Talón-Pagaré
            frmTESRemesasTP.Tipo = 2 ' talon pagare
            frmTESRemesasTP.Show vbModal
            
        Case 613 ' Norma 57: Pago por ventanilla
            frmTESNorma57.Opcion = 42
            frmTESNorma57.Show vbModal
            
        Case 614 ' transferencia abonos

            frmTESTransferencias.TipoTrans2 = 0
            frmTESTransferencias.Show vbModal
            
        Case ID_AnticipoFacturas
            frmTESTransferencias.TipoTrans2 = 4
            frmTESTransferencias.Show vbModal
            
        Case 709 ' Abono remesa
        Case 710 ' Devoluciones
        Case 711 ' Eliminar riesgo
        
        Case 801 ' Cartera de Pagos
            frmTESPagos.Show vbModal
        Case 802 ' Informe Pagos pendientes
            frmTESPagosPdtesList.Show vbModal
        Case 803 ' Informe Pagos bancos
            frmTESPagosBancoList.Show vbModal
        Case 804 ' Realizar Pago
            frmTESRealizarPagos.Show vbModal
        Case 805 ' Transferencias
            frmTESTransferencias.TipoTrans2 = 1 ' de pagos
            frmTESTransferencias.Show vbModal
        Case 806 ' Pagos domiciliados
            frmTESTransferencias.TipoTrans2 = 2 ' pagos domiciliados
            frmTESTransferencias.Show vbModal
        
        Case 807 ' Gastos Fijos
            frmTESGastosFijos.Show vbModal
        
        Case 808 ' Memoria Pagos proveedores
        
        Case 809 ' Compensar proveedor
            CadenaDesdeOtroForm = ""
            frmTESCompensaAboPro.Show vbModal
        
        Case 810 ' Confirming
            frmTESTransferencias.TipoTrans2 = 3 ' confirming
            frmTESTransferencias.Show vbModal
        
        Case 901 ' Informe por NIF
            frmTESInfSituacionNIF.Show vbModal
            
        Case 902 ' Informe por cuenta
            frmTESInfSituacionCta.Show vbModal
        
        Case 903 ' Situación Tesoreria
            frmTESInfSituacion.Show vbModal
        
        
        ' Analitica
        Case 1001 ' Centros de Coste
            frmCCCentroCoste.Show vbModal
            
        Case 1002 ' Consulta de Saldos
            frmCCConExtr.Show vbModal
        
        Case 1003 ' Cuenta de Explotación
            frmCCCtaExplo.Show vbModal
        Case 1004 ' Centros de coste por cuenta
            AbrirListado 17, False
        Case 1005 ' Detalle de explotación
            frmCCDetalleExplota.Show vbModal
            
        ' Presupuestaria
        Case 1101 ' Presupuestos
            Screen.MousePointer = vbHourglass
            'frmColPresu.Show vbModal
            frmPresu.Show vbModal
        Case 1102 ' Listado de Presupuestos
'            AbrirListado 9, False
        Case 1103 ' Balance Presupuestario
            frmPresuBal.Show vbModal
            
            
            
            
            
            
        ' Consolidado
        Case 1201 ' Sumas y Saldos
            AbrirListado 24, False
        Case 1202 ' Balance de Situación
            AbrirListado 51, False
        Case 1203 ' Pérdidas y Ganancias
            AbrirListado 50, False
        Case 1204 ' Cuenta de Explotación
            AbrirListado 31, False
        Case 1205 ' Listado Facturas Clientes
            AbrirListado 53, False
        Case 1206 ' Listado Facturas Proveedores
            AbrirListado 52, False
        
        ' Cierre de Ejercicio
        Case 1301 ' Renumeración de asientos
            frmCierre.Opcion = 0
            frmCierre.Show vbModal
        Case 1302 ' Simulación de cierre
            frmCierre.Opcion = 4
            frmCierre.Show vbModal
        Case 1303 ' Cierre de Ejercicio
            frmCierre.Opcion = 1
            frmCierre.Show vbModal
        Case 1304 ' Deshacer cierre
            frmCierre.Opcion = 5
            frmCierre.Show vbModal
        Case 1305 ' Diario Oficial
'            AbrirListado 14, False
        Case 1306 ' Diario Oficial Resumen
'            AbrirListado 18, False
            frmInfDiarioOficial.Show vbModal
        Case 1307 ' Presentación cuentas anuales
            Telematica 0
        Case 1308 ' Presentación Telemática de Libros
            Telematica 1
        Case 1309 ' memoria de Plazos de Pago
            frmTESMemoriaPlazos.Show vbModal
        
        ' Utilidades
        Case 1401 ' Comprobar cuadre
            Screen.MousePointer = vbHourglass
            frmMensajes.Opcion = 2
            frmMensajes.Show vbModal
        Case 1403 ' Revisar caracteres especiales
'            Screen.MousePointer = vbHourglass
'            frmMensajes.opcion = 14
'            frmMensajes.Show vbModal
        
        Case 1404 ' Agrupacion cuentas
        Case 1405 'Buscar ...
        
        Case 1407 'Desbloquear asientos
            mnHerrAriadnaCC_Click (0)
        Case 1408 'Mover cuentas
            mnHerrAriadnaCC_Click (1)
        Case 1409 'Renumerar registros proveedor
            mnHerrAriadnaCC_Click (5)
        Case 1410 'Aumentar dígitos contables
            mnHerrAriadnaCC_Click (3)
        Case 1411 'cambio de iva
            mnHerrAriadnaCC_Click (4)
        Case 1412 'log de acciones
            Screen.MousePointer = vbHourglass
            Load frmLog
            DoEvents
            frmLog.Show vbModal
            Screen.MousePointer = vbDefault
        Case 1413
            frmImportarUtil.Show vbModal
        Case 1414
            frmImportarNavarres.Show vbModal
            
            
        Case 1415
             frmMensajes.Opcion = 62
             frmMensajes.Show vbModal
                
            
            
        Case Else
  
    End Select
     
     
     
   If Timer - UltimaLecturaReminders > 300 Then
        frmReminders.OnReminders xtpCalendarRemindersFire, Nothing
        If frmReminders.CuantosAvisos > 0 Then frmReminders.Show vbModal, Me
        CerrarAvisos
        UltimaLecturaReminders = Timer
    End If
     
End Sub

Private Sub CerrarAvisos()
    On Error Resume Next
    Unload frmReminders
    Err.Clear
End Sub


Private Sub mnHerrAriadnaCC_Click(Index As Integer)
 
        If vUsu.Nivel > 1 Then
            MsgBox "No tiene permisos", vbExclamation
            Exit Sub
        End If
        'El index 3 , que es la barra, en frmCC es la opcion de NUEVA EMPRESA
        ' y no se llma desde aqui, con lo cual no hay problemo
        'Para el restro cojo el valor del helpidi
        
        frmCentroControl.Opcion = Index
        frmCentroControl.Show vbModal
    
End Sub






Private Sub AbrirMensajeBoxCodejock(QueMsg As Byte, OtrosDatos As String)

    
    Select Case QueMsg
    Case 0 To 10
        'Mensajes standard de la aplicacion
        
        
        
    Case 11
        
        Msg = "Importe descuadre: " & OtrosDatos
        'MuestraMsgCodejock2 "Ariconta6", "Existen asientos descuadrados", Msg, "Revise asientos", "", 0, False
         
        MuestraMsgAriadna "Ariconta6", "Existen asientos descuadrados", Msg, "Revise asientos", "", 0, False
    Case 12
        
        Msg = "Limite: " & UltimaFechaCorrectaSII(vParam.SIIDiasAviso, Now)
        
        'MuestraMsgCodejock2 "Ariadna software", "A.E.A.T.", Msg, "", "Ver facturas|Continuar|", 0, False
        MuestraMsgAriadna "Ariadna software", "A.E.A.T.", "Tiene facturas pendientes de comunicar al SII." & vbCrLf, Msg, "Continuar|Ver facturas|", 64, False
        
   '     Msg = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam ut elit sit amet quam tristique pretium ultricies ut nisl. Donec ultricies ante sodales bibendum dapibus. Suspendisse tincidunt tellus vel ante blandit, at lacinia enim tincidunt. Vivamus lectus libero, gravida eget augue a, lacinia finibus mi. Ut varius orci vehicula ipsum placerat cursus a quis nisi. Vivamus ut placerat lacus. Vivamus at euismod turpis, tincidunt commodo velit. Quisque a elementum erat. Nunc a malesuada urna, nec pretium ipsum. Nulla egestas metus vel lacus lobortis ullamcorper. Integer mollis tortor at velit pharetra aliquet sit amet et augue. Donec gravida imperdiet dui, a pretium leo pretium nec. In facilisis nunc arcu, non volutpat nibh ultricies in. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. In et accumsan ligula."

   ' Msg = Msg & " Proin gravida posuere convallis. Nunc eu diam in massa efficitur tristique vel porttitor metus. Nunc interdum urna metus. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Curabitur in tempus ex. Etiam sagittis placerat neque, non iaculis lorem faucibus et. Sed posuere purus in malesuada condimentum. Morbi nec commodo odio. Vivamus a vehicula ante, eget pulvinar quam. Praesent tristique purus mi, quis feugiat velit lobortis id. Suspendisse potenti. Donec volutpat cursus imperdiet. Curabitur ornare porta sem. Nunc fringilla dolor orci, nec rhoncus sapien tincidunt vitae."

   ' Msg = Msg & " Proin auctor quis massa non ornare. Sed ut aliquam nulla. Donec ornare consequat neque in dapibus. Aliquam volutpat aliquet lectus vel scelerisque. Donec fermentum iaculis tempor. Aliquam erat volutpat. Nulla ut magna urna. Nam semper non leo sit amet eleifend. Praesent sed dictum quam. Aenean ullamcorper elit neque. Vestibulum vehicula nulla sit amet scelerisque varius. Nam blandit turpis sed dolor finibus vehicula et vitae leo. Vivamus egestas elit a iaculis facilisis. Sed mollis at velit ac finibus. In fermentum ipsum ac massa eleifend, ut suscipit augue tincidunt."

   ' Msg = Msg & " Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Pellentesque vel auctor sapien, ut bibendum massa. Phasellus tincidunt metus risus, eget accumsan arcu viverra ut. Nunc rhoncus augue at laoreet rutrum. Phasellus et tempus odio. In eleifend placerat justo, et posuere lorem. Interdum et malesuada fames ac ante ipsum primis in faucibus. Nullam vel erat in odio placerat tincidunt. Etiam finibus purus turpis, non volutpat tellus finibus consequat. Phasellus ultrices magna congue metus dignissim hendrerit. Suspendisse nec massa nisl. Quisque sit amet nunc quis ligula tempus vulputate vel eget nunc. Nunc tincidunt arcu est, nec pellentesque velit dapibus at. Donec metus ex, tempus non lorem in, facilisis congue metus. Etiam porttitor rutrum tortor. Maecenas sollicitudin lacinia ornare."
      
       ' MuestraMsgAriadna "Ariadna software", "", Msg, "", "Ver facturas|Continuar|", 0, False
        
        
       ' MsgBoxA Msg, vbQuestion
    End Select


    
End Sub



Private Sub AccionesIncioAbrirProgramaEmpresa()
Dim C As String
Dim Tiene_A_cancelar As Byte    '0: NO    1:  Cobros      2 : Pagos     3 Los dos
    
    
    If vUsu.Nivel = 0 And vUsu.Id >= 0 And vEmpresa.TieneContabilidad Then
        'EmpresasQueYaHaComunicadoAsientosDescuadrados :  Para que solo lo haga una vez
        If InStr(1, EmpresasQueYaHaComunicadoAsientosDescuadrados, "|" & vEmpresa.codempre & "|") = 0 Then
            If EmpresasQueYaHaComunicadoAsientosDescuadrados = "" Then EmpresasQueYaHaComunicadoAsientosDescuadrados = "|"
            EmpresasQueYaHaComunicadoAsientosDescuadrados = EmpresasQueYaHaComunicadoAsientosDescuadrados & vEmpresa.codempre & "|"
            C = "sum(coalesce(timported,0))-sum(coalesce(timporteh,0))"
            C = DevuelveDesdeBD(C, "hlinapu", "numdiari>0 and numasien>=0 AND fechaent>=" & DBSet(vParam.FechaIni, "F") & " AND 1", "1")
            If C <> "" Then
                If CCur(C) <> 0 Then
                    
                    AbrirMensajeBoxCodejock 11, Format(CCur(C), FormatoImporte)
                    'If RespuestaMsgBox = xtpTaskButtonCancel Then St op
                
                    
                End If
            End If
        
        
            If vParam.PathFicherosInteg <> "" Then HacerDir vParam.PathFicherosInteg
            
        End If
        
        If vEmpresa.TieneTesoreria Then
            If vParamT.ComprobarAlInicio Then
                Tiene_A_cancelar = 0
                If vParamT.TalonesCtaPuente Or vParamT.TalonesCtaPuente Then
                    'Veremos si tiene remesas talon/pagare
                    C = DevuelveDesdeBD("count(*)", "remesas", "tiporem >1 and situacion>='Q' and situacion <'Z' AND 1", "1")
                    If C = "" Then C = "0"
                    If Val(C) > 0 Then Tiene_A_cancelar = 1
                End If
                
                
                'Si lleva confirming con cuenta puente
                If vParamT.ConfirmingCtaPuente Then
                    'Cta puente general
                    C = "transferencias.codmacta=bancos.codmacta  and transferencias.tipotrans = 0 and "
                    C = C & " transferencias.subtipo = 2 and situacion>='Q' and situacion <'Z' AND 1"
                    C = DevuelveDesdeBD("count(*)", "transferencias, bancos", C, "1")
                    
                Else
                
                    
 



                
                
                    C = " transferencias.tipotrans = 0 and transferencias.subtipo = 2 and situacion>='Q' and situacion <'Z'"
                    C = C & " AND (codigo,anyo) IN (select distinct nrodocum,anyodocum "
                    C = C & " from pagos, bancos  where bancos.codmacta =pagos.ctabanc1 and ctaconfirming<>'' and"
                    C = C & " DATE_ADD(now() , INTERVAL coalesce(confirmingriesgo,0) DAY) >fecefect"
                    C = C & " AND situacion=0 and situdocum>='Q' and situdocum <'Z') AND 1"
                    C = DevuelveDesdeBD("count(*)", "transferencias", C, "1")
                    
                End If
                If Val(C) > 0 Then Tiene_A_cancelar = Tiene_A_cancelar + 2

                
                
                If Tiene_A_cancelar > 0 Then
                    
                   ' If HayQueMostrarEliminarRiesgoTalPag Then
                        Screen.MousePointer = vbHourglass
                        frmMensajes.Tipo = CStr(Tiene_A_cancelar)
                        frmMensajes.Opcion = 63
                        frmMensajes.Show vbModal
                   ' End If
                End If
                
            End If
        
        
            'Aseguradas.
            ComprobarOperacionesAseguradas True
        
        End If
        
        
        
    End If
    
    
    
    
    
    
    
    If vParam.SIITiene Then
        DoEvents
        Screen.MousePointer = vbHourglass
        espera 0.1
    
        AbrirFormSII False
    End If
End Sub



Private Sub HacerDir(CadenaPath As String)
Dim Si As Boolean
    On Error Resume Next
    Si = False
    If Right(CadenaPath, 1) = "\" Then
        If Dir(CadenaPath & "*.*", vbArchive) <> "" Then Si = True
    Else
        If Dir(CadenaPath & "\*.*", vbArchive) <> "" Then Si = True
    End If
    If Err.Number <> 0 Then
        Err.Clear
    Else
        If Si Then MsgBox "Archivos pendientes de integrar", vbInformation
    End If
End Sub



Private Sub AbrirFormSII(AbrirSeguro As Boolean)
Dim B As Byte
Dim C As String

    If Not vParam.SIITiene Then Exit Sub

    
    
    If AbrirSeguro Then
        B = 1
    Else
        If vUsu.Nivel > 0 Then Exit Sub
        C = statusBar.Pane(1).Text
        Screen.MousePointer = vbHourglass
        statusBar.Pane(1).Text = "leyendo SII....    "
        
        
        B = DarAvisoPendientesSII()
        If B > 0 Then
            
            'MostrarMensaje 9, "A.E.A.T.", "Facturas pendientes de comunicar al SII", False
            AbrirMensajeBoxCodejock 12, ""
            'If MsgBox("AGENCIA TRIBUTARIA" & vbCrLf & vbCrLf & "Tiene facturas pendientes de comunicar al SII." & vbCrLf & vbCrLf & "¿Verlas ahora?", vbCritical + vbYesNoCancel + vbDefaultButton2) <> vbYes Then B = 0
            If RespuestaMsgBox = 30001 Then B = 0 'NO quiere verlas
            
        End If
    
        statusBar.Pane(1).Text = C
        Screen.MousePointer = vbDefault
    End If


    If B > 0 Then
        frmSII_Avisos.QueMostrarDeSalida = B
        frmSII_Avisos.Show vbModal
    End If

End Sub



'Esto lo tiene Moni "asin", ni digo ni pregunto
Private Sub AbrirListado(numero As Byte, Cerrado As Boolean)
'    Screen.MousePointer = vbHourglass
'    frmListado.EjerciciosCerrados = Cerrado
'    frmListado.Opcion = numero
'    frmListado.Show vbModal
End Sub

Private Sub Telematica(Caso As Integer)
        Me.Enabled = False
        frmTelematica.Opcion = Caso
        frmTelematica.Show vbModal
End Sub




'Establecer y fijar Skin
Public Sub EstablecerSkin(QueSkin As Integer)

    FijaSkin QueSkin

  ' Cargando el archivo del Skin
  ' ============================
    'frmPpal.SkinFramework1.LoadSkin Skn$, ""
    Me.SkinFramework1.ApplyWindow frmppal.hWnd
    Me.SkinFramework1.ApplyOptions = Me.SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
    


    
End Sub

Private Function FijaSkin(numero)
    Me.SkinFramework1.ExcludeModule "crviewer9.dll"

  Select Case (numero)
 
           
            Case 1:
                Skn$ = CStr(App.Path & "\Styles\Office2010.cjstyles")
                Me.SkinFramework1.LoadSkin Skn$, "NormalBlue.ini"
            Case 2:
                Skn$ = CStr(App.Path & "\Styles\Office2010.cjstyles")
                Me.SkinFramework1.LoadSkin Skn$, "NormalSilver.ini"
            Case 3:
                Skn$ = CStr(App.Path & "\Styles\Office2010.cjstyles")
                Me.SkinFramework1.LoadSkin Skn$, "NormalBlack.ini"
                
                  
                
        
        
  End Select
    
End Function



Private Sub PonerCaption()
        Caption = "AriCONTA 6    V-" & App.Major & "." & App.Minor & "." & App.Revision & "    usuario: " & vUsu.Nombre & "      Ejercicio: " & vParam.FechaIni & " - " & vParam.FechaFin
        'Label33.Caption = "   " & vEmpresa.nomempre
End Sub


Public Sub OpcionesMenuInformacion(Id As Long)
    
    Select Case Id
    Case ID_Licencia_Usuario_Final_txt
        LanzaVisorMimeDocumento Me.hWnd, "c:\programas\Ariadna.rtf"
    Case ID_Licencia_Usuario_Final_web
        LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & "Licenciadeuso.html"
    Case ID_Ver_Version_operativa_web
        LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & "Ariconta-6.html"  ' "http://www.ariadnasw.com/clientes/"
    Case ID_Ver_CambiosVersion
        LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & "Versiones.html"
    End Select
    
End Sub

Private Sub ComprobarOperacionesAseguradas(AlInicio As Boolean)

        If vParamT.TieneOperacionesAseguradas Then
            If vUsu.Nivel = 0 Then
                NumRegElim = 0
                'Avisos falta
                If Asegurados_HayAvisos(True) Then NumRegElim = 1
                'Siniestros
                If Asegurados_HayAvisos(False) Then NumRegElim = NumRegElim + 2
                    
                If NumRegElim > 0 Then
                    If NumRegElim > 2 Then
                        frmASeguradoAvisos.Opcion = 0
                    Else
                        'Solo uno de los dos
                        frmASeguradoAvisos.Opcion = CByte(NumRegElim)
                    End If
                    frmASeguradoAvisos.Show vbModal
                
                Else
                    'Si se lanza desde el menu (AUN NO ESTA)
                    If Not AlInicio Then MsgBox "Ningun valor devuelto", vbInformation
                End If
            End If
        End If
End Sub



