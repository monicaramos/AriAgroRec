Attribute VB_Name = "LibPrinter"
Option Explicit

Public Const HWND_BROADCAST = &HFFFF
Public Const WM_WININICHANGE = &H1A

' constants for DEVMODE structure
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

' constants for DesiredAccess member of PRINTER_DEFAULTS
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

' constant that goes into PRINTER_INFO_5 Attributes member
' to set it as default
Public Const PRINTER_ATTRIBUTE_DEFAULT = 4

' Constant for OSVERSIONINFO.dwPlatformId
Public Const VER_PLATFORM_WIN32_WINDOWS = 1

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type DEVMODE
     dmDeviceName As String * CCHDEVICENAME
     dmSpecVersion As Integer
     dmDriverVersion As Integer
     dmSize As Integer
     dmDriverExtra As Integer
     dmFields As Long
     dmOrientation As Integer
     dmPaperSize As Integer
     dmPaperLength As Integer
     dmPaperWidth As Integer
     dmScale As Integer
     dmCopies As Integer
     dmDefaultSource As Integer
     dmPrintQuality As Integer
     dmColor As Integer
     dmDuplex As Integer
     dmYResolution As Integer
     dmTTOption As Integer
     dmCollate As Integer
     dmFormName As String * CCHFORMNAME
     dmLogPixels As Integer
     dmBitsPerPel As Long
     dmPelsWidth As Long
     dmPelsHeight As Long
     dmDisplayFlags As Long
     dmDisplayFrequency As Long
     dmICMMethod As Long        ' // Windows 95 only
     dmICMIntent As Long        ' // Windows 95 only
     dmMediaType As Long        ' // Windows 95 only
     dmDitherType As Long       ' // Windows 95 only
     dmReserved1 As Long        ' // Windows 95 only
     dmReserved2 As Long        ' // Windows 95 only
End Type

Public Type PRINTER_INFO_5
     pPrinterName As String
     pPortName As String
     Attributes As Long
     DeviceNotSelectedTimeout As Long
     TransmissionRetryTimeout As Long
End Type

Public Type PRINTER_DEFAULTS
     pDatatype As Long
     pDevMode As Long
     DesiredAccess As Long
End Type

Declare Function GetProfileString Lib "kernel32" _
Alias "GetProfileStringA" _
(ByVal lpAppName As String, _
ByVal lpKeyName As String, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long) As Long

Declare Function WriteProfileString Lib "kernel32" _
Alias "WriteProfileStringA" _
(ByVal lpszSection As String, _
ByVal lpszKeyName As String, _
ByVal lpszString As String) As Long

Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" _
(ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lparam As String) As Long

Declare Function GetVersionExA Lib "kernel32" _
(lpVersionInformation As OSVERSIONINFO) As Integer

Public Declare Function OpenPrinter Lib "winspool.drv" _
Alias "OpenPrinterA" _
(ByVal pPrinterName As String, _
phPrinter As Long, _
pDefault As PRINTER_DEFAULTS) As Long

Public Declare Function SetPrinter Lib "winspool.drv" _
Alias "SetPrinterA" _
(ByVal hPrinter As Long, _
ByVal Level As Long, _
pPrinter As Any, _
ByVal Command As Long) As Long

Public Declare Function GetPrinter Lib "winspool.drv" _
Alias "GetPrinterA" _
(ByVal hPrinter As Long, _
ByVal Level As Long, _
pPrinter As Any, _
ByVal cbBuf As Long, _
pcbNeeded As Long) As Long

Public Declare Function lstrcpy Lib "kernel32" _
Alias "lstrcpyA" _
(ByVal lpString1 As String, _
ByVal lpString2 As Any) As Long

Public Declare Function ClosePrinter Lib "winspool.drv" _
(ByVal hPrinter As Long) As Long

    Public Sub SelectPrinter(NewPrinter As String)
        Dim Prt As Printer
        
        For Each Prt In Printers
            If Prt.DeviceName = NewPrinter Then
                Set Printer = Prt
            Exit For
            End If
        Next
    End Sub
            

