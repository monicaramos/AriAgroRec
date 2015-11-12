Attribute VB_Name = "ModCarpeta"
Option Explicit
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
Public Type BROWSEINFO
     hOwner As Long
     pidlRoot As Long
     pszDisplayName As String
     lpszTitle As String
     ulFlags As Long
     lpfn As Long
     lParam As Long
     iImage As Long
End Type
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
     Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
     ByVal pszPath As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" _
     Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)


Public Function GetFolder(Titulo As String) As String
     Dim bInf As BROWSEINFO
     Dim RetVal As Long
     Dim PathID As Long
     Dim RetPath As String
     Dim Offset As Integer
     'Establece las propiedades del dialogo
     bInf.hOwner = frmPpal.hWnd
     bInf.lpszTitle = Titulo
     bInf.ulFlags = BIF_RETURNONLYFSDIRS
     
     'Muestra el cuadro de dialogo del browse
     PathID = SHBrowseForFolder(bInf)
     RetPath = Space$(512)
     RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
     If RetVal Then
          Offset = InStr(RetPath, Chr$(0))
          GetFolder = Left$(RetPath, Offset - 1)
          CoTaskMemFree PathID
     Else
          GetFolder = ""
     End If
End Function




