Attribute VB_Name = "RuedaRaton"
Option Explicit

' Declaraciones japonesas.
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal Wparam As Long, ByVal Lparam As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A

Dim LocalHwnd As Long
Dim LocalPrevWndProc As Long
Dim MyGrid As DataGrid

' Función que captura (entre otros) el evento de la rueda del ratón.
Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal Wparam As Long, ByVal Lparam As Long) As Long
Dim Direccion As Long
On Error GoTo EWindowProc

  ' Si la acción es de la rueda del ratón, incrementamos/decrementamos el RowTop de MyGrid.
  If Lmsg = WM_MOUSEWHEEL Then
    Direccion = IIf(Wparam > 0, -1, 1)
    If MyGrid.FirstRow > 1 Or Direccion > 0 Then MyGrid.FirstRow = MyGrid.FirstRow + Direccion
  End If
  WindowProc = CallWindowProc(LocalPrevWndProc, Lwnd, Lmsg, Wparam, Lparam)

  Exit Function

EWindowProc:
  
  ' Se ha llegado al tope del Grid, y salta un error. Lo ignoramos.
  Err.Clear

End Function

' Activa la "escucha" para ThisGrid.
Public Sub WheelHook(ThisGrid As DataGrid)
On Error Resume Next

  Set MyGrid = ThisGrid
  LocalHwnd = ThisGrid.hWnd
  LocalPrevWndProc = SetWindowLong(LocalHwnd, GWL_WNDPROC, AddressOf WindowProc)

End Sub

' Desactiva la "escucha" para myGrid.
Public Sub WheelUnHook()
Dim WorkFlag As Long
On Error Resume Next

  WorkFlag = SetWindowLong(LocalHwnd, GWL_WNDPROC, LocalPrevWndProc)
  Set MyGrid = Nothing

End Sub

