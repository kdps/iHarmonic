Attribute VB_Name = "modSystem"
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Const LWA_ALPHA As Long = &H2
Private Const WS_EX_LAYERED As Long = &H80000
Private Const GWL_EXSTYLE As Long = -20
Private Const SW_SHOW As Long = 5
Private Const RDW_UPDATENOW As Long = &H100

Public Sub Trans(lngHwnd As Long, Optional ByVal Speed As Byte = 1, Optional ByVal OpaquePercent As Byte = 85)
Dim Cnt As Long
On Error Resume Next
SetWindowLong lngHwnd, GWL_EXSTYLE, GetWindowLong(lngHwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
SetLayeredWindowAttributes lngHwnd, 0, 0, LWA_ALPHA
ShowWindow lngHwnd, SW_SHOW
RedrawWindow lngHwnd, ByVal 0&, ByVal 0&, RDW_UPDATENOW
For Cnt = 0 To OpaquePercent Step Speed
SetLayeredWindowAttributes lngHwnd, 0, (Cnt / 100) * 255, LWA_ALPHA
Next Cnt
End Sub
