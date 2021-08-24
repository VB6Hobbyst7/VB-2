Attribute VB_Name = "Calendar"
Option Explicit

Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Declare Function SetCursorPos Lib "User32" (ByVal x As Long, ByVal y As Long) As Long

Type PointAPI
     x As Long
     y As Long
End Type
Global ReturnPos    As PointAPI


