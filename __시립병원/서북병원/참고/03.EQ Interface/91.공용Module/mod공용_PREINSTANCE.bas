Attribute VB_Name = "mod°ø¿ë_PREINSTANCE"
Option Explicit

'// API Needed
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'// Instance Code
Private Const InstanceCode = "A8F500EA:D54F:210F:ED0A:F4A5A20C038B"

Public Function PrevInstance() As Boolean
    '*** This function checks if there is another instance running ***
    
    '// Check for a window containing the InstanceCode
    '// If it is found, then return true (another instance is running)
    If FindWindow(vbNullString, ByVal InstanceCode) Then
        PrevInstance = True
        Exit Function
    End If
    
    '// Else, create the window with the InstanceCode
    CreateWindowEx 0&, "STATIC", InstanceCode, 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, 0&
    PrevInstance = False
End Function

