Attribute VB_Name = "Module1"
Option Explicit


'Public Declare Sub call_wdsl Lib "D:\프로젝트\수정구보건소\SOAP\CALL_WSDL.dll" Alias "GetOrder" (ByVal strOrder as String ) as String

Public Declare Function call_wdsl Lib "D:\프로젝트\수정구보건소\SOAP\CALL_WSDL.dll" Alias "GetOrder" (ByVal strOrder As String) As String

'Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

'Public Declare Function DownLoadOrder_BarCode_Dll Lib "D:\프로젝트\수정구보건소\SOAP\공공보건\보건정보 Sample\D2007_Dll\CALL_WSDL.DLL" Alias "SelectOrder" (ByVal strOrder1 As String)
Declare Function DownLoadOrder_WorkList_Dll Lib "C:\HealthWSDL.dll" Alias "DownLoadOrder_WorkList_DllA" (ByVal strOrder1 As String) As Boolean


'#If Win32 Then
'Declare Function SelectOrder& Lib "D:\프로젝트\수정구보건소\SOAP\공공보건\보건정보 Sample\D2007_Dll\CALL_WSDL.DLL" (ByVal strOrder1 As Boolean, ByVal strOrder2$, ByVal strOrder3$)
'#Else
'Declare Function SelectOrder% Lib "D:\프로젝트\수정구보건소\SOAP\공공보건\보건정보 Sample\D2007_Dll\CALL_WSDL.DLL" (ByVal strOrder1 As Boolean, ByVal strOrder2 As String, ByVal strOrder3 As Variant)
'#End If

'
#If Win32 Then
Declare Function dce_setenv& Lib "odet30.dll" (ByVal s1$, ByVal s2$, ByVal s3$)
#Else
Declare Function dce_setenv% Lib "odet30.dll" (ByVal s1$, ByVal s2$, ByVal s3$)
#End If

'Public Declare Function DownLoadOrder_WorkList_Dll Lib "C:\CALL_WSDL.DLL" (ByVal strOrder1 As String) As String

'New_SelectOrder
