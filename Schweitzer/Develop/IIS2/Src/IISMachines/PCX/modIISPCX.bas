Attribute VB_Name = "modIISPCX"
'-----------------------------------------------------------------------------'
'   ���ϸ�  : modIISPCX.cls
'   �ۼ���  : ������
'   ��  ��  : PCX Ini File Read (Machine Set Info)
'   �ۼ���  : 2007-05-17
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

