Attribute VB_Name = "modIP"
Option Explicit

Const MAX_IP = 5

Type IPINFO
    dwAddr As Long
    dwIndex As Long
    dwMask As Long
    dwBCastAddr As Long
    dwReasmSize  As Long
    unused1 As Integer
    unused2 As Integer
End Type

Type MIB_IPADDRTABLE
    dEntrys As Long
    mIPInfo(MAX_IP) As IPINFO
End Type

Type IP_Array
    mBuffer As MIB_IPADDRTABLE
    BufferLen As Long
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long

Public Function ConvertAddressToString(longAddr As Long) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function

'Public Sub Start()
'    Dim Ret As Long, Tel As Long
'    Dim bBytes() As Byte
'    Dim Listing As MIB_IPADDRTABLE
'
'    Form1.Text1 = ""
'
'    On Error GoTo END1
'    GetIpAddrTable ByVal 0&, Ret, True
'
'    If Ret <= 0 Then Exit Sub
'    ReDim bBytes(0 To Ret - 1) As Byte
'    GetIpAddrTable bBytes(0), Ret, False
'
'    CopyMemory Listing.dEntrys, bBytes(0), 4
'    For Tel = 0 To Listing.dEntrys - 1
'        CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
'        Debug.Print "IP address                   : " & ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr) & vbCrLf
'        Debug.Print "IP Subnetmask            : " & ConvertAddressToString(Listing.mIPInfo(Tel).dwMask) & vbCrLf
'        Debug.Print "BroadCast IP address  : " & ConvertAddressToString(Listing.mIPInfo(Tel).dwBCastAddr) & vbCrLf
'        Debug.Print "**************************************" & vbCrLf
'    Next
'    Exit Sub
'
'END1:
'    MsgBox "ERROR"
'End Sub

