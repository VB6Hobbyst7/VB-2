Attribute VB_Name = "Net_Func"
Option Explicit

Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1
Public Const WSADescription_Len = 256
Public Const WSASYS_Status_Len = 128

Public Type HOSTENT
hName As Long
hAliases As Long
hAddrType As Integer
hLength As Integer
hAddrList As Long
End Type

Public Type WSADATA
wversion As Integer
wHighVersion As Integer
szDescription(0 To WSADescription_Len) As Byte
szSystemStatus(0 To WSASYS_Status_Len) As Byte
iMaxSockets As Integer
iMaxUdpDg As Integer
lpszVendorInfo As Long
End Type

Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long

Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal _
wVersionRequired&, lpWSAData As WSADATA) As Long

Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, _
ByVal HostLen As Long) As Long

Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal _
hostname$) As Long

Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal _
hpvSource&, ByVal cbCopy&)


Function hibyte(ByVal wParam As Integer)
hibyte = wParam \ &H100 And &HFF&
End Function

Function lobyte(ByVal wParam As Integer)
lobyte = wParam And &HFF&
End Function

Sub SocketsInitialize()
    Dim WSAD As WSADATA
    Dim iReturn As Integer
    Dim sLowByte As String, sHighByte As String, sMsg As String
    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)
    
    If iReturn <> 0 Then
    MsgBox "Winsock.dll is not responding."
    End
    End If
    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
    WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then
    sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
    sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
    sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
    sMsg = sMsg & " is not supported by winsock.dll "
    MsgBox sMsg
    End
    End If
    
    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
    sMsg = "This application requires a minimum of "
    sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
    MsgBox sMsg
    End
    End If
End Sub

Sub SocketsCleanup()
    Dim lReturn As Long
    lReturn = WSACleanup()
    
    
    If lReturn <> 0 Then
    MsgBox "Socket Error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
    
    End
    End If
End Sub

Public Function GetTheIP()
Dim hostname As String * 256
Dim hostent_addr As Long
Dim host As HOSTENT
Dim hostip_addr As Long
Dim temp_ip_address() As Byte
Dim i As Integer
Dim ip_address As String


    If gethostname(hostname, 256) = SOCKET_ERROR Then
    MsgBox "Windows Sockets Error " & Str(WSAGetLastError())
    Exit Function
    Else
    hostname = Trim$(hostname)
    End If
    hostent_addr = gethostbyname(hostname)
    
    
    If hostent_addr = 0 Then
    'MsgBox "Winsock.dll is not responding."
    Exit Function
    End If
    RtlMoveMemory host, hostent_addr, LenB(host)
    RtlMoveMemory hostip_addr, host.hAddrList, 4
    'MsgBox hostname
    'get all of the IP address if machine is
    ' multi-homed
    
    
    Do
    ReDim temp_ip_address(1 To host.hLength)
    RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength
    
    
    For i = 1 To host.hLength
    ip_address = ip_address & temp_ip_address(i) & "."
    Next
    ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
    'MsgBox ip_address
    
    host.hAddrList = host.hAddrList + LenB(host.hAddrList)
    RtlMoveMemory hostip_addr, host.hAddrList, 4
    Loop While (hostip_addr <> 0)
    
    GetTheIP = ip_address
    ip_address = ""
End Function



