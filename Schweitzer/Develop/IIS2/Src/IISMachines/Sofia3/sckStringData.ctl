VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl sckStringData 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   660
   ScaleHeight     =   300
   ScaleWidth      =   660
   Begin VB.Timer tmrChkSendMsg 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   2355
   End
   Begin VB.Timer tmrChkRecvMsg 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   870
      Top             =   1665
   End
   Begin VB.Timer tmrChkStr 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   870
      Top             =   1095
   End
   Begin MSWinsockLib.Winsock sckData 
      Left            =   315
      Top             =   1095
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      Height          =   285
      Left            =   15
      Top             =   0
      Width           =   630
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Socket"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   510
   End
End
Attribute VB_Name = "sckStringData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const HEADER_CHAR As String = "[START]"
Private Const FOOTER_CHAR As String = "[END]"

Dim m_BufferString      As String
Dim m_BufferRecvMsg     As New Collection
Dim m_BufferSendMsg     As New Collection

Event ConnectOpen(RemoteIP As String, RemotePort As String, RemoteHost As String)
Event ConnectClose()
Event ConnectRequest(RemoteIP As String, RemotePort As String, RemoteHost As String)

Event SendComplete()
Event ProcRecvMessage(sMessage As String)

Event AddLog(sMsg As String, sPos As String)
Event ErrLog(sNum As String, sMsg As String, sPos As String)


Private Sub InitBuffer()
    m_BufferString = ""
    
    Dim idx As Long
    
    For idx = 1 To m_BufferRecvMsg.Count
        m_BufferRecvMsg.Remove idx
    Next
    
    For idx = 1 To m_BufferSendMsg.Count
        m_BufferSendMsg.Remove idx
    Next
End Sub

Public Function ProcSendMessage(sMessage As String)
    Dim sMsg    As String
    
'    sMsg = HEADER_CHAR & sMessage & FOOTER_CHAR
    m_BufferSendMsg.Add sMessage
End Function

Public Sub SckConnect(sIP As String, nPort As Integer, Optional bForceConnect As Boolean = True)
    If sckData.state <> sckClosed Then
        If bForceConnect = True Then
            sckData.Close
        Else
            Exit Sub
        End If
    End If
    
    sckData.Connect sIP, nPort
End Sub

Public Sub SckClose()
    If sckData.state <> sckClosed Then
        sckData.Close
    End If
End Sub

Private Sub sckData_Close()
    tmrChkStr.Enabled = False
    tmrChkRecvMsg.Enabled = False
    tmrChkSendMsg.Enabled = False

    RaiseEvent ConnectClose
End Sub

Private Sub sckData_Connect()
    tmrChkStr.Enabled = True
    tmrChkSendMsg.Enabled = True
    tmrChkRecvMsg.Enabled = True

    RaiseEvent ConnectOpen(sckData.RemoteHostIP, sckData.RemotePort, sckData.RemoteHost)
End Sub

Public Function state() As String
    state = GetState2String
End Function

Public Function GetState2String() As String
    Dim nState As Byte
    nState = sckData.state
    Select Case nState
        Case sckClosed: GetState2String = "Closed"                          '닫혀있음
        Case sckOpen: GetState2String = "Open"                              '열려있음
        Case sckListening: GetState2String = "Listening"                    '기다리는 중(서버)
        Case sckConnectionPending: GetState2String = "Connection Pending"   '연결 보류 중
        Case sckResolvingHost: GetState2String = "Resolving Host"           '호스트 고정 중
        Case sckHostResolved: GetState2String = "Host Resolved"             '호스트 고정 완료
        Case sckConnecting: GetState2String = "Connecting"                 '연결 중
        Case sckConnected:  GetState2String = "Connected"                   '연결 완료
        Case sckClosing:    GetState2String = "Closing"                     '피어가 연결을 닫고 있음
        Case sckError: GetState2String = "ERROR"                            '오류
    End Select
End Function

Public Property Get StateConnIP() As String
    If sckData.state <> sckConnected Then Exit Property
    StateConnIP = sckData.RemoteHostIP
End Property

Public Property Get StateConnPort() As String
    If sckData.state <> sckConnected Then Exit Property
    StateConnPort = sckData.RemotePort
End Property

Public Function Accept(requestID As Long) As Boolean
    If sckData.state <> sckClosed Then
        sckData.Close
    End If
    
   ' Call InitBuffer
    sckData.Accept requestID
    
    tmrChkStr.Enabled = True
    tmrChkSendMsg.Enabled = True
    tmrChkRecvMsg.Enabled = True
    RaiseEvent ConnectRequest(sckData.RemoteHostIP, sckData.RemotePort, sckData.RemoteHost)
    Accept = True
End Function

Private Sub sckData_DataArrival(ByVal bytesTotal As Long)
    Dim sBuffer         As String
    Static runFlag      As Byte
    Dim strRcvBuffer As String
    Dim strSndBuffer As String
    
    If runFlag = 1 Then Exit Sub
    
    runFlag = 1
    sckData.GetData sBuffer, vbString, bytesTotal
'    m_BufferString = m_BufferString & sBuffer
    m_BufferString = sBuffer
    
    runFlag = 0
    
    Debug.Print sBuffer
    
    Call frmIISSofia3.RcvSocketData(sBuffer)
    
    If state <> "Connected" Then Exit Sub
    
End Sub

Private Sub sckData_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent AddLog("[Socket Error(" & Number & ")] " & Description, UserControl.Name & "-sckData_Error()")
End Sub

Private Function ProcStrPasing(TotStr As String, strFront As String, strEnd As String) As String
 On Error GoTo err
    If InStr(1, TotStr, strFront, vbTextCompare) = 0 Then
        ProcStrPasing = ""
    Else
        ProcStrPasing = Mid(TotStr, InStr(1, TotStr, strFront, vbTextCompare) + Len(strFront), (InStr(1, TotStr, strEnd, vbTextCompare)) - (InStr(1, TotStr, strFront, vbTextCompare) + Len(strFront)))
    End If
    Exit Function
err:
    err.Clear
End Function

Private Sub sckData_SendComplete()
    RaiseEvent SendComplete
End Sub

Private Sub tmrChkRecvMsg_Timer()
'받는 컬렉션에서 메시지가 있으면 메시지를 밖으로 고고싱한다.
    Static runFlag          As Byte
    If m_BufferRecvMsg.Count <= 0 Then Exit Sub
    If runFlag = 1 Then Exit Sub
    
    runFlag = 1
    
    RaiseEvent ProcRecvMessage(m_BufferRecvMsg.Item(1))
    m_BufferRecvMsg.Remove 1
    
    runFlag = 0
End Sub

Public Function RecvMessageCollection() As Collection
    Set RecvMessageCollection = m_BufferRecvMsg
End Function

Public Function SendMessageCollection() As Collection
    Set SendMessageCollection = m_BufferSendMsg
End Function

Private Sub tmrChkSendMsg_Timer()
'보내는 컬렉션에서 메시지가 있으면 메시지를 윈삭으로 보낸다.
    Static runFlag          As Byte
    If m_BufferSendMsg.Count <= 0 Then Exit Sub
    If runFlag = 1 Then Exit Sub
    
    runFlag = 1
    
    sckData.SendData m_BufferSendMsg.Item(1)
    m_BufferSendMsg.Remove 1
    
    runFlag = 0
End Sub

Private Sub tmrChkStr_Timer()
'통스트링에서 메시지단위로 짤라 메시지 컬렉션에 넣는다.
    Static runFlag  As Byte
    
    Dim nFindHeaderPos      As Long         '버퍼에서 젤 처음 HEADER_CHAR 포지션 값
    Dim nFindFooterPos      As Long         '버퍼에서 발견된 HEADER_CHAR로부터 젤 처음 FOOTER_CHAR 포지션 값
    Dim nFindFooterEndPos   As Long         '버퍼에서 메시지를 짜르기 위한 FOOTER_CHAR 포지션 마지막 값
    Dim sTotMessage         As String       '추출된 헤더 포함된 메시지 문자열
    Dim sMessage            As String       '추출된 헤더 뺀 메시지 문자열
    
    If m_BufferString = "" Then Exit Sub
    If runFlag = 1 Then Exit Sub
    
    runFlag = 1
    
    nFindHeaderPos = InStr(1, m_BufferString, HEADER_CHAR)
    
    If nFindHeaderPos <> 0 Then
        nFindFooterPos = InStr(nFindHeaderPos, m_BufferString, FOOTER_CHAR)
        If nFindFooterPos <> 0 Then
            nFindFooterEndPos = nFindFooterPos + Len(FOOTER_CHAR)
            '헤더 포함한 메시지 추출
            sTotMessage = Mid(m_BufferString, 1, nFindFooterEndPos - 1)
            
            '헤더를 뺀 메시지 추출
            sMessage = ProcStrPasing(sTotMessage, HEADER_CHAR, FOOTER_CHAR)
            
            '추출된 메시지를 제외한 메시지 조합
            m_BufferString = Mid(m_BufferString, nFindFooterEndPos)
            
            '컬렉션에 메시지 추가
            m_BufferRecvMsg.Add sMessage
        End If
    End If
    runFlag = 0
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 660
    UserControl.Height = 300
End Sub
