VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISIntOrder.cls
'   작성자  : 이상대
'   내  용  : LPIA-NV7 장비 오더정보 클래스
'   작성일  : 2004-10-15
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo      As String   '바코드번호
Private mMsgType    As String   'Message Type (1: Routine, 4:Stat)
Private mWorkNo     As Long     'Worksheet No

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get MsgType() As String
    MsgType = mMsgType
End Property

Public Property Let MsgType(ByVal vData As String)
    mMsgType = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 문자열 조회
'   인수 :
'       - pAccInfo : 접수내역 클래스
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult       As clsIISResult     '결과내역 클래스
    Dim objQCResult     As clsIISQCResult   'QC결과내역 클래스
    Dim strIntBase      As String       '장비기준 검사명
    Dim strItems        As String       '송신할 검사항목
    Dim strWorkNo       As String       '송신할 WorkNo
    Dim strMsgType      As String       '송신할 Message Type(2:Routin, 3:Stat)
    Dim strOutput       As String       '송신할 데이터
    Dim lngSendCnt      As Long         '송신할 검사항목 개수
    
    '## 송신할 검사항목 조회
    If pAccInfo.QcFg = "0" Then     '## 일반검체
        For Each objResult In pAccInfo.Results
            strIntBase = Format$(Format$(objResult.IntNm.IntBase, "###"), "@@@")
            strItems = strItems & strIntBase
            lngSendCnt = lngSendCnt + 1
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then '## QC검체
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = Format$(Format$(objQCResult.IntNm.IntBase, "###"), "@@@")
            strItems = strItems & strIntBase
            lngSendCnt = lngSendCnt + 1
        Next
        Set objQCResult = Nothing
    End If
    
    '## 최대 16개까지 전송가능하고 남는부분은 공백으로!
    strItems = strItems & Space(3 * (16 - lngSendCnt))
    
    '## Worksheet No, Message Type
    strMsgType = IIf(mMsgType = "1", "2", "3")
    If strMsgType = "3" Then
        strWorkNo = Space(4)
    Else
        mWorkNo = mWorkNo + 1
        strWorkNo = Format$(mWorkNo, "0000")
    End If
    
    strOutput = STX & strMsgType & strWorkNo & Format$(mBarNo, "!" & String$(13, "@")) & _
                strItems & ETX

    GetOrder = strOutput
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수정보가 없을때의 문자열 조회
'   반환 : 접수정보가 없을때의 문자열
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    Dim strWorkNo   As String   '송신할 Worksheet No
    Dim strMsgType  As String   '송신할 Message Type(2:Routin, 3:Stat)
    Dim strItems    As String   '송신할 검사항목
    Dim strOutput   As String   '송신할 데이터
    
    '## Worksheet No, Message Type
    strItems = Space(3 * 16)
    strMsgType = IIf(mMsgType = "1", "2", "3")
    If strMsgType = "3" Then
        strWorkNo = Space(4)
    Else
        mWorkNo = mWorkNo + 1
        strWorkNo = Format$(mWorkNo, "0000")
    End If
    
    strOutput = STX & strMsgType & strWorkNo & Format$(mBarNo, "!" & String$(13, "@")) & _
                strItems & ETX
    
    GetNoOrder = strOutput
End Function

'-----------------------------------------------------------------------------'
'   기능 : Control Information 문자열 조회
'-----------------------------------------------------------------------------'
Public Function GetControlInfo() As String
    Dim strOutput As String
    
    strOutput = STX & "900" & ETX
    GetControlInfo = strOutput
End Function

'-----------------------------------------------------------------------------'
'   기능 : 클래스 멤버변수 초기화
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mMsgType = ""
End Sub

