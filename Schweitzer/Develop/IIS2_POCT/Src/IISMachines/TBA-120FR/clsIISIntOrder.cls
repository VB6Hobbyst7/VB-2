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
'   내  용  : TBA-120FR 오더전송 클래스
'   작성일  : 2005-02-04
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo   As String   '바코드번호
Private mDiskNo  As String   'Disk No
Private mPos     As String   'Position
Private mOrder   As String   '오더정보 문자열
Private mSendCnt As Long     '오더 재전송 회수

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get DiskNo() As String
    DiskNo = mDiskNo
End Property

Public Property Let DiskNo(ByVal vData As String)
    mDiskNo = vData
End Property

Public Property Get Pos() As String
    Pos = mPos
End Property

Public Property Let Pos(ByVal vData As String)
    mPos = vData
End Property

Public Property Get Order() As String
    Order = mOrder
End Property

Public Property Let Order(ByVal vData As String)
    mOrder = vData
End Property

Public Property Get SendCnt() As Long
    SendCnt = mSendCnt
End Property

Public Property Let SendCnt(ByVal vData As Long)
    mSendCnt = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 문자열 조회
'   인수 :
'       - pAccInfo : 접수내역 클래스
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult     As clsIISResult     '결과내역 클래스
    Dim objQCResult   As clsIISQCResult   'QC결과내역 클래스
    Dim strIntBase    As String           '장비기준 검사명
    Dim strItems      As String           '송신할 검사항목 문자열
    
    '## 검사항목 문자열 생성
    If pAccInfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In pAccInfo.Results
            strIntBase = Format$(Format$(objResult.IntNm.IntBase, "#"), "@@@@") & "1"
            strItems = strItems & strIntBase
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = Format$(Format$(objResult.IntNm.IntBase, "#"), "@@@@") & "1"
            strItems = strItems & strIntBase
        Next
        Set objQCResult = Nothing
    End If
    
    mOrder = STX & "O " & mBarNo & mDiskNo & mPos & "  1" & strItems & ETB & String$(12, "0") & _
             ETB & ETX
    GetOrder = mOrder
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수정보가 없을때 오더 문자열 조회
'   반환 : 접수정보가 없을때 오더 문자열
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    mOrder = STX & "O " & mBarNo & mDiskNo & mPos & ETB & ETB & ETX
    GetNoOrder = mOrder
End Function

'-----------------------------------------------------------------------------'
'   기능 : 클래스 멤버변수 초기화
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mDiskNo = ""
    mPos = ""
    mOrder = ""
    mSendCnt = 0
End Sub

