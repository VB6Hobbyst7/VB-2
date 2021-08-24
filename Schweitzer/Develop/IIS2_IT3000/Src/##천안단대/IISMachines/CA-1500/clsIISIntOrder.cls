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
'   내  용  : CA-1500 오더전송 클래스
'   작성일  : 2005-07-25
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo      As String           '바코드번호
Private mRackNo     As String           'Rack No
Private mPos        As String           'Tube Position
Private mNoOrder    As Boolean          '접수정보 존재유무(True:비존재, False:존재)
Private mOrder      As String           'Order Information Text Format
Private mSendCnt    As String           '오더 전송회수

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get RackNo() As String
    RackNo = mRackNo
End Property

Public Property Let RackNo(ByVal vData As String)
    mRackNo = vData
End Property

Public Property Get Pos() As String
    Pos = mPos
End Property

Public Property Let Pos(ByVal vData As String)
    mPos = vData
End Property

Public Property Get NoOrder() As Boolean
    NoOrder = mNoOrder
End Property

Public Property Let NoOrder(ByVal vData As Boolean)
    mNoOrder = vData
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
    Dim strStatFg     As String           '응급유무 (R:Routin, E:Stat)
    Dim strTemp       As String

    '## 검사항목 문자열 생성
    If pAccInfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In pAccInfo.Results
            strIntBase = objResult.IntNm.IntBase
            strIntBase = Mid$(strIntBase, 1, Len(strIntBase) - 1) & "0" & Space(6)

            If strIntBase <> strTemp Then
                strItems = strItems & strIntBase
                strTemp = strIntBase
            End If
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = objQCResult.IntNm.IntBase
            strIntBase = Mid$(strIntBase, 1, Len(strIntBase) - 1) & "0" & Space(6)

            If strIntBase <> strTemp Then
                strItems = strItems & strIntBase
                strTemp = strIntBase
            End If
        Next
        Set objQCResult = Nothing
    End If

    strStatFg = IIf(pAccInfo.StatFg = "1", "E", "U")

    mOrder = STX & "S2210101" & strStatFg & Space(6) & Space(4) & mRackNo & mPos & mBarNo & _
             "B" & Space(15) & strItems & ETX
    GetOrder = mOrder
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수정보가 없을때 오더 문자열 조회
'   반환 : 접수정보가 없을때 오더 문자열
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    mOrder = STX & "S2210101U" & Space(6) & Space(4) & mRackNo & mPos & mBarNo & "B" & _
             Space(15) & ETX
    
    GetNoOrder = mOrder
End Function

'-----------------------------------------------------------------------------'
'   기능 : 클래스 멤버변수 초기화
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mRackNo = ""
    mPos = ""
    mNoOrder = True
    mOrder = ""
    mSendCnt = 0
End Sub
