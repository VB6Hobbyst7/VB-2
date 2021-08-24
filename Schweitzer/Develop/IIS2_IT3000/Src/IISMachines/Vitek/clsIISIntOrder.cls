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
'   내  용  : Vitek 장비 오더정보 클래스
'   작성일  : 2004-09-21
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mCheckSum    As String       'Check Sum
Private mSeq         As Long         '오더를 전송한 Spread의 Row
Private mSendCnt     As Long         '전송할 검체개수

Public Property Get CheckSum() As String
    CheckSum = mCheckSum
End Property

Public Property Let CheckSum(ByVal vData As String)
    mCheckSum = vData
End Property

Public Property Get Seq() As Long
    Seq = mSeq
End Property

Public Property Let Seq(ByVal vData As Long)
    mSeq = vData
End Property

Public Property Get SendCnt() As Long
    SendCnt = mSendCnt
End Property

Public Property Let SendCnt(ByVal vData As Long)
    mSendCnt = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 오더문자열 조회
'   인수 :
'       - pAccInfo  : 접수정보 클래스
'       - pVitekNo  : Vitek No
'   반환 : 오더문자열
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo, ByVal pVitekNo As String) As String
    Dim strSexAge As String  '송신할 Sex/Age
    Dim strColTm  As String  '송신할 채혈시간 (HH:MM형식)
    Dim strRcvTm  As String  '송신할 접수시간 (HH:MM형식)
    Dim strOrder  As String  '오더정보 문자열
    
    '## px(진료과), ss(검체), ct(검사유형), pl(병동) 필드는 Vitek에 미리 정의되어야 함!!
    '   ct필드는 항상 cs가 Culture & Sensitivity로 등록하게 해야함!
    With pAccInfo
        strSexAge = .Sex & "/" & Mid$(.Age, 1, Len(.Age) - 1)
        strColTm = Format$(Format$(.ColTm, "##:##:##"), "HH:MM")
        strRcvTm = Format$(Format$(.ColTm, "##:##:##"), "HH:MM")
        
        '## 운영시 사용할것!
        strOrder = "mtmpr|pi" & .PtId & "|w1" & strSexAge & "|px" & .DeptCd & "|pl" & .WardId & _
                   "|si|ss" & .SpcCd & "|s1" & .ColDt & "|s2" & strColTm & "|s3" & .RcvDt & _
                   "|s4" & strRcvTm & "|ci" & pVitekNo & "|ctcs|zz|" & vbCrLf
    End With
    
    mCheckSum = GetChkSum(strOrder & "")
    GetOrder = strOrder
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열의 CheckSum을 구함
'   인수 :
'       - pMsg : 문자열
'   반환 : CheckSum
'-----------------------------------------------------------------------------'
Public Function GetChkSum(ByVal pMsg As String) As String
    Dim lngChkSum   As Long
    Dim i           As Long

    For i = 1 To Len(pMsg)
        lngChkSum = (lngChkSum + Asc(Mid(pMsg, i, 1))) Mod 256
    Next

    If lngChkSum = 0 Then
        GetChkSum = "00"
    Else
        GetChkSum = LCase(Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2))
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : 클래스 멤버변수 초기화
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mCheckSum = ""
    mSeq = 0
End Sub
