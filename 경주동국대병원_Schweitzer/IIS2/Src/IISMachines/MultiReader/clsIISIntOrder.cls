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
'   작성자  : 오세원
'   내  용  : MultiReader 장비 오더정보 클래스
'   작성일  : 2010-05-31
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mAccInfo    As clsIISAccInfo    '접수내역 클래스
Private mBarNo      As String           '바코드번호
Private mNoOrder    As Boolean          '접수정보 존재유무 (True:비존재, False:존재)
Private mItems()    As String           '전송할 검사항목들
Private mCount      As Long             '전송할 총 검사항목 개수
Private mSendCnt    As Long             '전송한 검사항목 개수
Private mPtId       As String           '전송할 환자ID
Private mSex        As String           '전송할 성별
Private mStatFg     As String           '전송할 응급유무

Public Property Get AccInfo() As clsIISAccInfo
    Set AccInfo = mAccInfo
End Property

Public Property Let AccInfo(ByVal vData As clsIISAccInfo)
    Set mAccInfo = vData
End Property

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get NoOrder() As Boolean
    NoOrder = mNoOrder
End Property

Public Property Let NoOrder(ByVal vData As Boolean)
    mNoOrder = vData
End Property

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Let Count(ByVal vData As Long)
    mCount = vData
End Property

Public Property Get SendCnt() As Long
    SendCnt = mSendCnt
End Property

Public Property Let SendCnt(ByVal vData As Long)
    SendCnt = vData
End Property

Public Property Get PtId() As String
    PtId = mPtId
End Property

Public Property Let PtId(ByVal vData As String)
    mPtId = vData
End Property

Public Property Get Sex() As String
    Sex = mSex
End Property

Public Property Let Sex(ByVal vData As String)
    mSex = vData
End Property

Public Property Get StatFg() As String
    StatFg = mStatFg
End Property

Public Property Let StatFg(ByVal vData As String)
    mStatFg = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 조회
'-----------------------------------------------------------------------------'
Public Sub GetItems()
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim strIntBase  As String           '장비기준 검사명
    Dim strTemp     As String

    ReDim mItems(mAccInfo.SendCnt)
    
    If mAccInfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In mAccInfo.Results
            strIntBase = mGetP(objResult.IntNm.IntBase, 1, "|")
            If strIntBase <> strTemp Then
                mItems(mCount) = "^^^" & strIntBase
                strTemp = strIntBase
                mCount = mCount + 1
            End If
        Next
        Set objResult = Nothing
        
        mPtId = mAccInfo.PtId
        mPtId = mAccInfo.Sex
        mStatFg = IIf(mAccInfo.StatFg = "1", "S", "")
    ElseIf mAccInfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In mAccInfo.QCResults
            strIntBase = mGetP(objQCResult.IntNm.IntBase, 1, "|")
            If strIntBase <> strTemp Then
                mItems(mCount) = "^^^" & strIntBase
                strTemp = strIntBase
                mCount = mCount + 1
            End If
        Next
        Set objQCResult = Nothing
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 검사항목 조회
'   반환 : 검사항목
'-----------------------------------------------------------------------------'
Public Function GetItem() As String
    Dim strIntBase  As String
    
    GetItem = mItems(mSendCnt)
    mSendCnt = mSendCnt + 1
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
        GetChkSum = Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2)
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : 클래스 멤버변수 초기화
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    Set mAccInfo = Nothing
    mBarNo = ""
    mNoOrder = False
    mCount = 0
    mSendCnt = 0
    mPtId = ""
    mSex = ""
    mStatFg = ""
End Sub

Private Sub Class_Terminate()
    Set mAccInfo = Nothing
End Sub

