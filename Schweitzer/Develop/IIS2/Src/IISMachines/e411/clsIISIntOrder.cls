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
'   내  용  : CobasE411 장비 오더정보 클래스
'   작성일  : 2014-01-21
'-----------------------------------------------------------------------------'

Option Explicit

Private mAccInfo    As clsIISAccInfo   '접수내역 클래스

Private mBarNo      As String   '바코드번호
Private mSeq        As String   'Seq
Private mRackNo     As String   'Rack No
Private mTubePos    As String   'Tube Position
Private mNoOrder    As String   '접수정보 존재유무 (True:비존재, False:존재)
Private mIsSending  As Boolean  '오더 전송후 남은 문자열의 존재유무
Private mOrder      As String   '오더 전송후 남은오더

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

Public Property Get Seq() As String
    Seq = mSeq
End Property

Public Property Let Seq(ByVal vData As String)
    mSeq = vData
End Property

Public Property Get RackNo() As String
    RackNo = mRackNo
End Property

Public Property Let RackNo(ByVal vData As String)
    mRackNo = vData
End Property

Public Property Get TubePos() As String
    TubePos = mTubePos
End Property

Public Property Let TubePos(ByVal vData As String)
    mTubePos = vData
End Property

Public Property Get NoOrder() As Boolean
    NoOrder = mNoOrder
End Property

Public Property Let NoOrder(ByVal vData As Boolean)
    mNoOrder = vData
End Property

Public Property Get IsSending() As Boolean
    IsSending = mIsSending
End Property

Public Property Let IsSending(ByVal vData As Boolean)
    mIsSending = vData
End Property

Public Property Get Order() As String
    Order = mOrder
End Property

Public Property Let Order(ByVal vData As String)
    mOrder = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 문자열을 조회
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetOrder() As String
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim strIntBase  As String           '장비기준 검사명
    Dim strItems    As String           '송신할 검사항목
    Dim strTemp     As String

    '## 1.0.1: 이상대(2004-12-07)
    '   - 같은 검사항목에 중복전송문제 수정
    '## 오더문자열 형식: ^^^1\^^^2
    If mAccInfo.QcFg = "0" Then
        For Each objResult In mAccInfo.Results
            strIntBase = objResult.IntNm.IntBase
            strIntBase = Mid$(strIntBase, 1, Len(strIntBase) - 1)
            
            If strIntBase <> strTemp Then
                If strItems = "" Then
                    strItems = "^^^" & strIntBase
                Else
                    strItems = strItems & "\^^^" & strIntBase
                End If
                strTemp = strIntBase
            End If
        Next
        Set objResult = Nothing
    ElseIf mAccInfo.QcFg = "1" Then
        For Each objQCResult In mAccInfo.QCResults
            strIntBase = objQCResult.IntNm.IntBase
            strIntBase = Mid$(strIntBase, 1, Len(strIntBase) - 1)
            
            If strIntBase <> strTemp Then
                If strItems = "" Then
                    strItems = "^^^" & strIntBase
                Else
                    strItems = strItems & "\^^^" & strIntBase
                End If
                strTemp = strIntBase
            End If
        Next
        Set objQCResult = Nothing
    End If

    GetOrder = strItems
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
    mBarNo = ""
    mSeq = ""
    mRackNo = ""
    mTubePos = ""
    mNoOrder = False
    mIsSending = False
    mOrder = ""
    Set mAccInfo = Nothing
End Sub

Private Sub Class_Terminate()
    Set mAccInfo = Nothing
End Sub
