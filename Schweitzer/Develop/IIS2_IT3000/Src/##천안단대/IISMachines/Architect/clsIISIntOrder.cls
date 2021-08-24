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
'   내  용  : Architect 장비 오더정보 클래스
'   작성일  : 2004-10-13
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mAccInfo    As clsIISAccInfo    '접수내역 클래스
Private mBarNo      As String           '바코드번호
Private mOrder      As String           '오더 전송후 남은오더
Private mNoOrder    As Boolean          '접수정보 존재유무 (True:비존재, False:존재)
Private mIsSending  As Boolean          '오더 전송후 남은 문자열의 존재유무

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

Public Property Get Order() As String
    Order = mOrder
End Property

Public Property Let Order(ByVal vData As String)
    mOrder = vData
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

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 문자열을 조회
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetOrder() As String
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim objIntNm    As clsIISIntNm      '장비별 검사항목 클래스
    Dim strIntBase  As String           '장비기준 검사명
    Dim strItems    As String           '장비기준 검사명 문자열
    Dim strTemp     As String

    '## 오더문자열 형식: ^^^1\^^^2
    '## 장비기준 검사명 형식: 정량(Assay No+N), 정성(Assay No+C)
    If mAccInfo.QcFg = "0" Then
        For Each objResult In mAccInfo.Results
            Set objIntNm = objResult.IntNm
            strIntBase = Mid$(objIntNm.IntBase, 1, Len(objIntNm.IntBase) - 1)
            If strIntBase <> strTemp Then
                If strItems = "" Then
                    strItems = "^^^" & strIntBase
                    strTemp = strIntBase
                Else
                    strItems = strItems & "\^^^" & strIntBase
                    strTemp = strIntBase
                End If
            End If
        Next
        Set objResult = Nothing
    ElseIf mAccInfo.QcFg = "1" Then
        For Each objQCResult In mAccInfo.QCResults
            Set objIntNm = objQCResult.IntNm
            strIntBase = Mid$(objIntNm.IntBase, 1, Len(objIntNm.IntBase) - 1)
            If strIntBase <> strTemp Then
                If strItems = "" Then
                    strItems = "^^^" & strIntBase
                    strTemp = strIntBase
                Else
                    strItems = strItems & "\^^^" & strIntBase
                    strTemp = strIntBase
                End If
            End If
        Next
        Set objQCResult = Nothing
    End If
    Set objIntNm = Nothing

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
    Set mAccInfo = Nothing
    mBarNo = ""
    mOrder = ""
    mNoOrder = False
    mIsSending = False
End Sub

Private Sub Class_Terminate()
    Set mAccInfo = Nothing
End Sub
