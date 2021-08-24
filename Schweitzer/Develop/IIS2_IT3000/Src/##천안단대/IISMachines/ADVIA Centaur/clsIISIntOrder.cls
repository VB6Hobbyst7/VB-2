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
'   내  용  : ADVIA Centaur 오더전송 클래스
'   작성일  : 2004-08-10
'   버  전  :
'       1.0.1 : 이상대 (2004-11-08)
'               같은 검사항목에 대해서는 한번만 오더를 전송하도록 수정(GetOrder)
'-----------------------------------------------------------------------------'


Option Explicit

Private mAccInfo    As clsIISAccInfo   '접수내역 클래스

Private mBarNo      As String   '바코드번호
Private mNoOrder    As String   '접수정보 존재유무 (True:비존재, False:존재)
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

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 문자열을 조회
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetOrder() As String
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim strIntBase  As String           '장비기준 검사명
    Dim strItems    As String           '송신할 장비기준 검사명 문자열
    Dim strTemp     As String

    '## 1.0.1: 같은 검사항목에 대해서는 한번만 오더를 전송하도록 수정
    If mAccInfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In mAccInfo.Results
            strIntBase = objResult.IntNm.IntBase
            strIntBase = Mid$(strIntBase, 1, Len(strIntBase) - 1)
            
            If strTemp <> strIntBase Then
                If strItems = "" Then
                    strItems = "^^^" & strIntBase
                Else
                    strItems = strItems & "\^^^" & strIntBase
                End If
                strTemp = strIntBase
            End If
        Next
        Set objResult = Nothing
    ElseIf mAccInfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In mAccInfo.QCResults
            strIntBase = objQCResult.IntNm.IntBase
            strIntBase = Mid$(strIntBase, 1, Len(strIntBase) - 1)
            
            If strTemp <> strIntBase Then
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
'   기능 : 환자ID 조회
'   반환 :
'       - 환자ID
'       - -1: 접수정보 객체가 Null일때
'-----------------------------------------------------------------------------'
Public Function GetPtId() As String
    If Not (mAccInfo Is Nothing) Then
        GetPtId = mAccInfo.PtId
    Else
        GetPtId = "-1"
    End If
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
    mNoOrder = False
    mOrder = ""
    
    Set mAccInfo = Nothing
End Sub

Private Sub Class_Terminate()
    Set mAccInfo = Nothing
End Sub
