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
'   파일명  : clsIISBNII.cls
'   작성자  : 이상대
'   내  용  : BN II 오더전송 클래스
'   작성일  : 2005-01-27
'   버  전  :
'       1. 1.0.3: 이상대(2005-02-07)
'       2. 1.0.5: 이상대(2005-02-21)
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo      As String           '바코드번호
Private mOrder      As String           '오더정보 문자열

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
    Dim strBarNo      As String           '송신할 바코드번호
    Dim strPtId       As String           '송신할 환자ID
    Dim strSex        As String           '송신할 성별
    Dim strStatFg     As String           '송신할 응급유무
    Dim strTemp       As String
    
    '## 1.0.3: 이상대(2005-02-07)
    '   - 일반검체 일때는 환자의 성별, QC검체일때는 Space(1)을 전송하도록 수정
    '## 1.0.5: 이상대(2005-02-21)
    '   - 일반검체 일때는 응급유무를, QC검체일때는 항상 비응급을 전송하도록 수정
    '## 검사항목 문자열 생성
    If pAccInfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In pAccInfo.Results
            strIntBase = Mid$(objResult.IntNm.IntBase, 1, Len(objResult.IntNm.IntBase) - 1)
            If strTemp <> strIntBase Then
                strItems = strItems & strIntBase & "0"
                strTemp = strIntBase
            End If
        Next
        Set objResult = Nothing
        strSex = pAccInfo.Sex
        strStatFg = pAccInfo.StatFg
    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = Mid$(objQCResult.IntNm.IntBase, 1, Len(objQCResult.IntNm.IntBase) - 1)
            If strTemp <> strIntBase Then
                strItems = strItems & strIntBase & "0"
                strTemp = strIntBase
            End If
        Next
        Set objQCResult = Nothing
        strSex = Space(1)
        strStatFg = "0"
    End If
    
    strBarNo = Format$(mBarNo, "!" & String$(29, "@"))
    strPtId = Format$(pAccInfo.PtId, "!" & String$(30, "@"))
    mOrder = STX & "J" & strBarNo & Space(8) & strSex & strStatFg & Space(2) & "1" & _
             Space(40) & strPtId & Space(27) & strItems & Space(1) & vbCr & ETX & ACK
    GetOrder = mOrder
End Function

'-----------------------------------------------------------------------------'
'   기능 : 클래스 멤버변수 초기화
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mOrder = ""
End Sub
