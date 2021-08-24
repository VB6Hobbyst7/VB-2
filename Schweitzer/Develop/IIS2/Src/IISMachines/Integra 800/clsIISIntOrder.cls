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
'   내  용  : Integra 800 장비 오더정보 클래스
'   작성일  : 2005-07-01
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo As String

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 문자열을 조회
'   인수 :
'       - pAccInfo : 접수내역 객체
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim strIntBase  As String           '장비기준 검사명
    Dim strItems    As String           '전송할 검사항목
    Dim strStat     As String           '응급유무(A:Routin, S:Stat)
    Dim strOutput   As String           '송신할 데이터
    
    If pAccInfo.QcFg = "0" Then     '## 일반검체
        For Each objResult In pAccInfo.Results
            strIntBase = "55 " & Format$(objResult.IntNm.IntBase, "000") & mEndChar
            strItems = strItems & strIntBase
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then '## QC검체
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = "55 " & Format$(objQCResult.IntNm.IntBase, "000") & mEndChar
            strItems = strItems & strIntBase
        Next
        Set objQCResult = Nothing
    End If
    
    strStat = IIf(pAccInfo.StatFg = "1", "S", "A")
    strOutput = SOH & mEndChar & _
                "09 " & Format$(mEqpNm, "!" & String$(16, "@")) & " 10" & mEndChar & _
                STX & mEndChar & _
                "53 " & Format$(mBarNo, "!" & String(15, "@")) & " 00/00/0000" & mEndChar & _
                "54 00000 00 " & strStat & mEndChar & strItems & _
                ETX & mEndChar & _
                EOT & mEndChar
    
              ' "54 000 00 " & strStat & mEndChar & strItems & _

    
    GetOrder = strOutput
End Function

'-----------------------------------------------------------------------------'
'   기능 : 클래스 멤버변수 초기화
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
End Sub
