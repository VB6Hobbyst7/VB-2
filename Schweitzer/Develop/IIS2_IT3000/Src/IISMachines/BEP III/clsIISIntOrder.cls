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
'   파일명  : clsIISBEPIII.cls
'   작성자  : 이상대
'   내  용  : BEP III 오더전송 클래스
'   작성일  : 2005-01-27
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo      As String           '바코드번호

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 문자열 조회
'   인수 :
'       - pAccInfo : 접수내역 클래스
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult     As clsIISResult     '결과내역 객체
    Dim objQCResult   As clsIISQCResult   'QC결과내역 객체
    Dim strIntBase    As String           '장비기준 검사명
    Dim strItems      As String           '송신할 검사항목 문자열
    Dim strOutput     As String           '송신할 문자열

    If pAccInfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In pAccInfo.Results
            strIntBase = Format$(mGetP(objResult.IntNm.IntBase, 1, "|"), "!" & String$(6, "@"))
            If strItems = "" Then
                strItems = strIntBase
            Else
                strItems = strItems & ";" & strIntBase
            End If
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = Format$(mGetP(objQCResult.IntNm.IntBase, 1, "|"), "!" & String$(6, "@"))
            If strItems = "" Then
                strItems = strIntBase
            Else
                strItems = strItems & ";" & strIntBase
            End If
        Next
        Set objQCResult = Nothing
    End If
    
    strOutput = pAccInfo.PtId & ";" & mBarNo & ";" & strItems
    GetOrder = strOutput
End Function

'-----------------------------------------------------------------------------'
'   기능 : 클래스 멤버변수 초기화
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
End Sub
