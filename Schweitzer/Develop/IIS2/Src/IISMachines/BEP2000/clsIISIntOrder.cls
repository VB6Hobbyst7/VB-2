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
'   파일명  : clsIISBEP2000.cls
'   작성자  : 오세원
'   내  용  : BEP2000 오더전송 클래스
'   작성일  : 2007-09-04
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
    Dim strIntBase()  As String           '장비기준 검사명
    Dim strItems      As String           '송신할 검사항목 문자열
    Dim strOutput     As String           '송신할 문자열
    Dim intCnt        As Integer
    Dim intCnt2       As Integer
    
    intCnt = 0
    If pAccInfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In pAccInfo.Results
            intCnt = intCnt + 1
            ReDim Preserve strIntBase(intCnt)
            If InStr(objResult.IntNm.IntBase, "II") > 0 Then
                '검사명 컬럼사이즈 문제???
                strIntBase(intCnt) = "DB_" & mGetP(objResult.IntNm.IntBase, 1, "|") & "-NEW"
            Else
                strIntBase(intCnt) = "DB_" & mGetP(objResult.IntNm.IntBase, 1, "|")
            End If
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In pAccInfo.QCResults
            ReDim Preserve strIntBase(intCnt)
            strIntBase(intCnt) = "DB_A" & mGetP(objResult.IntNm.IntBase, 1, "|")
        Next
        Set objQCResult = Nothing
    End If
    
    '17000680329,marrie,DB_A-CMV-G,DB_A-HSV-G,19650101,M,20030101010101
    For intCnt2 = 1 To intCnt
        If intCnt2 = intCnt Then
            strOutput = strOutput & mBarNo & "," & "" & "," & strIntBase(intCnt2) & "," & strIntBase(intCnt2) & "," & pAccInfo.AgeDay & "," & pAccInfo.Sex & "," & Format(Now, "yyyymmddhhmmss")
        Else
            strOutput = strOutput & mBarNo & "," & "" & "," & strIntBase(intCnt2) & "," & strIntBase(intCnt2) & "," & pAccInfo.AgeDay & "," & pAccInfo.Sex & "," & Format(Now, "yyyymmddhhmmss") & vbCrLf
        End If
    Next
    GetOrder = strOutput

End Function

'-----------------------------------------------------------------------------'
'   기능 : 클래스 멤버변수 초기화
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
End Sub
