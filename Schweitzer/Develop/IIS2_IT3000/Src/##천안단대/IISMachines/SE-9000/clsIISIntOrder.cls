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
'   내  용  : SE-9000 장비 오더정보 클래스
'   작성일  : 2004-10-13
'   버  전  :
'       1. 1.0.1: 이상대(2004-12-07)
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo  As String   '바코드번호
Private mRackNo As String   'Rack No
Private mPos    As String   'Tube Position

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

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 문자열을 조회
'   인수 :
'       - pAccInfo : 접수내역 클래스
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim strIntBase  As String           '장비기준 검사명
    Dim strItems    As String           '장비기준 검사명 문자열
    Dim strBarNo    As String           '송신할 바코드번호
    Dim strPtId     As String           '송신할 환자ID
    Dim strSex      As String           '송신할 성별
    Dim strOutput   As String           '송신할 데이터

    strItems = String$(24, "0")
    If pAccInfo.QcFg = "0" Then     '## 일반검체
        For Each objResult In pAccInfo.Results
            strIntBase = mGetP(objResult.IntNm.IntBase, 1, "|")
            Mid$(strItems, CLng(strIntBase), 1) = 1
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then '## QC검체
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = mGetP(objQCResult.IntNm.IntBase, 1, "|")
            Mid$(strItems, CLng(strIntBase), 1) = 1
        Next
        Set objQCResult = Nothing
    End If
    
    '## 1.0.1: 이상대(2004-12-07)
    '   - 환자ID가 영문인 경우 "0"붙지않는 문제수정
    strBarNo = Format$(mBarNo, String$(13, "0"))
    strPtId = pAccInfo.PtId
    If IsNumeric(strPtId) Then
        strPtId = Format$(strPtId, String$(13, "0"))
    Else
        strPtId = String$(13 - Len(strPtId), "0") & strPtId
    End If
    strSex = IIf(UCase$(pAccInfo.Sex) = "M", 1, 2)
    
    strOutput = STX & "S1        " & strBarNo & mRackNo & mPos & "1" & strPtId & Space(25) & _
                strSex & Space(71) & strItems & Space(7) & ETX
    GetOrder = strOutput
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수정보가 없을때 오더문자열을 조회
'   반환 : 접수정보가 없을때 오더문자열
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    Dim strBarNo    As String   '송신할 바코드번호
    Dim strOutput   As String   '송신할 데이터
    
    strBarNo = Format$(mBarNo, String$(13, "0"))
    
    strOutput = STX & "S0        " & strBarNo & mRackNo & mPos & "1" & Space(110) & _
                String$(24, "0") & Space(7) & ETX
    Debug.Print "Len(strOutput)=" & Len(strOutput)
    GetNoOrder = strOutput
End Function

'-----------------------------------------------------------------------------'
'   기능 : 클래스 멤버변수 초기화
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mRackNo = ""
    mPos = ""
End Sub
