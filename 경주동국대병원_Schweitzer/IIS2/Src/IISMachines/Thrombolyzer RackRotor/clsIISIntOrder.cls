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
'   내  용  : Thrombolyzer RackRotor 장비 오더정보 클래스
'   작성일  : 2005-03-04
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private Const RS As String = ""    'Record Separator
Private Const GS As String = ""    'Group Separator

Private mBarNo  As String   '바코드번호
Private mPos    As String   'Tube Position

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get Pos() As String
    Pos = mPos
End Property

Public Property Let Pos(ByVal vData As String)
    mPos = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 문자열 조회
'   인수 :
'       - pAccInfo : 접수내역 클래스
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim strIntBase  As String           '장비기준 검사명
    Dim strItems    As String           '송신할 검사항목
    Dim strOutput   As String           '송신할 데이터
    Dim lngCount    As Long             '송신할 검사항목 개수
    Dim strTemp     As String
    Dim i           As Long

    '## 장비기준 검사명
    '   PT(Time):1T, PT(%):1%, PT(INR):1I
    '   aPTT    :2T
    If pAccInfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In pAccInfo.Results
            strIntBase = Mid$(objResult.IntNm.IntBase, 1, Len(objResult.IntNm.IntBase) - 1)
            If strIntBase <> strTemp Then
                strItems = strItems & "|" & Format$(strIntBase, "@@")
                lngCount = lngCount + 1
                strTemp = strIntBase
            End If
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = Mid$(objQCResult.IntNm.IntBase, 1, Len(objQCResult.IntNm.IntBase) - 1)
            If strIntBase <> strTemp Then
                strItems = strItems & "|" & Format$(strIntBase, "@@")
                lngCount = lngCount + 1
                strTemp = strIntBase
            End If
        Next
        Set objQCResult = Nothing
    End If

    '## 최대 9개항목까지 전송가능, 9개 미만일때는 "99"로 전송
    For i = 1 To 9 - lngCount
        strItems = strItems & "|99"
    Next i

    strOutput = "0|" & Format$(mPos, "@@") & "|" & Format$(mBarNo, "!" & String$(30, "@")) & strItems
    strOutput = RS & strOutput & vbCrLf & GS
    strOutput = STX & vbCrLf & strOutput & GetChkSum(strOutput) & vbCrLf
    GetOrder = strOutput
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수정보가 없을때 오더정보 문자열 조회
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    Dim strOutput   As String   '송신할 데이터

    strOutput = "1|" & Format$(mPos, "@@") & "|" & Format$(mBarNo, "!" & String$(30, "@")) & _
                "|99|99|99|99|99|99|99|99|99"
    strOutput = RS & strOutput & vbCrLf & GS
    strOutput = STX & vbCrLf & strOutput & GetChkSum(strOutput) & vbCrLf
    GetNoOrder = strOutput
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
    mPos = ""
End Sub
