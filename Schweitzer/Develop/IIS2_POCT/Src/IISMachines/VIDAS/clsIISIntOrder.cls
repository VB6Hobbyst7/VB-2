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
'   내  용  : VIDAS 장비 오더정보 클래스
'   작성일  : 2004-08-23
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo      As String   '바코드번호
Private mPtId       As String   '환자Id
Private mRow        As Long     '현재 바코드번호의 Row No

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get PtId() As String
    PtId = mPtId
End Property

Public Property Let PtId(ByVal vData As String)
    mPtId = vData
End Property

Public Property Get Row() As Long
    Row = mRow
End Property

Public Property Let Row(ByVal vData As Long)
    mRow = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 문자열 조회 - 1.1.0: 이상대(2005-06-30)
'   인수 :
'       - pAccInfo : 접수내역 객체
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByRef pAccInfo As clsIISAccInfo) As String
'    Dim objResult   As clsIISResult     '결과내역 객체
'    Dim objQCResult As clsIISQCResult   'QC결과내역 객체
'    Dim strIntBase  As String           '장비기준 검사명
'    Dim strOrder    As String           '오더정보 문자열
'
'    If pAccInfo.QcFg = "0" Then     '## 일반검체
'        For Each objResult In pAccInfo.Results
'            If objResult.SndChk = False Then
'                strIntBase = objResult.IntNm.IntBase
'                strIntBase = Mid$(strIntBase, 1, Len(strIntBase) - 1)
'                objResult.SndChk = True
'                Exit For
'            End If
'        Next
'        Set objResult = Nothing
'    ElseIf pAccInfo.QcFg = "1" Then '## QC검체
'        For Each objQCResult In pAccInfo.QCResults
'            If objResult.SndChk = False Then
'                strIntBase = objQCResult.IntNm.IntBase
'                strIntBase = Mid$(strIntBase, 1, Len(strIntBase) - 1)
'                objQCResult.SndChk = True
'                Exit For
'            End If
'        Next
'        Set objQCResult = Nothing
'    End If
'
'    strOrder = Chr(30) & "mtmpr|pi" & pAccInfo.PtId & "||si|ci" & mBarNo & "|" & _
'               "rt" & strIntBase & "|" & Chr(29)
'    GetOrder = STX & strOrder & GetChkSum(strOrder) & ETX
End Function

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 문자열 조회 - 1.1.0: 이상대(2005-06-30)
'   인수 :
'       - pIntBase : 장비기준 검사명
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetOrderX(ByVal pIntBase As String) As String
    Dim strOrder As String      '오더정보 문자열
    
    strOrder = Chr(30) & "mtmpr|pi" & mPtId & "||si|ci" & mBarNo & "|" & "rt" & pIntBase & "|" & Chr(29)
    GetOrderX = STX & strOrder & GetChkSum(strOrder) & ETX
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
        GetChkSum = LCase(Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2))
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : 클래스 멤버변수 초기화
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mPtId = ""
    mRow = 0
End Sub
