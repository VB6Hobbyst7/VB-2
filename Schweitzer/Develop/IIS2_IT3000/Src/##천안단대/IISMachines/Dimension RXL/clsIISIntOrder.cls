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
'   내  용  : Dimension RXL 장비 오더정보 클래스
'   작성일  : 2004-07-14
'   버  전  :
'       1. 1.0.1: 이상대(2004-12-16)
'          - Cardiac QC(MMB,MYO,CTNI)인 경우에는 QC1,2,3를 사용해서 IISQCNORMAL
'            검체구분 추가!
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo  As String   '바코드번호

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : No Request Message 조회
'   반환 : No Request Message
'-----------------------------------------------------------------------------'
Public Function GetNoRequest() As String
    GetNoRequest = STX & "N6A" & ETX
End Function

'-----------------------------------------------------------------------------'
'   기능 : Wait Poll Message 조회
'   반환 : Wait Poll Message
'-----------------------------------------------------------------------------'
Public Function GetWaitPoll() As String
    GetWaitPoll = STX & "W73" & ETX
End Function

'-----------------------------------------------------------------------------'
'   기능 : Result Acceptance Message 조회
'   반환 : Result Acceptance Message
'-----------------------------------------------------------------------------'
Public Function GetResultAccept() As String
    GetResultAccept = STX & "MAE2" & ETX
End Function

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 문자열 조회
'   인수 :
'       - pAccInfo : 접수내역 클래스
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult     As clsIISResult     '결과내역 클래스
    Dim objQCResult   As clsIISQCResult   'QC결과내역 클래스
    Dim objIntNm      As clsIISIntNm      '장비별 검사항목 클래스
    Dim strIntBase    As String           '전송할 검사항목
    Dim strSampleType As String           '전송할 검체종류
    Dim strOrder      As String           '오더정보 문자열
    
    '## OSMO 검사는 NA, GLU, BUN검사결과의 계산된 검사이므로 오더정보를 전송하지 않음
    If pAccInfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In pAccInfo.Results
            Set objIntNm = objResult.IntNm
            If objIntNm.IntBase = "OSMO" Then
                pAccInfo.SendCnt = pAccInfo.SendCnt - 1
            Else
                strIntBase = strIntBase & objIntNm.IntBase & ""
            End If
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In pAccInfo.QCResults
            Set objIntNm = objQCResult.IntNm
            
            '## QC중 HA1C는 검체지정을 "W"로 해야함
            Select Case objIntNm.IntBase
                Case "OSMO"
                    pAccInfo.SendCnt = pAccInfo.SendCnt - 1
                Case "HA1C"
                    pAccInfo.SpcNm = "B"
                    strIntBase = strIntBase & objIntNm.IntBase & ""
                Case Else
                    strIntBase = strIntBase & objIntNm.IntBase & ""
            End Select
        Next
        Set objQCResult = Nothing
    End If
    Set objIntNm = Nothing

    '## 1.0.1: 이상대(2004-12-16)
    '   - IISQCNORMAL 검체구분 추가
    '## 검체종류 조회
    Select Case Mid$(pAccInfo.SpcNm, 1, IISSPCLEN)
        Case IISSPCSERUM:   strSampleType = "1"
        Case IISSPCPLASMA:  strSampleType = "2"
        Case IISSPCURINE:   strSampleType = "3"
        Case IISSPCCSF:     strSampleType = "4"
        Case IISSPCBLOOD:   strSampleType = "W"
        Case IISQCLOW:      strSampleType = "5"
        Case IISQCNORMAL:   strSampleType = "6"
        Case IISQCHIGH:     strSampleType = "7"
        Case Else:          strSampleType = "1"
    End Select

    strOrder = "D00A" & pAccInfo.PtId & "" & mBarNo & "" & strSampleType & "01**1" & _
               pAccInfo.SendCnt & "" & strIntBase

    GetOrder = STX & strOrder & GetChkSum(strOrder) & ETX
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
End Sub

