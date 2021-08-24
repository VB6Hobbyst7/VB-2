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
'   내  용  : CX3 Delta 장비 오더정보 클래스
'   작성일  : 2004-10-14
'   버  전  :
'       1. 1.0.1: 이상대(2004-12-14)
'-----------------------------------------------------------------------------'

Option Explicit

Private mAccInfo    As clsIISAccInfo    '접수내역 클래스
Private mDeviceId   As String           'Device ID
Private mBarNo      As String           '바코드번호
Private mSendCnt    As String           '전송할 검체개수

Public Property Get AccInfo() As clsIISAccInfo
    Set AccInfo = mAccInfo
End Property

Public Property Let AccInfo(ByVal vData As clsIISAccInfo)
    Set mAccInfo = vData
End Property

Public Property Get DeviceId() As String
    DeviceId = mDeviceId
End Property

Public Property Let DeviceId(ByVal vData As String)
    mDeviceId = vData
End Property

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get SendCnt() As Long
    SendCnt = mSendCnt
End Property

Public Property Let SendCnt(ByVal vData As Long)
    mSendCnt = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 문자열 조회
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetOrder() As String
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim strIntBase  As String           '장비기준 검사명
    Dim strItems    As String           '송신할 검사항목 문자열
    Dim strTestTp   As String           '송신할 Test Type(CO:Control, RO:Rountine, SC:STAT Control, ST:STAT)
    Dim strSpcTp    As String           '송신할 Sample Type(SE:Serum, SF:Spinal Fluid, UR:Urine, TU:Timed Urine, PL:Plasma)
    Dim strPtId     As String           '송신할 환자ID
    Dim strAge      As String           '송신할 Age
    Dim strSex      As String           '송신할 Sex
    Dim strOutput   As String           '송신한 데이터
    Dim strSendCnt  As Long
    
    '## 검사항목 문자열 생성
    If mAccInfo.QcFg = "0" Then
        For Each objResult In mAccInfo.Results
            strIntBase = objResult.IntNm.IntBase
            strItems = strItems & "," & Format$(strIntBase, "!@@@@") & ",0"
        Next
        Set objResult = Nothing
    ElseIf mAccInfo.QcFg = "1" Then
        For Each objQCResult In mAccInfo.QCResults
            strIntBase = objQCResult.IntNm.IntBase
            strItems = strItems & "," & Format$(strIntBase, "!@@@@") & ",0"
        Next
        Set objQCResult = Nothing
    End If
    
    '## Test Type, Sample Type 조회
    strTestTp = IIf(mAccInfo.StatFg = "1", "ST", "RO")
    Select Case Mid$(mAccInfo.SpcNm, 1, IISSPCLEN)
        Case IISSPCSERUM:   strSpcTp = "SE"
        Case IISSPCURINE:   strSpcTp = "UR"
        Case IISSPCPLASMA:  strSpcTp = "PL"
        Case IISSPCFLUID:   strSpcTp = "SF"
        Case IISSPCCSF:     strSpcTp = "SF"
        Case Else:          strSpcTp = "SE"
    End Select
    
    '## 1.0.1: 이상대(2004-12-14)
    '   - QC오더인 경우 나이, 성별을 강제 셋팅!
    '## 환자ID, 나이 조회
    strPtId = Format$(mAccInfo.PtId, "!" & String$(12, "@"))
    If mAccInfo.QcFg = "0" Then
        strAge = Format$(Mid$(mAccInfo.Age, 1, Len(mAccInfo.Age) - 1), "000")
        strSex = mAccInfo.Sex
    ElseIf mAccInfo.QcFg = "1" Then
        strAge = "000"
        strSex = "M"
    End If
    
    strOutput = "[" & mDeviceId & ",701,01,00,00,0," & strTestTp & "," & strSpcTp & "," & mBarNo & _
                "," & Space(20) & "," & Space(25) & "," & Space(25) & "," & Space(18) & "," & _
                Space(15) & "," & Space(1) & "," & strPtId & "," & Space(18) & "," & Space(6) & _
                "," & Space(4) & "," & Space(20) & "," & strAge & ",5," & Space(6) & "," & strSex & _
                "," & Space(25) & "," & Space(7) & "," & Space(4) & "," & Space(4) & "," & Space(6) & _
                "," & Format$(mAccInfo.SendCnt, "000") & strItems & "]"
    strOutput = strOutput & GetChkSum(strOutput) & vbCrLf
    GetOrder = strOutput
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
        lngChkSum = lngChkSum + Asc(Mid$(pMsg, i, 1))
    Next i
    
    lngChkSum = lngChkSum Mod 256
    lngChkSum = 256 - lngChkSum
    
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
End Sub
