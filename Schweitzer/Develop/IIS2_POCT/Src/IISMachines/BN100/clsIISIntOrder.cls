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
'   내  용  : BN100 장비 오더정보 클래스
'   작성일  : 2004-07-19
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo As String    '바코드번호
Private mSeq   As String    '오더를 전송한 Spread의 Row

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get Seq() As String
    Seq = mSeq
End Property

Public Property Let Seq(ByVal vData As String)
    mSeq = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 문자열 조회
'   인수 :
'       - 접수내역 클래스
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByRef pAccinfo As clsIISAccInfo) As String
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim strIntBase  As String           '장비기준 검사명 문자열
    Dim strOrder    As String           '오더정보 문자열
    Dim strTemp1    As String
    Dim strTemp2    As String
    
    If pAccinfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In pAccinfo.Results
            strTemp1 = Mid$(objResult.IntNm.IntBase, 1, Len(objResult.IntNm.IntBase) - 1)
            
            If strTemp1 <> strTemp2 Then
                strIntBase = strIntBase & Format$(strTemp1, "00")
                strTemp2 = strTemp1
            Else
                pAccinfo.SendCnt = pAccinfo.SendCnt - 1
            End If
        Next
        Set objResult = Nothing
    ElseIf pAccinfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In pAccinfo.QCResults
            strTemp1 = Mid$(objQCResult.IntNm.IntBase, 1, Len(objQCResult.IntNm.IntBase) - 1)
            
            If strTemp1 <> strTemp2 Then
                strIntBase = strIntBase & Format$(strTemp1, "00")
                strTemp2 = strTemp1
            Else
                pAccinfo.SendCnt = pAccinfo.SendCnt - 1
            End If
        Next
        Set objQCResult = Nothing
    End If
    strIntBase = strIntBase & String$(40 - Len(strIntBase), Space(1))
    
    strOrder = STX & "J" & mBarNo & String$(29 - SPCLEN, Space(1)) & Space(8) & _
               strIntBase & Space(1) & vbCr & ETX
    GetOrder = strOrder
End Function

'-----------------------------------------------------------------------------'
'   기능 : Request Result Message 문자열 조회
'   인수 :
'       - 접수내역 클래스
'   반환 :
'       - Request Result Message
'       - 더이상 요구할 검사항목이 없을경우에는 "-1" 반환
'-----------------------------------------------------------------------------'
Public Function GetReqResult(ByVal pAccinfo As clsIISAccInfo) As String
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim strIntBase  As String           '장비기준 검사명 문자열
    Dim strReqMsg   As String           'Request Result Message
    Dim strTemp1    As String
    Dim strTemp2    As String
    
    '## 결과전송요구시 검사항목을 한개씩 요구해야 한다
    If pAccinfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In pAccinfo.Results
            If objResult.SndChk = False Then
                strTemp1 = Mid$(objResult.IntNm.IntBase, 1, Len(objResult.IntNm.IntBase) - 1)
                
                If strTemp1 <> strTemp2 Then
                    strIntBase = Format$(strTemp1, "00")
                    strTemp2 = strTemp1
                End If
                objResult.SndChk = True
                Exit For
            End If
        Next
        Set objResult = Nothing
    ElseIf pAccinfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In pAccinfo.QCResults
            If objQCResult.SndChk = False Then
                strTemp1 = Mid$(objQCResult.IntNm.IntBase, 1, Len(objQCResult.IntNm.IntBase) - 1)
                
                If strTemp1 <> strTemp2 Then
                    strIntBase = Format$(strTemp1, "00")
                    strTemp2 = strTemp1
                End If
                objQCResult.SndChk = True
                Exit For
            End If
        Next
        Set objQCResult = Nothing
    End If
    
    If strIntBase = "" Then
        strReqMsg = "-1"
    Else
        strReqMsg = STX & "D" & CStr(mBarNo) & String$(29 - SPCLEN, Space(1)) & Space(1) & _
                    strIntBase & Space(1) & vbCr & ETX
    End If
    GetReqResult = strReqMsg
End Function

'-----------------------------------------------------------------------------'
'   기능 : Positive Acknowledge Message 조회
'   반환 : Positive Acknowledge Message
'-----------------------------------------------------------------------------'
Public Function GetPositiveMsg() As String
    GetPositiveMsg = STX & "A " & vbCr & ETX
End Function

'-----------------------------------------------------------------------------'
'   기능 : Positive Acknowledge Message 조회
'   반환 : Positive Acknowledge Message
'-----------------------------------------------------------------------------'
Public Function GetNegativeMsg() As String
    GetNegativeMsg = STX & "E " & vbCr & ETX
End Function

'-----------------------------------------------------------------------------'
'   기능 : 클래스 멤버변수 초기화
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mSeq = ""
End Sub
