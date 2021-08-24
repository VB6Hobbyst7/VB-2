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
'   내  용  : ADVIA 120 장비 오더정보 클래스
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mMT     As Long     'MT Value
Private mBarNo  As String   '바코드번호

Public Property Get MT() As String
    MT = Chr(mMT)
End Property

Public Property Let MT(ByVal vData As String)
    mMT = Asc(vData)
End Property

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
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim objIntNm    As String           '장비별 검사항목 클래스
    Dim strIntBase  As String           '전송할 검사항목
    Dim strStatFg   As String           '응급유무
    Dim strOrder    As String           '오더정보 문자열
    
    '## 전송할 검사항목 조회
    '## 장비기준 검사명에 12자리 공백은 붙히지 않음
    If pAccInfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In pAccInfo.Results
            strIntBase = strIntBase & objResult.IntNm.IntBase
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = strIntBase & objQCResult.IntNm.IntBase
        Next
        Set objQCResult = Nothing
    End If
    
    '## 응급유무 조회 (비응급항목은 공백으로)
    strStatFg = IIf(pAccInfo.QcFg = "1", " U", "  ")
    
    '## 전송하지 않는 항목은 모두 공백으로 전송
    strOrder = GetMT & "Y " & strStatFg & "  " & Format$(mBarNo, String$(14, "0")) & Space(25) & _
               Format$(pAccInfo.PtId, String$(14, "0")) & Space(75) & vbCrLf & strIntBase & vbCrLf
    
    GetOrder = STX & strOrder & GetLRC(strOrder) & ETX

End Function

'-----------------------------------------------------------------------------'
'   기능 : 오더정보가 없을경우의 문자열 조회
'   반환 : No Workorder 문자열
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    Dim strNoOrder As String    'No Workorder 문자열
    
    strNoOrder = GetMT & "N R " & Format$(mBarNo, String$(14, "0")) & vbCrLf
    strNoOrder = STX & strNoOrder & GetLRC(strNoOrder) & ETX
    GetNoOrder = strNoOrder
    
End Function

'------------------------------------------------------------------------------'
'   기능 : Initialization Message 조회
'   반환 : Initialization Message
'------------------------------------------------------------------------------'
Public Function GetInit() As String
    Dim strInit As String   'Initialization 문자열
    
    mMT = &H30
    strInit = MT & "I " & vbCrLf
    strInit = STX & strInit & GetLRC(strInit) & ETX
    GetInit = strInit
End Function

'-----------------------------------------------------------------------------'
'   기능 : Token Transfer Message 조회
'   반환 : Token Transfer Message
'-----------------------------------------------------------------------------'
Public Function GetToken() As String
    Dim strToken As String  'Token Transfer 문자열
    
    strToken = GetMT & "S" & Space(10) & vbCrLf
    strToken = STX & strToken & GetLRC(strToken) & ETX
    GetToken = strToken
End Function

'-----------------------------------------------------------------------------'
'   기능 : Result Validation Message 조회
'   반환 : Result Validation Message
'-----------------------------------------------------------------------------'
Public Function GetResultValid() As String
    Dim strValid As String  'Result Validation 문자열
    
    strValid = GetMT & "Z" & Space(17) & " 0" & vbCrLf
    strValid = STX & strValid & GetLRC(strValid) & ETX
    GetResultValid = strValid
End Function

'-----------------------------------------------------------------------------'
'   기능 : MT 조회
'   반환 : MT
'-----------------------------------------------------------------------------'
Public Function GetMT() As String
    If mMT = 0 Or mMT = 90 Then '## H30=0, H5A=90
        mMT = &H30
    Else
        mMT = mMT + 1
    End If
    
    GetMT = Chr(mMT)
End Function

'-----------------------------------------------------------------------------'
'   기능 : LRC(Longgitudinal Redundancy Check) 조회
'   반환 : LRC
'-----------------------------------------------------------------------------'
Public Function GetLRC(ByVal pData As String) As String
    Dim lngLRC  As Long     'LRC
    Dim i       As Long
    
    For i = 1 To Len(pData)
        lngLRC = lngLRC Xor Asc(Mid$(pData, i, 1))
    Next i
    
    If lngLRC = 3 Then
        GetLRC = Chr(&H7F)
    Else
        GetLRC = Chr(lngLRC)
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : 클래스 멤버변수 초기화
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
End Sub
