VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIntOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : clsIISIntOrder.cls
'   작성자  : 이상대
'   내  용  : Elecsys 2010 장비 오더정보 클래스
'   작성일  : 2004-08-13
'   버  전  :
'       1. 1.0.1: 이상대(2004-12-07)
'-----------------------------------------------------------------------------'

Option Explicit

'Private mAccInfo    As clsIISAccInfo   '접수내역 클래스

Private mBarNo      As String   '바코드번호
Private mSpcType    As String   '검체유형
Private mSeq        As String   'Seq
Private mRackNo     As String   'Rack No
Private mTubePos    As String   'Tube Position
Private mKind       As String   'Run(R1), Rerun(R2), QC(Q) 구분
Private mPriority    As String  '응급구분(R:Routine, S:Stat)
Private mNoOrder    As String   '접수정보 존재유무 (True:비존재, False:존재)
Private mIsSending  As Boolean  '오더 전송후 남은 문자열의 존재유무
Private mOrder      As String   '오더 전송후 남은오더

'Public Property Get AccInfo() As clsIISAccInfo
'    Set AccInfo = mAccInfo
'End Property
'
'Public Property Let AccInfo(ByVal vData As clsIISAccInfo)
'    Set mAccInfo = vData
'End Property

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get SpcType() As String
    SpcType = mSpcType
End Property

Public Property Let SpcType(ByVal vData As String)
    mSpcType = vData
End Property

Public Property Get Seq() As String
    Seq = mSeq
End Property

Public Property Let Seq(ByVal vData As String)
    mSeq = vData
End Property

Public Property Get RackNo() As String
    RackNo = mRackNo
End Property

Public Property Let RackNo(ByVal vData As String)
    mRackNo = vData
End Property

Public Property Get TubePos() As String
    TubePos = mTubePos
End Property

Public Property Let TubePos(ByVal vData As String)
    mTubePos = vData
End Property

Public Property Get Kind() As String
    Kind = mKind
End Property

Public Property Let Kind(ByVal vData As String)
    mKind = vData
End Property

Public Property Get Priority() As String
    Priority = mPriority
End Property

Public Property Let Priority(ByVal vData As String)
    mPriority = vData
End Property

Public Property Get NoOrder() As Boolean
    NoOrder = mNoOrder
End Property

Public Property Let NoOrder(ByVal vData As Boolean)
    mNoOrder = vData
End Property

Public Property Get IsSending() As Boolean
    IsSending = mIsSending
End Property

Public Property Let IsSending(ByVal vData As Boolean)
    mIsSending = vData
End Property

Public Property Get Order() As String
    Order = mOrder
End Property

Public Property Let Order(ByVal vData As String)
    mOrder = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 문자열을 조회
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal f_Channel As Variant) As String
    Dim strIntBase  As String           '장비기준 검사명
    Dim strItems    As String           '송신할 검사항목
'    Dim strTemp     As String
    Dim blnISE      As Boolean          'ISE모듈 검사항목 유무
    Dim i As Integer

'    '## 오더문자열 형식: ^^^1/\^^^2
'    For Each objResult In mAccInfo.Results
'        Set objIntNm = objResult.IntNm
'        strIntBase = objIntNm.IntBase
'
'        '## 1.Na(989), K(990), Cl(991) 검사항목은 (Na,K), (Na,K,Cl) 조합으로만 오더를 전송
'        '## 2.계산공식 오더는 전송하지 않음(961~968)
'        If strIntBase = "989" Or strIntBase = "990" Or strIntBase = "991" Then
'            '## ISE검사항목 나중에 일괄전송
'            blnISE = True
'        ElseIf CInt(strIntBase) >= 961 And CInt(strIntBase) <= 968 Then
'            '## 계산공식 검사항목은 전송하지 않음
'        Else
'            If strItems = "" Then
'                strItems = "^^^" & strIntBase & "/"
'            Else
'                strItems = strItems & "\^^^" & strIntBase & "/"
'            End If
'        End If
'    Next
'
'    '## 1.blnISE=True이면 ISE검사항목(Na,K,Cl) 추가
'    '## 2.검체가 Serum이면 Serum Index 검사항목 추가(병원별로 설정할수 있도록 할것!)
'    If blnISE = True Then
'        If strItems = "" Then
'            strItems = "^^^989/\^^^990/\^^^991/"
'        Else
'            strItems = strItems & "\^^^989/\^^^990/\^^^991/"
'        End If
'    End If
'
'    GetOrder = strItems

    For i = 0 To UBound(f_Channel)
        strIntBase = f_Channel(i)
        '## 1.Na(989), K(990), Cl(991) 검사항목은 (Na,K), (Na,K,Cl) 조합으로만 오더를 전송
        '## 2.계산공식 오더는 전송하지 않음(961~968)
        If IsNumeric(strIntBase) Then
            If strIntBase = "989" Or strIntBase = "990" Or strIntBase = "991" Then
                '## ISE검사항목 나중에 일괄전송
                blnISE = True
            ElseIf CInt(strIntBase) >= 961 And CInt(strIntBase) <= 968 Then
                '## 계산공식 검사항목은 전송하지 않음
            Else
                If strItems = "" Then
                    strItems = "^^^" & strIntBase & "/"
                Else
                    strItems = strItems & "\^^^" & strIntBase & "/"
                End If
            End If
        End If
    Next
    
    '## 1.blnISE=True이면 ISE검사항목(Na,K,Cl) 추가
    '## 2.검체가 Serum이면 Serum Index 검사항목 추가(병원별로 설정할수 있도록 할것!)
    If blnISE = True Then
        If strItems = "" Then
            strItems = "^^^989/\^^^990/\^^^991/"
        Else
            strItems = strItems & "\^^^989/\^^^990/\^^^991/"
        End If
    End If

    GetOrder = strItems

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
    mSeq = ""
    mRackNo = ""
    mTubePos = ""
    mNoOrder = False
    mIsSending = False
    mOrder = ""
'    Set mAccInfo = Nothing
End Sub

Private Sub Class_Terminate()
'    Set mAccInfo = Nothing
End Sub

