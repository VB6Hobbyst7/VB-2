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
Option Explicit

Private mAccInfo    As clsIISAccInfo   '접수내역 클래스

Private mBarNo      As String   '바코드번호
Private mSeq        As String   'Seq
Private mDiskNo     As String   'Disk No
Private mTubePos    As String   'Tube Position
Private mNoOrder    As String   '접수정보 존재유무 (True:비존재, False:존재)
Private mIsSending  As Boolean  '
Private mOrder      As String   '오더를 전송후 남은오더

Public Property Get AccInfo() As clsIISAccInfo
    Set AccInfo = mAccInfo
End Property

Public Property Let AccInfo(ByVal vData As clsIISAccInfo)
    Set mAccInfo = vData
End Property

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

Public Property Get DiskNo() As String
    DiskNo = mDiskNo
End Property

Public Property Let DiskNo(ByVal vData As String)
    mDiskNo = vData
End Property

Public Property Get TubePos() As String
    TubePos = mTubePos
End Property

Public Property Let TubePos(ByVal vData As String)
    mTubePos = vData
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
Public Function GetOrder() As String
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim objIntNm    As clsIISIntNm      '장비별 검사항목 클래스
    Dim strIntBase  As String           '장비기준 검사명 문자열
    
    '## 오더문자열 형식: ^^^1\^^^2
    If mAccInfo.QcFg = "0" Then
        For Each objResult In mAccInfo.Results
            Set objIntNm = objResult.IntNm
            If strIntBase = "" Then
                strIntBase = "^^^" & CStr(CInt(Mid$(objIntNm.IntBase, 1, Len(objIntNm.IntBase) - 1)))
            Else
                strIntBase = strIntBase & "\^^^" & CStr(CInt(Mid$(objIntNm.IntBase, 1, Len(objIntNm.IntBase) - 1)))
            End If
        Next
        Set objResult = Nothing
    ElseIf mAccInfo.QcFg = "1" Then
        For Each objQCResult In mAccInfo.QCResults
            Set objIntNm = objQCResult.IntNm
            If strIntBase = "" Then
                strIntBase = "^^^" & CStr(CInt(Mid$(objIntNm.IntBase, 1, Len(objIntNm.IntBase) - 1)))
            Else
                strIntBase = strIntBase & "\^^^" & CStr(CInt(Mid$(objIntNm.IntBase, 1, Len(objIntNm.IntBase) - 1)))
            End If
        Next
        Set objQCResult = Nothing
    End If
    Set objIntNm = Nothing
    
    GetOrder = strIntBase
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
    mDiskNo = ""
    mTubePos = ""
    mNoOrder = False
    mIsSending = False
    mOrder = ""
    Set mAccInfo = Nothing
End Sub

Private Sub Class_Terminate()
    Set mAccInfo = Nothing
End Sub
