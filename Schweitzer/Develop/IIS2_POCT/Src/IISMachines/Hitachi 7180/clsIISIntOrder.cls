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
'   작성자  : 오세원
'   내  용  : Hitachi 7180 장비 오더정보 클래스
'   작성일  : 2008-07-09
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo      As String   '바코드번호
Private mSampleData As String   'Sample Data
Private mFu         As String   'Function Character
Private mDiskNo     As String   'Disk No
Private mPos        As String   'Position No

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get SampleData() As String
    SampleData = mSampleData
End Property

Public Property Let SampleData(ByVal vData As String)
    mSampleData = vData
End Property

Public Property Get Fu() As String
    Fu = mFu
End Property

Public Property Let Fu(ByVal vData As String)
    mFu = vData
End Property

Public Property Get DiskNo() As String
    DiskNo = mDiskNo
End Property

Public Property Let DiskNo(ByVal vData As String)
    mDiskNo = vData
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
'       - pAccInfo : 접수내역 객체
'   반환 : 오더정보 문자열
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim lngIntBase  As Long             '장비기준 검사명
    Dim strItems    As String           '전송할 검사항목
    Dim blnISE      As Boolean          'Na, K, Cl 검사여부

    '## 검사항목 문자열 생성
    strItems = String$(88, "0")
    If pAccInfo.QcFg = "0" Then
        For Each objResult In pAccInfo.Results
            lngIntBase = CLng(objResult.IntNm.IntBase)
            
            '## 계산항목: 93~100
            If lngIntBase >= 93 And lngIntBase <= 100 Then GoTo Skip1
            
            '## Na, K, Cl 검사여부 Check
            If lngIntBase = 87 Or lngIntBase = 88 Or lngIntBase = 89 Then
                blnISE = True
            Else
                Mid$(strItems, lngIntBase, 1) = "1"
            End If
Skip1:
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then
        For Each objQCResult In pAccInfo.QCResults
            lngIntBase = CLng(objQCResult.IntNm.IntBase)
            
            '## 계산항목: 93~100
            If lngIntBase >= 93 And lngIntBase <= 100 Then GoTo Skip2
            
            '## Na, K, Cl 검사여부 Check
            If lngIntBase = 87 Or lngIntBase = 88 Or lngIntBase = 89 Then
                blnISE = True
            Else
                Mid$(strItems, lngIntBase, 1) = "1"
            End If
Skip2:
        Next
        Set objQCResult = Nothing
    End If
    
    '## Na, K, Cl 검사여부 Check
    If blnISE Then
        Mid$(strItems, 87, 1) = "1"
    End If

    GetOrder = STX & ";" & mFu & mSampleData & " 88" & strItems & "000000" & ETX
End Function

'-----------------------------------------------------------------------------'
'   기능 : 접수정보가 없을때 전송할 문자열 조회
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    GetNoOrder = STX & ";" & mFu & mSampleData & Space(15) & " 88" & String$(94, "0") & ETX
End Function

'-----------------------------------------------------------------------------'
'   기능 : 클래스 멤버변수 초기화
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mSampleData = ""
    mFu = ""
    mDiskNo = ""
    mPos = ""
End Sub
