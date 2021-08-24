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
'   내  용  : AU680 장비 오더정보 클래스
'-----------------------------------------------------------------------------'

Option Explicit

Private mRackNo     As String   'Rack No
Private mCupPos     As String   'Cup Position
Private mBarNo      As String   '바코드번호
Private mSampleNo   As String   'Sample No

Public Property Get RackNo() As String
    RackNo = mRackNo
End Property

Public Property Let RackNo(ByVal vData As String)
    mRackNo = vData
End Property

Public Property Get CupPos() As String
    CupPos = mCupPos
End Property

Public Property Let CupPos(ByVal vData As String)
    mCupPos = vData
End Property

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get SampleNo() As String
    SampleNo = mSampleNo
End Property

Public Property Let SampleNo(ByVal vData As String)
    mSampleNo = vData
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
    Dim objIntNm    As clsIISIntNm      '장비별 검사항목 클래스
    Dim strOrder    As String           '오더정보 문자열
    
    '## 1.0.1: 이상대(2004-12-29)
    '   - LDL-Calculated값이 23->28로 변경되어 프로그램 수정
    '## TP검사중 검체가 Urine이면 검사를 하면 안됨
    If pAccInfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In pAccInfo.Results
            Set objIntNm = objResult.IntNm
            
            '## LDL-C(28)항목은 계산항목이라 오더를 보내면 안됨(CHOL, T.G, HDL-C)
            '## T.P(01)는 검체가 Urine일때 검사를 하면 안됨
            'If objIntNm.IntBase <> "28" And _
                Not (Mid$(pAccInfo.SpcNm, 1, IISSPCLEN) = IISSPCURINE And objIntNm.IntBase = "01") Then
                'strOrder = strOrder & "0" & objIntNm.IntBase & "0"
                strOrder = strOrder & objIntNm.IntBase & "0"
            'End If
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In pAccInfo.QCResults
            Set objIntNm = objQCResult.IntNm
            
            '## LDL-C(28)항목은 계산항목이라 오더를 보내면 안됨(CHOL, T.G, HDL-C)
            '## T.P(01)는 검체가 Urine일때 검사를 하면 안됨
            If objIntNm.IntBase <> "28" And _
                Not (Mid$(pAccInfo.SpcNm, 1, IISSPCLEN) = IISSPCURINE And objIntNm.IntBase = "01") Then
                strOrder = strOrder & objIntNm.IntBase
            End If
        Next
        Set objResult = Nothing
    End If
    Set objIntNm = Nothing
    'Space(20 - Len(mOrder.BarNo))
    GetOrder = STX & "S " & mRackNo & mCupPos & mSampleNo & Space(26 - Len(mBarNo)) & mBarNo & Space(4) & "E" & strOrder & ETX
              'STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(20 - Len(mOrder.BarNo)) & mOrder.BarNo & Space(4) & "E" & strItems & ETX

End Function




'-----------------------------------------------------------------------------'
'   기능 : 오더정보가 없을경우의 문자열 조회
'   반환 : 오더정보가 없을경우의 문자열
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    GetNoOrder = STX & "S " & mRackNo & mCupPos & Format$(mBarNo, String$(SPCLEN, "@")) & mSampleNo & ETX
End Function

'-----------------------------------------------------------------------------'
'   기능 : 클래스 멤버변수 초기화
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mRackNo = ""
    mCupPos = ""
    mBarNo = ""
    mSampleNo = ""
End Sub
