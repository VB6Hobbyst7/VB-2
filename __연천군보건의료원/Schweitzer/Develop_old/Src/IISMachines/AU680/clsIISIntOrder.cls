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
    
    '## TP검사중 검체가 Urine이면 검사를 하면 안됨
    If pAccInfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In pAccInfo.Results
            Set objIntNm = objResult.IntNm
            
            '## LDL-C(027)항목은 계산항목이라 오더를 보내면 안됨(CHOL, T.G, HDL-C)
            '## B/C  (025)항목은 계산항목이라 오더를 보내면 안됨(BUN,CREA)
            '## A/G  (026)항목은 계산항목이라 오더를 보내면 안됨
            '## GLOB (032)항목은 계산항목이라 오더를 보내면 안됨
            '## I.Bil(033)항목은 계산항목이라 오더를 보내면 안됨
            '## T.P  (002)항목은 검체가 Urine일때 검사를 하면 안됨
            
            If objIntNm.IntBase <> "027" And objIntNm.IntBase <> "026" And objIntNm.IntBase <> "025" And _
                objIntNm.IntBase <> "032" And objIntNm.IntBase <> "033" And _
                Not (Mid$(pAccInfo.SpcNm, 1, IISSPCLEN) = IISSPCURINE And objIntNm.IntBase = "002") Then
                                
                '## HbA1C(23)항목은 Hgb(20)와 A1C(21) 오더를 보내야 함
                If objIntNm.IntBase = "023" Then
                    strOrder = strOrder & "020" & "0"
                    strOrder = strOrder & "021" & "0"
                Else
                    strOrder = strOrder & objIntNm.IntBase & "0"
                End If
            End If
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In pAccInfo.QCResults
            Set objIntNm = objQCResult.IntNm
            
            '## LDL-C(027)항목은 계산항목이라 오더를 보내면 안됨(CHOL, T.G, HDL-C)
            '## T.P  (002)는 검체가 Urine일때 검사를 하면 안됨
            If objIntNm.IntBase <> "027" And _
                Not (Mid$(pAccInfo.SpcNm, 1, IISSPCLEN) = IISSPCURINE And objIntNm.IntBase = "002") Then
                strOrder = strOrder & objIntNm.IntBase
            End If
        Next
        Set objResult = Nothing
    End If
    Set objIntNm = Nothing
    
    GetOrder = STX & "S " & mRackNo & mCupPos & mSampleNo & Space(26 - Len(mBarNo)) & mBarNo & Space(4) & "E" & strOrder & ETX

End Function




'-----------------------------------------------------------------------------'
'   기능 : 오더정보가 없을경우의 문자열 조회
'   반환 : 오더정보가 없을경우의 문자열
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    'GetNoOrder = STX & "S " & mRackNo & mCupPos & Format$(mBarNo, String$(SPCLEN, "@")) & mSampleNo & ETX

    GetNoOrder = STX & "S " & mRackNo & mCupPos & mSampleNo & Space(26 - Len(mBarNo)) & mBarNo & Space(4) & "E" & ETX

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
