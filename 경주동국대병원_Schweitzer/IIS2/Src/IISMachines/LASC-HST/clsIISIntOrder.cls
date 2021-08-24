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
'   내  용  : LASC-HST 오더전송 클래스
'   작성일  : 2005-09-15
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo  As String           '바코드번호
Private mRow    As Long             'Spread의 Row

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get Row() As Long
    Row = mRow
End Property

Public Property Let Row(ByVal vData As Long)
    mRow = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 문자열 조회
'   인수 :
'       - pAccInfo : 접수내역 클래스
'   반환 : 오더정보 문자열1
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult       As clsIISResult     '결과내역 클래스
    Dim strIntBase      As String           '장비기준 검사명
    Dim strItems        As String           '송신할 검사항목 문자열
    Dim strOutput       As String           '송신할 데이터
    Dim strPrintNo1     As String
    Dim strPrintNo2     As String
    Dim strPrintNo3     As String
    Dim blnSPOrder      As Boolean

    strItems = String$(38, "0")
    For Each objResult In pAccInfo.Results
        strIntBase = mGetP(objResult.IntNm.IntBase, 1, "|")
        
        Select Case strIntBase
            Case "RET"
                Mid$(strItems, 1, 1) = 1
            Case "SP"
                Mid$(strItems, 2, 1) = 1
                blnSPOrder = True
            Case Else
                Mid$(strItems, CLng(strIntBase) + 3, 1) = 1
        End Select
    Next
    Set objResult = Nothing
    
    If blnSPOrder Then
        Mid$(strItems, 3, 1) = "0"
    Else
        Mid$(strItems, 3, 1) = "1"
    End If
    
    mBarNo = Format$(mBarNo, String$(13, "0"))
    strPrintNo1 = Format$(pAccInfo.AccDt & "-" & CStr(pAccInfo.AccSeq), String$(13, "@"))
    strPrintNo2 = Space(13)
    strPrintNo3 = Space(13)
    
    strOutput = STX & "S" & Space(8) & String$(2, "0") & mBarNo & "********" & strItems & _
                String$(19, "0") & strPrintNo1 & strPrintNo2 & strPrintNo3 & "100" & _
                String$(25, "*") & Space(15) & ETX
    
    GetOrder = strOutput
End Function

''-----------------------------------------------------------------------------'
'   기능 : 클래스 멤버변수 초기화
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mRow = 0
End Sub
