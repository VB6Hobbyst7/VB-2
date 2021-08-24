VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISTestCd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISTestCd.cls
'   작성자  : 이상대
'   내  용  : 검사코드 클래스
'   작성일  : 2004-01-19
'   버  전  :
'   메  모  : 검사항목 마스터의 이전, 다음기능을 위해 특별히 만들어진 클래스
'-----------------------------------------------------------------------------'

Option Explicit

Private mTestCd     As String       '검사코드
Private mTestSeq    As Long         'Sequence
Private mTestNm     As String       '검사명

Public Property Get Cd() As String
    Cd = mTestCd
End Property

Public Property Get CdNm() As String
    CdNm = mTestNm
End Property

Public Property Get TestCd() As String
    TestCd = mTestCd
End Property

Public Property Let TestCd(ByVal vData As String)
    mTestCd = vData
End Property

Public Property Get TestSeq() As Long
    TestSeq = mTestSeq
End Property

Public Property Let TestSeq(ByVal vData As Long)
    mTestSeq = vData
End Property

Public Property Get TestNm() As String
    TestNm = mTestNm
End Property

Public Property Let TestNm(ByVal vData As String)
    mTestNm = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 현재 객체와 전달된 객체의 TestCd를 비교
'   인수 :
'       1.pObj : clsIISTestCd 객체
'   반환 :
'       1.현재객체 < 전달객체 이면 -1
'       2.현재개체 = 전달객체 이면 0
'       3.현재개체 > 전달객체 이면 1
'-----------------------------------------------------------------------------'
Public Function Compare(ByVal pObj As Object)
    Compare = StrComp(mTestCd, pObj.TestCd, vbTextCompare)
End Function
