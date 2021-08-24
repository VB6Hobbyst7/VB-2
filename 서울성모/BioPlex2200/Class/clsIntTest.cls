VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIntTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mMachCd     As String       ' 장비코드
Dim mTESTCD     As String       ' 검사코드
Dim mTestNm     As String       ' 검사명
Dim mMcTestCd   As String       ' 장비검사코드
Dim mMcTestNm   As String       ' 장비검사명
Dim mFrval      As String       ' Ref상한
Dim mToval      As String       ' Ref하한

Private mvarDictionary As New Scripting.Dictionary
Private mvarRecordCount As Long


Public Property Get Testcd() As Variant
    Testcd = mTESTCD
End Property

Public Property Let Testcd(ByVal vNewValue As Variant)
    mTESTCD = vNewValue
End Property

Public Property Get TestNm() As Variant
    TestNm = mTestNm
End Property

Public Property Let TestNm(ByVal vNewValue As Variant)
    mTestNm = vNewValue
End Property

Public Property Get McTestCd() As Variant
    McTestCd = mMcTestCd
End Property

Public Property Let McTestCd(ByVal vNewValue As Variant)
    mMcTestCd = vNewValue
End Property

Public Property Get McTestNm() As Variant
    McTestNm = mMcTestNm
End Property

Public Property Let McTestNm(ByVal vNewValue As Variant)
    mMcTestNm = vNewValue
End Property

Public Property Get FrVal() As Variant
    FrVal = mFrval
End Property

Public Property Let FrVal(ByVal vNewValue As Variant)
    mFrval = vNewValue
End Property

Public Property Get ToVal() As Variant
    ToVal = mToval
End Property

Public Property Let ToVal(ByVal vNewValue As Variant)
    mToval = vNewValue
End Property

Public Function Exists(ByVal Key As String) As Boolean
    
    Exists = mvarDictionary.Exists(Key)
    
End Function

Public Function AddNew(ByVal Key As String, ByVal Data As String, Optional ByVal Tags As String)
    Dim aryTmp() As String
    Dim aryKey() As String
    Dim aryTag() As String
    
    
    aryKey = Split(Key, "|")
    aryTmp = Split(Data, "|")
    aryTag = Split(Tags, "|")

    mvarDictionary.Add Join(aryKey, "|"), Join(aryTmp, "|")
    mvarRecordCount = mvarDictionary.Count
    
End Function


'-----------------------------------------------------------------------------'
'   기능 : 장비기준 검사명을 이용해 검사명 조회
'   인수 :
'       - pIntBase : 장비기준 검사명
'   반환 : 검사명
'-----------------------------------------------------------------------------'
Public Function GetIntNm(ByVal pIntBase As String) As String
    
    GetIntNm = mvarDictionary(pIntBase)

End Function

