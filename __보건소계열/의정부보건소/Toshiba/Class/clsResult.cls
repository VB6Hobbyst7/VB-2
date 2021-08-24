VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mRst_Eid        As String '장비 번호
Private mRst_Sid        As String '검체 번호
Private mRst_Test       As String '검사 코드
Private mRst_Values     As String '값
Private mRst_Type       As String '검사 타입
Private mRst_Error      As String '에러 코드
Private mRst_Tag        As String '테그

Public Property Get Rst_Eid() As String
    Rst_Eid = mRst_Eid
End Property

Public Property Let Rst_Eid(ByVal vNewValue As String)
    mRst_Eid = vNewValue
End Property

Public Property Get Rst_Sid() As String
    Rst_Sid = mRst_Sid
End Property

Public Property Let Rst_Sid(ByVal newValue As String)
    mRst_Sid = newValue
End Property

Public Property Get Rst_Test() As String
    Rst_Test = mRst_Test
End Property

Public Property Let Rst_Test(ByVal newValue As String)
    mRst_Test = Trim(newValue)
End Property

Public Property Get Rst_Values() As String
    Rst_Values = mRst_Values
End Property

Public Property Let Rst_Values(ByVal newValue As String)
    mRst_Values = Trim(newValue)
End Property

Public Property Get Rst_Type() As String
    Rst_Type = mRst_Type
End Property

Public Property Let Rst_Type(ByVal newValue As String)
    mRst_Type = Trim(newValue)
End Property

Public Property Get Rst_Error() As String
    Rst_Error = mRst_Error
End Property

Public Property Let Rst_Error(ByVal newValue As String)
    mRst_Error = Trim(newValue)
End Property

Public Property Get Rst_Tag() As String
    Rst_Tag = mRst_Tag
End Property

Public Property Let Rst_Tag(ByVal vNewValue As String)
    mRst_Tag = vNewValue
End Property

Private Sub Class_Initialize()
    mRst_Eid = ""       '장비 번호
    mRst_Sid = ""       '검체 번호
    mRst_Test = ""      '검사 코드
    mRst_Values = ""    '값
    mRst_Type = ""      '검사 타입
    mRst_Error = ""     '에러 코드
    mRst_Tag = ""       '테그
End Sub
