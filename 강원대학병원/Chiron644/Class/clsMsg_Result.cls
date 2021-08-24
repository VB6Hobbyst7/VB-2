VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsg_Result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Event MsgBegin(ByVal SID As String)
Public Event MsgSplit(ByVal Result As clsResult)
Public Event MsgComplit(ByVal SID As String)

'속성값을 가질 지역 변수
Private mSID            As String
Private mType           As String

Public Property Let MSG_SID(ByVal vData As String)
    mSID = vData
End Property

Public Property Let MSG_TYPE(ByVal vData As String)
    mType = vData
End Property

Public Property Let MSG_RESULT(ByVal strMsg As String)
    
    Dim strRst()        As String
    Dim strAssayNo()    As String
    Dim objResult       As clsResult
    
'-- YEJ
'    Set objResult = New clsResult
'    strRst = Split(strMsg, DLM_F)
'    With objResult
'        .Rst_Eid = mSID
'        .Rst_Sid = mSID
'        strAssayNo = Split(strRst(2), DLM_C)
'        .Rst_Test = Trim(strAssayNo(3))
'        .Rst_Values = Trim(strRst(3))
'        .Rst_Type = mType
'        .Rst_Error = ""
'        .Rst_Tag = "" 'mTYPE
'    End With
'    If strAssayNo(7) = "F" Then
'        RaiseEvent MsgSplit(objResult)
'    End If
'    Set objResult = Nothing
    
    Set objResult = New clsResult
    With objResult
        .Rst_Eid = Trim$(Mid$(strMsg, 3, 29))
        .Rst_Sid = Trim$(Mid$(strMsg, 3, 29))
        .Rst_Test = Format$(Mid$(strMsg, 32, 2), "00")
        .Rst_Values = Mid$(strMsg, 38, 15)
        .Rst_Type = mType
        .Rst_Error = ""
        .Rst_Tag = "" 'mTYPE
    End With
    Set objResult = Nothing

End Property

Private Sub Class_Initialize()
    mSID = ""
    mType = ""
End Sub
