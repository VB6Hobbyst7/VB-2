VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'속성값을 가질  변수
Private mvarErr As Long
Private mvarErrText As String

Public Sub Reset()
   glngErrorNo = 0
   gstrErrorMsg = ""
   mvarErr = 0
   mvarErrText = ""
End Sub

Public Property Get ErrText() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.ErrText
   Select Case glngErrorNo
      Case 457
         mvarErrText = "데이터 중복 Error." & vbNewLine & "(" & gstrErrorMsg & ")"
      Case 90001
         mvarErrText = "데이터 중복 Error." & vbNewLine & "(" & gstrErrorMsg & ")"
      Case Else
         mvarErrText = gstrErrorMsg
   End Select
   '
   ErrText = mvarErrText
End Property

Public Property Get Err() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.Err
   mvarErr = glngErrorNo
   Err = mvarErr
End Property
