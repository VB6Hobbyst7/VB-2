VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsg_Token"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Implements clsMsg_Analysis

'속성값을 가질 지역 변수
Private mMTC As String * 1

Private mvarMsg_Token As String '지역 복사

Public Property Let MSG_TOKEN(ByVal vData As String)
    mMTC = Mid(vData, 2, 1)
End Property

Public Property Get MSG_TOKEN() As String
    mvarMsg_Token = GET_MT & MID_TOKEN & Space(10) & Chr(COM_CR) & Chr(COM_LF)
    MSG_TOKEN = Chr(COM_STX) & mvarMsg_Token & Get_CRC(mvarMsg_Token) & Chr(COM_ETX)
End Property

Private Function clsMsg_Analysis_GetMessage() As String
    clsMsg_Analysis_GetMessage = mMTC
End Function

