VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsg_Init"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Implements clsMsg_Analysis

'속성값을 가질 지역 변수
Private mvarMsg_Init As String '지역 복사

Public Property Get MSG_INIT() As String
    mvarMsg_Init = Chr(&H30) & MID_INIT & Chr(COM_SP) & Chr(COM_CR) & Chr(COM_LF)
    MSG_INIT = Chr(COM_STX) & mvarMsg_Init & Get_CRC(mvarMsg_Init) & Chr(COM_ETX)
End Property

Private Function clsMsg_Analysis_GetMessage() As String
    clsMsg_Analysis_GetMessage = MSG_INIT
End Function
