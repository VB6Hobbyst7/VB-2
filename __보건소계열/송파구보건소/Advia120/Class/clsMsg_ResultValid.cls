VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsg_ResultValid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Implements clsMsg_Analysis
Private Const ECD As String * 2 = " 0"

Private mvarMSG_VALIDRESULT As String

Public Property Let MSG_VALIDRESULT(ByVal vData As String)

End Property

Public Property Get MSG_VALIDRESULT() As String

    mvarMSG_VALIDRESULT = GET_MT & MID_VALIDR & Space(17) & ECD & Chr(COM_CR) & Chr(COM_LF)
    MSG_VALIDRESULT = Chr(COM_STX) & mvarMSG_VALIDRESULT & Get_CRC(mvarMSG_VALIDRESULT) & Chr(COM_ETX)

End Property
