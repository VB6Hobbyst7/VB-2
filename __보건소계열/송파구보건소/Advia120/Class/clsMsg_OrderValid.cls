VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsg_OrderValid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements clsMsg_Analysis

Private Const DM_MV As String * 2 = " 0"
Private Const DM_MI As String * 2 = " 4"
Private Const QM_MV As String * 2 = "10"
Private Const QM_MI As String * 2 = "14"

Private MTC As String
Private CDF As String

'속성값을 가질 지역 변수
Private mvarMSG_VALIDOREDER As String '지역 복사

Public Property Let MSG_VALIDOREDER(ByVal vData As String)
    MTC = Mid(vData, 2, 1)
    CDF = Mid(vData, 12, 2)
End Property

Public Property Get MSG_VALIDOREDER() As String

    mvarMSG_VALIDOREDER = GET_MT & MID_VALIDO & Chr(COM_SP) & QM_MV & Chr(COM_CR) & Chr(COM_LF)
    MSG_VALIDOREDER = Chr(COM_STX) & mvarMSG_VALIDOREDER & Get_CRC(mvarMSG_VALIDOREDER) & Chr(COM_ETX)

End Property

Private Function clsMsg_Analysis_GetMessage() As String
    clsMsg_Analysis_GetMessage = MTC
End Function
