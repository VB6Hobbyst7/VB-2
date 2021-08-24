VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsg_OrderNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarMSG_ORDERNO As String '지역 복사
Private mSID            As String * 14

Public Property Let MSG_ORDERNO(ByVal vData As String)
Attribute MSG_ORDERNO.VB_Description = "오더없음"
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.MSG_ORDERNO = 5
    mSID = vData
    mvarMSG_ORDERNO = GET_MT & MID_ORDERN & Chr(COM_SP) & "W" & Chr(COM_SP) & mSID & Chr(COM_CR) & Chr(COM_LF)
    mvarMSG_ORDERNO = Chr(COM_STX) & mvarMSG_ORDERNO & Get_CRC(mvarMSG_ORDERNO) & Chr(COM_ETX)
    
End Property


Public Property Get MSG_ORDERNO() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.MSG_ORDERNO
    MSG_ORDERNO = mvarMSG_ORDERNO
End Property



