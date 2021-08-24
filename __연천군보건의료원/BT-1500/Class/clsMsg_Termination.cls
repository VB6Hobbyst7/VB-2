VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsg_Termination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mTermination(1)    As String
Private mMSG_TERMIN     As String '지역 복사

Public Property Get MSG_TERMIN() As String
    Dim strTemp     As String
    
    strTemp = Join(mTermination, DLM_F)
    strTemp = "4" & strTemp & Chr(COM_CR) & Chr(COM_ETX)
    
    MSG_TERMIN = Chr(COM_STX) & strTemp & CHK_SUM(strTemp, 2) & Chr(COM_CR) & Chr(COM_LF)
End Property

Private Sub Class_Initialize()
    mTermination(0) = MSG_L
    mTermination(1) = "1"
End Sub
