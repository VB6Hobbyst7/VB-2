Attribute VB_Name = "modAXSYM"
Option Explicit

Private TimerID         As Long

Public Sub Com_Output(ByVal Msg_Send As String)
'    frmInterface.comADVIA120.Output = Msg_Send
End Sub

Public Sub Send_Token()
'    Call SetTimer(frmInterface.hwnd, TimerID, 5000, AddressOf Token_Proc)
End Sub

Public Sub Stop_Token()
'    Call KillTimer(frmInterface.hwnd, TimerID)
End Sub

Public Sub Token_Proc(ByVal hwnd&, ByVal Msg&, ByVal ID&, ByVal nTime&)
'    Dim objToken As clsMsg_Token
'
'    Set objToken = New clsMsg_Token
'    frmInterface.comADVIA120.Output = objToken.MSG_TOKEN
'    Set objToken = Nothing

End Sub
