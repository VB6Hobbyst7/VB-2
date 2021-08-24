Attribute VB_Name = "modInstrument"
Option Explicit

Global Const REG_INSNAME    As String = "Advia120"

Public Const MSG_H      As String = "H" 'Header
Public Const MSG_P      As String = "P" 'Patient Informaiton
Public Const MSG_Q      As String = "Q" 'Request Information
Public Const MSG_O      As String = "O" 'Test Order
Public Const MSG_R      As String = "R" 'Result
Public Const MSG_C      As String = "C" 'Comment
Public Const MSG_M      As String = "M" 'Manufacturer Information
Public Const MSG_S      As String = "S" 'Scientific
Public Const MSG_L      As String = "L" 'Message Terminator

Public Const DLM_F      As String = "|" 'Separates Field
Public Const DLM_R      As String = "\" 'Separates Repeat
Public Const DLM_C      As String = "^" 'Separates Component
Public Const DLM_E      As String = "&" 'Escape Delimiter

Public CU_STATUS        As String * 1

Public COM_MODE         As String             '// 0 = 출력안함, 1 = 출력

Public Const MID_INIT   As String = "I"
Public Const MID_TOKEN  As String = "S"
Public Const MID_ORDER  As String = "Y"
Public Const MID_VALIDO As String = "E"
Public Const MID_RESULT As String = "R"
Public Const MID_QUERY  As String = "Q"
Public Const MID_ORDERN As String = "N"
Public Const MID_VALIDR As String = "Z"

Private Const MT_SRT As Long = &H30
Private Const MT_END As Long = &H5A

Public CHR_MT       As String
Public IS_INIT      As Boolean
Public IS_ORDER     As Boolean

Public Enum CU_MODE
    MOD_INIT = 1
    MOD_ORDER = 2
    MOD_RESULT = 3
End Enum

Public CURRENT_MOD As CU_MODE
Private TimerID         As Long

' 데이타 받기
Public Function COM_INPUT(ByVal strRec As String)
    
    Dim strCnvString        As String
    
'    If COM_MODE = "1" Then
        With frmComm.txtCom
            If .Text <> "" Then .Text = .Text & vbNewLine
            strCnvString = comChar_Convert(strRec)
            
            .Text = .Text & "장비 : " & strCnvString & vbNewLine
        End With
'    End If

End Function

' 데이타 받기2
Public Function COM_INPUT_String(ByVal strRec As String)
    Dim strCnvString        As String
    
    If COM_MODE = "1" Then
        With frmComm.txtCom
            strCnvString = comChar_Convert(strRec)
            .Text = .Text & _
                    String(Len(strCnvString), "-") & vbNewLine & _
                    strCnvString & vbNewLine & _
                    String(Len(strCnvString), "-") & vbNewLine & vbNewLine
        End With
    End If
End Function

' 데이타 내보내기
Public Function Com_Output(ByVal DataValue As String)
    
    Dim strCnvData      As String
    
    If COM_MODE = "1" Then
        With frmComm.txtCom
            .Text = .Text & "HOST : " & comChar_Convert(DataValue) & vbNewLine
        End With
    End If
    frmComm.comEQP.Output = DataValue
        
End Function

' 장비에 따른 통신문자 변환
Public Function comChar_Convert(ByVal OrgString As String) As String
    Dim NewString   As String
    
    OrgString = Replace(OrgString, Chr(COM_SOH), "{SOH}")
    OrgString = Replace(OrgString, Chr(COM_NUL), "{NUL}")
    OrgString = Replace(OrgString, Chr(COM_STX), "{STX}")
    OrgString = Replace(OrgString, Chr(COM_ETX), "{ETX}")
    OrgString = Replace(OrgString, Chr(COM_ACK), "{ACK}")
    OrgString = Replace(OrgString, Chr(COM_NACK), "{NAK}")
    OrgString = Replace(OrgString, Chr(COM_ENQ), "{ENQ}")
    OrgString = Replace(OrgString, Chr(COM_EOT), "{EOT}")
    OrgString = Replace(OrgString, Chr(COM_CR), "{CR}")
    NewString = Replace(OrgString, Chr(COM_LF), "{LF}")
    
    comChar_Convert = NewString
End Function

' 장비에 따른 통신문자 변환
Public Function charCOM_Convert(ByVal OrgString As String) As String
    Dim NewString   As String
    
    OrgString = Replace(OrgString, "{NUL}", Chr(COM_NUL))
    OrgString = Replace(OrgString, "{STX}", Chr(COM_STX))
    OrgString = Replace(OrgString, "{ETX}", Chr(COM_ETX))
    OrgString = Replace(OrgString, "{ACK}", Chr(COM_ACK))
    OrgString = Replace(OrgString, "{NAK}", Chr(COM_NACK))
    OrgString = Replace(OrgString, "{ENQ}", Chr(COM_ENQ))
    OrgString = Replace(OrgString, "{EOT}", Chr(COM_EOT))
    OrgString = Replace(OrgString, "{CR}", Chr(COM_CR))
    NewString = Replace(OrgString, "{LF}", Chr(COM_LF))
    
    charCOM_Convert = NewString
End Function

Public Function GET_MT() As String
'    Static lngMT As Long
'
'    Select Case lngMT
'        Case MT_END
'            lngMT = MT_SRT
'        Case Chr(MT_SRT)
'            lngMT = MT_SRT
'        Case Else
'            lngMT = lngMT + 1
'    End Select
'
'    GET_MT = lngMT

    Select Case Asc(CHR_MT)
        Case MT_END
            GET_MT = Chr(MT_SRT)
        Case Else
            GET_MT = Chr(Asc(CHR_MT) + 1)
    End Select
End Function

'Public Sub Com_Output(ByVal Msg_Send As String)
'    frmComm.comEQP.Output = Msg_Send
'End Sub

Public Sub Send_Token()
    Call SetTimer(frmComm.hwnd, TimerID, 5000, AddressOf Token_Proc)
End Sub

Public Sub Stop_Token()
    Call KillTimer(frmComm.hwnd, TimerID)
End Sub

Public Sub Token_Proc(ByVal hwnd&, ByVal Msg&, ByVal ID&, ByVal nTime&)
    Dim objToken As clsMsg_Token
    
    Set objToken = New clsMsg_Token
    Call Com_Output(objToken.MSG_TOKEN)
    Set objToken = Nothing

End Sub
