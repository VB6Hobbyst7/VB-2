Attribute VB_Name = "modInstrument"
Option Explicit

Global Const REG_INSNAME    As String = "coaguChek"

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

