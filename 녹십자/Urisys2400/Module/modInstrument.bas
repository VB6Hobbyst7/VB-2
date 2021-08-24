Attribute VB_Name = "modInstrument"
Option Explicit

Global Const REG_INSNAME    As String = "Urisys2400"

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


'
'   ASTM Protocol CheckSum 계산
'
Public Function ChkSum_ASTM(ByVal Para As String) As String

    Dim I   As Integer
    Dim Tmp As Integer
    Dim ChkS1   As Integer
    Dim ChkS2   As String
    
    For I = 1 To Len(Para)
        Tmp = Asc(Mid$(Para, I, 1))
        ChkS1 = ChkS1 + Tmp
    Next I
    ChkS1 = ChkS1 Mod 256
    ChkS2 = Right$("0" & Hex$(ChkS1), 2)
    
    ChkSum_ASTM = ChkS2
    
End Function

' 데이타 받기
Public Function COM_INPUT(ByVal strRec As String)
    
    Dim strCnvString        As String
    
    If COM_MODE = "1" Then
        With frmComm.txtCom
            If .text <> "" Then .text = .text & vbNewLine
            strCnvString = comChar_Convert(strRec)
            
            .text = .text & "장비 : " & strCnvString & vbNewLine
        End With
    End If

End Function

' 데이타 받기2
Public Function COM_INPUT_String(ByVal strRec As String)
    Dim strCnvString        As String
    
    If COM_MODE = "1" Then
        With frmComm.txtCom
            strCnvString = comChar_Convert(strRec)
            .text = .text & _
                    String(Len(strCnvString), "-") & vbNewLine & _
                    strCnvString & vbNewLine & _
                    String(Len(strCnvString), "-") & vbNewLine & vbNewLine
        End With
    End If
End Function

' 데이타 내보내기
Public Function COM_OUTPUT(ByVal DataValue As String)
    
    Dim strCnvData      As String
    
    If COM_MODE = "1" Then
        With frmComm.txtCom
            .text = .text & "HOST : " & comChar_Convert(DataValue) & vbNewLine
        End With
    End If
    frmComm.comEQP.Output = DataValue
        
End Function

Public Sub SaveLog(ByVal strdata As String)
On Error Resume Next

    Dim objFile     As FileSystemObject
    Dim logFile     As TextStream
    Dim FileName    As String
    
    Set objFile = New FileSystemObject
    
    FileName = Format(Now, "YYYYMMDD") & ".LOG"
    With objFile
        If Not .FolderExists(DirPath & "LOG\") Then
            Call .CreateFolder(DirPath & "LOG\")
        End If
        Set logFile = .OpenTextFile(DirPath & "LOG\" & FileName, ForAppending, True)
    End With
    
    Call logFile.Write(strdata)
    Call logFile.Close
    
    Set objFile = Nothing
    Set logFile = Nothing
End Sub


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

