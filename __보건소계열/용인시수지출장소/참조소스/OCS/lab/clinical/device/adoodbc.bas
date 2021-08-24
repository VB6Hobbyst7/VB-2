Attribute VB_Name = "AdoConst"
Option Explicit

Public strSQL               As String
Public Rowindicator         As Long
Public Result               As Boolean

Public GnMousePointer       As Integer

Public adoConnect           As ADODB.Connection
Public Rs                   As ADODB.Recordset

'Public GstrMsgList          As String
'Public GstrMsgTitle         As String
'Public GstrMsgOpt           As Integer
'Public GstrMsgRet           As Integer

Public GstrPassProgramID    As String * 8
Public GstrPassWord         As String
Public GstrPassGrade        As String
Public GstrPassClass        As String
Public GstrPassName         As String
Public GstrPassPart         As String * 2
Public GstrPassDept         As String
Public GstrIdnumber         As String
Public GstrPassIDnumber     As String


Public Sub DbAdoDisConnect()

    On Error Resume Next
    adoConnect.Close
    If Not adoConnect Is Nothing Then
        Set adoConnect = Nothing
    End If
    
End Sub


Public Function DbAdoConnect(ByVal ArgUser$, ByVal ArgPassword$, ByVal ArgSource$) As Integer
    
    Dim sConString          As String
    
    sConString = ""
    sConString = sConString & "Provider=Microsoft OLE DB Provider for Oracle;"
    sConString = sConString & "Data Source=" & ArgSource & ";"
    sConString = sConString & "Persist Security Info=False"
    
    On Error GoTo DBConnect_Error
    
    GnMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    Set adoConnect = New ADODB.Connection
    
    adoConnect.CursorLocation = adUseClient
    adoConnect.Open sConString, ArgUser, ArgPassword
    
    Screen.MousePointer = GnMousePointer
    
    Exit Function
    
'/-----------------------------------------------------------------------------/

DBConnect_Error:
    
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description & vbCrLf & vbCrLf & _
           "ConnectString : " & sConString & vbCrLf & _
           "User          : " & ArgUser & vbCrLf & _
           "Password      : " & ArgPassword
    
    End
    
End Function

Public Function AdoExecute(ByVal SQL As String) As Integer
    
    GnMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ExecError:
    
    AdoExecute = 0
'    AdoExecute = False
    Rowindicator = 0
    
    adoConnect.Execute SQL, Rowindicator, adCmdText + ADODB.adExecuteNoRecords
    Screen.MousePointer = GnMousePointer
    AdoExecute = True
    
Exit Function

ExecError:
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description & vbCrLf & _
           SQL
''    AdoExecute = -1
    AdoExecute = False
    Rowindicator = -1
    
    Screen.MousePointer = GnMousePointer
    
End Function

Public Function AdoExecute1(ByVal SQL As String) As Integer

    On Error GoTo ExecError:
    AdoExecute1 = 0
    Rowindicator = 0
    
    adoConnect.Execute SQL, Rowindicator, adCmdText
    
Exit Function

ExecError:
    
''    AdoExecute1 = -1
    AdoExecute1 = False
    Rowindicator = -1
    
End Function

Public Function AdoOpenSet(ByRef sAdoset As ADODB.Recordset, ByVal SQL As String, Optional ByVal nRowCnt As Boolean = True, Optional ByVal nMousePointer = True) As Integer
    
    Set sAdoset = New ADODB.Recordset
    
    If nMousePointer = True Then
        GnMousePointer = Screen.MousePointer
        Screen.MousePointer = vbHourglass
    End If
    On Error GoTo OpenError:
    
    AdoOpenSet = 0
    Rowindicator = 0
    
    'Set sAdoset = adoConnect.Execute(SQL, Rowindicator, adCmdText)
    Call sAdoset.Open(SQL, adoConnect, adOpenStatic, adLockReadOnly, adCmdText)
    
    If Not sAdoset.EOF Then
        If nRowCnt = True Then
            Rowindicator = sAdoset.RecordCount
            AdoOpenSet = True
        Else
            Rowindicator = -1
        End If
    End If
    
    If nMousePointer = True Then
        Screen.MousePointer = GnMousePointer
    End If
    
    Exit Function
            
OpenError:
    
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description & vbCrLf & _
           SQL
''    AdoOpenSet = -1
    AdoOpenSet = False
    
    Screen.MousePointer = GnMousePointer
        
End Function

Public Sub AdoCloseSet(ByRef sAdoset As ADODB.Recordset)

    On Error GoTo SetClose_Error
    sAdoset.Close
    If Not sAdoset Is Nothing Then Set sAdoset = Nothing
    
    Exit Sub
    
SetClose_Error:
    
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description
    
End Sub

Public Function AdoGetString(ByRef adoS As ADODB.Recordset, ByVal adoCol As String, Optional ByVal AbsPos As Long = -1) As String

    On Error GoTo ReadError

    If AbsPos > -1 Then adoS.AbsolutePosition = AbsPos + 1
    AdoGetString = adoS.Fields(adoCol).Value & ""
    
    Exit Function
    
ReadError:
    
    AdoGetString = ""
    
End Function

Public Function AdoGetNumber(ByRef adoS As ADODB.Recordset, ByVal adoCol As String, Optional ByVal AbsPos As Long = 1) As Double

    On Error GoTo ReadError

    If AbsPos > -1 Then adoS.AbsolutePosition = AbsPos + 1
    AdoGetNumber = Val(adoS.Fields(adoCol).Value & "")
    
    Exit Function
    
ReadError:
    
    AdoGetNumber = 0
    
End Function


Public Function Quot(ByVal strString As String) As String
    Dim i                   As Integer
    Dim nPos                As Integer
    nPos = 1
    Do
        For i = nPos To Len(strString)
            If Mid(strString, i, 1) = "'" Then
                strString = Left(strString, i - 1) & "''" & Mid(strString, i + 1)
                Exit For
            End If
        Next i
        nPos = i + 2
        If nPos > Len(strString) Then Exit Do
    Loop While (True)
    Quot = strString
    
End Function
