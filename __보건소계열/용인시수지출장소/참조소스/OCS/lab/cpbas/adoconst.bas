Attribute VB_Name = "adoCnst"
Option Explicit

'Public adoConnect       As ADODB.Connection
Public strSql               As String
Public rowindicator         As Long
Public Result               As Integer
Public GnMousePointer       As Integer

Public adoConnect           As ADODB.Connection
Public rs                   As ADODB.Recordset

Public GstrMsgList          As String
Public GstrMsgTitle         As String
Public GstrMsgOpt           As Integer
Public GstrMsgRet           As Integer

Public GstrSysDate          As String
Public GstrPassProgramID    As String * 8
Public GstrPassWord         As String
Public GstrPassGrade        As String
Public GstrPassClass        As String
Public GstrPassName         As String
Public GstrPassPart         As String * 2
Public GstrPassDept         As String
Public GstrIdnumber         As String
Public GstrPassIDnumber     As String
Public GstrPassId           As String

Public adoSet           As ADODB.Recordset
Public lngExeCount      As Long


Public Function adoDbConnect(ByVal sUser As String, ByVal sPassword As String, ByVal sDataSRC As String) As Integer
    Dim sConString          As String
    
        
    sConString = ""
    sConString = sConString & "Provider=Microsoft OLE DB Provider for Oracle" & ";"
    sConString = sConString & "User ID=" & sUser & ";"
    sConString = sConString & "Data Source=" & sDataSRC & ";"
    sConString = sConString & "Persist Security info=False"
    
    
    On Error GoTo DBConnect_Error
    
    Set adoConnect = New ADODB.Connection
    adoConnect.CursorLocation = adUseClient
    adoConnect.Open sConString, sUser, sPassword
    
    Exit Function
    
    
DBConnect_Error:
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description & vbCrLf & vbCrLf & _
           "ConnectString : " & sConString & vbCrLf & _
           "Username : " & sUser & vbCrLf & _
           "Password : " & sPassword
    End
    Return

End Function

Public Function adoDbDisconnect() As Integer
    
    adoConnect.Close
    If Not adoConnect Is Nothing Then
        Set adoConnect = Nothing
    End If
    
End Function
Public Function adoSetOpen(ByVal sSql As String, ByRef sAdoset As ADODB.Recordset) As Integer
    
    On Error GoTo SetOpen_Error
    
    Set sAdoset = New ADODB.Recordset
    
    'Set sAdoset = adoConnect.Execute(sSql)
    
    Call sAdoset.Open(sSql, adoConnect, adOpenStatic, adLockReadOnly, adCmdText)
    
    If sAdoset.EOF Then
        adoSetOpen = False
        Exit Function
    End If
    
    If sAdoset.RecordCount = 0 Then
        adoSetOpen = False
    Else
        adoSetOpen = True
    End If
    
    Exit Function
    
    
SetOpen_Error:
    adoSetOpen = False
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description & vbCrLf & _
           sSql
           
    Exit Function
    
    Return
    
End Function

Public Function adoExec(ByVal sSql As String, Optional nRetCount As Integer) As Integer
    
    
    On Error GoTo SetOpen_Error
    
    adoExec = True
    Call adoConnect.Execute(sSql, nRetCount, adCmdText + ADODB.adExecuteNoRecords)
    adoExec = True
    Exit Function
    
SetOpen_Error:
    adoExec = False
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description & vbCrLf & _
           sSql
    Exit Function
    Return
    
End Function


Public Function adoSetClose(ByRef sAdoset As ADODB.Recordset) As Integer
    
    On Error GoTo SetClose_Error
    
    sAdoset.Close
    If Not sAdoset Is Nothing Then Set sAdoset = Nothing
    
    Exit Function
    
    
SetClose_Error:
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description
    adoSetClose = False
    
    Exit Function
    Return

End Function

Public Sub DbAdoDisConnect()

    On Error Resume Next
    adoConnect.Close
    If Not adoConnect Is Nothing Then
        Set adoConnect = Nothing
    End If
    
End Sub

Public Function DbAdoConnect(ByVal ArgUser$, ByVal ArgPassword$, ByVal ArgSource$) As Integer
    
    Dim sConString      As String
    
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
    rowindicator = 0
    
    adoConnect.Execute SQL, rowindicator, adCmdText + ADODB.adExecuteNoRecords
    Screen.MousePointer = GnMousePointer
    
Exit Function

ExecError:
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description & vbCrLf & _
           SQL
    AdoExecute = -1
    rowindicator = -1
    
    Screen.MousePointer = GnMousePointer
    
End Function

Public Function AdoExecute1(ByVal SQL As String) As Integer

    On Error GoTo ExecError:
    AdoExecute1 = 0
    rowindicator = 0
    
    adoConnect.Execute SQL, rowindicator, adCmdText
    
Exit Function

ExecError:
    
    AdoExecute1 = -1
    rowindicator = -1
    
End Function

Public Function AdoOpenSet(ByRef sAdoset As ADODB.Recordset, ByVal SQL As String, Optional ByVal nRowCnt As Boolean = True, Optional ByVal nMousePointer = True) As Integer
    
    Set sAdoset = New ADODB.Recordset
    
    If nMousePointer = True Then
        GnMousePointer = Screen.MousePointer
        Screen.MousePointer = vbHourglass
    End If
    On Error GoTo OpenError:
    
    AdoOpenSet = 0
    rowindicator = 0
    
    If nRowCnt = True Then
        adoConnect.CursorLocation = adUseClient
    Else
        adoConnect.CursorLocation = adUseServer
    End If
    
    'Set sAdoset = adoConnect.Execute(SQL, Rowindicator, adCmdText)
    Call sAdoset.Open(SQL, adoConnect, adOpenStatic, adLockReadOnly, adCmdText)
    If Not sAdoset.EOF Then
        If nRowCnt = True Then
            rowindicator = sAdoset.RecordCount
        Else
            rowindicator = -1
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
    AdoOpenSet = -1
    
    Screen.MousePointer = GnMousePointer
        
End Function

Public Function Quot(ByVal strString As String) As String

    Dim i       As Integer
    Dim nPos    As Integer
    
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

Public Function AdoGetString(ByRef adoS As ADODB.Recordset, ByVal adoCol As String, Optional ByVal AbsPos As Long = -1) As String

    On Error GoTo ReadError

    If AbsPos > -1 Then adoS.AbsolutePosition = AbsPos + 1
    AdoGetString = adoS.Fields(adoCol).Value & ""
    
    Exit Function
    
ReadError:
    
    AdoGetString = ""
    
End Function

Public Function AdoGetNumber(ByRef adoS As ADODB.Recordset, ByVal adoCol As String, Optional ByVal AbsPos As Long = -1) As Double

    On Error GoTo ReadError

    If AbsPos > -1 Then adoS.AbsolutePosition = AbsPos + 1
    AdoGetNumber = Val(adoS.Fields(adoCol).Value & "")
    
    Exit Function
    
ReadError:
    
    AdoGetNumber = 0
    
End Function
