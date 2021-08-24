Attribute VB_Name = "VbODBC"
Option Explicit

Public strSQL               As String
Public Rowindicator         As Long
Public Result               As Integer

Public GnMousePointer       As Integer

Public RdoEnv               As rdoEnvironment
Public RdoDb                As rdoConnection

Public GstrMsgList          As String
Public GstrMsgTitle         As String
Public GstrMsgOpt           As Integer
Public GstrMsgRet           As Integer

Public Sub DbRdoDisConnect()

    On Error Resume Next
    
    RdoDb.Close
    
    RdoEnv.Close
    
End Sub

Public Sub DbRdoConnect(ByVal ArgUser$, ByVal ArgPassword$, ByVal ArgConnectString$)
    
    On Error GoTo Error_Process
    
    GnMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    Call Create_DSN_ORACLE(ArgConnectString$, ArgConnectString$)
    
    Set RdoEnv = rdoEngine.rdoEnvironments(0)
    
    RdoEnv.UserName = Trim(ArgUser)
    RdoEnv.Password = Trim(ArgPassword)
    RdoEnv.CursorDriver = rdUseOdbc
    
    Set RdoDb = RdoEnv.OpenConnection(Trim(ArgConnectString), rdDriverNoPrompt, False)
    
    Screen.MousePointer = GnMousePointer
    
    Exit Sub
    
'/-----------------------------------------------------------------------------/

Error_Process:
    
    GstrMsgList = "DB Connection(RDO)을 하지 못했습니다" & Chr(13) & Chr(13)
    GstrMsgList = GstrMsgList & "Error Number : " & Err.Number & Chr$(13)
    GstrMsgList = GstrMsgList & "Description  : " & Err.Description
    
    MsgBox GstrMsgList, vbCritical, "DB Connect Error"
    
    End
    
End Sub

Public Function RdoExecute(ByVal SQL As String) As Integer
    
    GnMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ExecError:
    
    RdoExecute = 0
    Rowindicator = 0
    
    RdoDb.Execute SQL
    Rowindicator = RdoDb.RowsAffected ' 한번의 Transaction 에 의해 처리된 행의 숫자
    
    Screen.MousePointer = GnMousePointer
    
Exit Function

ExecError:

    Dim I       As Integer
    
    If Err.Number = 40002 Then  '응용 프로그램 정의 오류 또는 개체 정의 오류
        GstrMsgList = rdoErrors(0).Description
        I = InStr(1, GstrMsgList, "ORA-", vbBinaryCompare)
        If I <= 0 Then I = 1
        GstrMsgList = Mid(GstrMsgList, I, Len(GstrMsgList))
        GstrMsgList = GstrMsgList & vbCrLf & vbCrLf & SQL
        GstrMsgTitle = "Error - " & Trim(Str(rdoErrors(0).Number))
        MsgBox GstrMsgList, vbCritical, GstrMsgTitle
    Else
        MsgBox Err.Description, vbCritical, "VB Error - " & Trim(Str(Err.Number))
    End If
    
    RdoExecute = -1
    
    Screen.MousePointer = GnMousePointer
    
End Function

Public Function RdoExecute1(ByVal SQL As String) As Integer
    
    On Error GoTo ExecError:
    
    RdoExecute1 = 0
    Rowindicator = 0
    
    RdoDb.Execute SQL
    Rowindicator = RdoDb.RowsAffected ' 한번의 Transaction 에 의해 처리된 행의 숫자
    
Exit Function

ExecError:

    Dim I       As Integer
    
    If Err.Number = 40002 Then  '응용 프로그램 정의 오류 또는 개체 정의 오류
        GstrMsgList = rdoErrors(0).Description
        I = InStr(1, GstrMsgList, "ORA-", vbBinaryCompare)
        If I <= 0 Then I = 1
        GstrMsgList = Mid(GstrMsgList, I, Len(GstrMsgList))
        GstrMsgList = GstrMsgList & vbCrLf & vbCrLf & SQL
        GstrMsgTitle = "Error - " & Trim(Str(rdoErrors(0).Number))
        Debug.Print "RdoExecute Error - " & GstrMsgTitle, GstrMsgList,
    Else
        Debug.Print "RdoExecute Error - " & Trim(Str(Err.Number)), Err.Description
    End If
    
    RdoExecute1 = -1
    
End Function

Public Function RdoOpenSet(ByRef Rs As rdoResultset, ByVal SQL As String, Optional ByVal nRowCnt As Boolean = True, Optional ByVal nMousePointer = True) As Integer
    
    If nMousePointer = True Then
        GnMousePointer = Screen.MousePointer
        Screen.MousePointer = vbHourglass
    End If
    
    On Error GoTo OpenError:
    
    RdoOpenSet = 0
    Rowindicator = 0
    
    Set Rs = RdoDb.OpenResultset(SQL, rdOpenStatic, rdConcurReadOnly)
    
    If Not Rs.EOF Then
        If nRowCnt = True Then
            Rowindicator = Rs.RowCount
        Else
            Rowindicator = -1
        End If
    End If
    
    If nMousePointer = True Then
        Screen.MousePointer = GnMousePointer
    End If
    
    Exit Function
    
OpenError:
    
    Dim I   As Integer
    
    If Err.Number = 40002 Then  '응용 프로그램 정의 오류 또는 개체 정의 오류
        GstrMsgList = rdoErrors(0).Description
        I = InStr(1, GstrMsgList, "ORA-", vbBinaryCompare)
        If I <= 0 Then I = 1
        GstrMsgList = Mid(GstrMsgList, I, Len(GstrMsgList))
        GstrMsgList = GstrMsgList & vbCrLf & vbCrLf & SQL
        GstrMsgTitle = "Error - " & Trim(Str(rdoErrors(0).Number))
        MsgBox GstrMsgList, vbCritical, GstrMsgTitle
    Else
        MsgBox Err.Description, vbCritical, "VB Error - " & Trim(Str(Err.Number))
    End If

    Set Rs = Nothing
    RdoOpenSet = -1
    
    Screen.MousePointer = GnMousePointer

End Function

Public Sub RdoCloseSet(ByRef Rs As rdoResultset)

    If Not Rs Is Nothing Then
        Rs.Close
        Set Rs = Nothing
    End If
    
End Sub

Public Function RdoGetString(ByRef Rs As rdoResultset, ByVal rdoCol As String, Optional ByVal AbsPos As Long = -1) As String

    On Error GoTo ReadError
    
    If AbsPos > -1 Then Rs.AbsolutePosition = AbsPos + 1
    RdoGetString = Rs.rdoColumns(rdoCol).Value
    
    Exit Function

ReadError:
    
    RdoGetString = ""
    
    Select Case Err.Number
        Case 40041      'Invalid Column Name
            Debug.Print "RdoGetString Error - 40041", "ORA-00904 : Invalid Column Name - " & rdoCol
        Case 40022      'Invalid Position Number
            Debug.Print "RdoGetString Error - 40022", "Invalid Position Number - " & rdoCol & "(" & CStr(AbsPos) & ")"
        Case 91
        Case 94
        Case Else
            Debug.Print "RdoGetString Error - " & RTrim(Str(Err.Number)), Err.Description
    End Select
    
End Function

Public Function RdoGetNumber(ByRef Rs As rdoResultset, ByVal rdoCol As String, Optional ByVal AbsPos As Long = -1) As Double

    On Error GoTo ReadError
    
    If AbsPos > -1 Then Rs.AbsolutePosition = AbsPos + 1
    RdoGetNumber = IIf(IsNull(Rs.rdoColumns(rdoCol).Value), 0, Rs.rdoColumns(rdoCol).Value)
    
    Exit Function

ReadError:
    
    RdoGetNumber = 0
    
    Select Case Err.Number
        Case 40041      'Invalid Column Name
            Debug.Print "RdoGetNumber Error - 40041", "ORA-00904 : Invalid Column Name - " & rdoCol
        Case 40022      'Invalid Position Number
            Debug.Print "RdoGetNumber Error - 40022", "Invalid Position Number - " & rdoCol & "(" & CStr(AbsPos) & ")"
        Case 91
        Case 94
        Case Else
            Debug.Print "RdoGetNumber Error - " & RTrim(Str(Err.Number)), Err.Description
    End Select
    
End Function

Public Function RdoIsNull(ByRef Rs As rdoResultset, ByVal rdoCol As String, Optional ByVal AbsPos As Long = -1) As Boolean

    On Error GoTo ReadError
    
    If AbsPos > -1 Then Rs.AbsolutePosition = AbsPos + 1
    RdoIsNull = IsNull(Rs.rdoColumns(rdoCol).Value)
    
    Exit Function

ReadError:
    
    RdoIsNull = False
    
    Select Case Err.Number
        Case 40041      'Invalid Column Name
            Debug.Print "RdoIsNull Error - 40041", "ORA-00904 : Invalid Column Name - " & rdoCol
        Case 40022      'Invalid Position Number
            Debug.Print "RdoIsNull Error - 40022", "Invalid Position Number - " & rdoCol & "(" & CStr(AbsPos) & ")"
        Case 91
        Case 94
        Case Else
            Debug.Print "RdoIsNull Error - " & RTrim(Str(Err.Number)), Err.Description
    End Select
    
End Function

Public Function Quot(ByVal strString As String) As String

    Dim I       As Integer
    Dim nPos    As Integer
    
    nPos = 1
    Do
        For I = nPos To Len(strString)
            If Mid(strString, I, 1) = "'" Then
                strString = Left(strString, I - 1) & "''" & Mid(strString, I + 1)
                Exit For
            End If
        Next I
        nPos = I + 2
        If nPos > Len(strString) Then Exit Do
    Loop While (True)
    
    Quot = strString
    
End Function

Public Sub Create_DSN_TEXT(ByVal ArgDSN As String, ByVal ArgDefaultDir As String, ByVal ArgDescription As String)

    Dim strAttribs  As String
    
    strAttribs = "Description=" & ArgDescription & vbCr & "DefaultDir=" & ArgDefaultDir
    rdoEngine.rdoRegisterDataSource ArgDSN, "Microsoft Text Driver (*.txt; *.csv)", True, strAttribs

End Sub

Public Sub Create_DSN_ORACLE(ByVal ArgDSN As String, ByVal ArgServer As String)

    Dim strAttribs  As String
    
    strAttribs = "Description=Oracle ODBC for TWMIS" & vbCr & "Server=" & ArgServer
    rdoEngine.rdoRegisterDataSource ArgDSN, "Microsoft ODBC for Oracle", True, strAttribs

End Sub

