Attribute VB_Name = "AdoConst"
Option Explicit

Public rowindicator         As Long
Public SQL                  As String
Public Result               As Integer
Public strSQL               As String

Public GnMousePointer       As Integer

Public adoConnect           As ADODB.Connection
Public ADORES               As ADODB.Recordset
Public Rs                   As ADODB.Recordset

Public GstrMsgList          As String
Public GstrMsgTitle         As String
Public GstrMsgOpt           As Integer
Public GstrMsgRet           As Integer
    
Public GnJobSabun           As Long
Public GstrJobName          As String
Public GstrJobPart          As String
Public GstrJobGrade         As String

Public GstrPassProgramID    As String * 8
Public GstrPassWord         As String
Public GstrPassGrade        As String
Public GstrPassClass        As String
Public GstrSubClass         As String * 1
Public GstrPassRank         As String
Public GstrPassName         As String
Public GstrPassPart         As String * 1
Public GstrSubPart          As String * 2

Public GstrPassDept         As String
Public GstrIdnumber         As String
Public GstrPassIDnumber     As String
Public GstrSysDate          As String
Public GstrPmpaServer       As String



Public Function DbOdbcConnect(ByVal ArgUser$, ByVal ArgPassword$, ByVal ArgSource$) As Integer

    strSQL = "DSN=" & ArgSource & ";"
    strSQL = strSQL & "UID=" & ArgUser & ";"
    strSQL = strSQL & "PWD=" & ArgPassword & "; "

    On Error GoTo DBConnect_Error

    GnMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    Set adoConnect = New ADODB.Connection

    adoConnect.CursorLocation = adUseClient

    adoConnect.ConnectionString = strSQL
    adoConnect.Open

    Screen.MousePointer = GnMousePointer

    Exit Function

'/-----------------------------------------------------------------------------/
DBConnect_Error:
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description & vbCrLf & vbCrLf & _
           "ConnectString : " & strSQL & vbCrLf & _
           "User          : " & ArgUser & vbCrLf & _
           "Password      : " & ArgPassword, 48, "DbAdoConnect Error"
    End


End Function



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
Public Sub DbAdoDisConnect()

    On Error Resume Next

    If Not Rs Is Nothing Then Rs.Close
    
    adoConnect.Close
    If Not adoConnect Is Nothing Then
        Set adoConnect = Nothing
    End If
    
End Sub

Public Function VarToStr(ByVal sVariable As String) As String
    VarToStr = "'" & sVariable & "'" & vbLf
End Function

Public Function VarToComma(ByVal sVariable As String) As String
    VarToComma = "'" & sVariable & "'," & vbLf
End Function

Public Function NumToComma(ByVal sVariable As String) As String
    NumToComma = sVariable & "," & vbLf
End Function


Public Function adoSQL(ByVal SQL As String) As Integer
    Select Case UCase(left(Trim(SQL), 6))
            Case "SELECT", "FETCH "                              'select�� fetch�� ���� �Լ� ȣ���Ѵ�.
                adoSQL = AdoOpenSet(Rs, SQL)
            Case Else                                   'select �� �ƴѰ��� ���� ���Լ��� ����ϵ��� �Ѵ�.
                adoSQL = AdoExecute(SQL)
    End Select
    
End Function

Public Function AdoExecute(ByVal SQL As String) As Integer
    
    GnMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ExecError:
    
    AdoExecute = 0
    rowindicator = 0
    
    Call adoConnect.Execute(SQL, rowindicator, adCmdText + ADODB.adExecuteNoRecords)
    Screen.MousePointer = GnMousePointer
    
    Exit Function
    
'/----------------------------------------------------------------
ExecError:
    MsgBox "Error.Number : " & adoConnect.Errors(0).Number & vbCrLf & _
           "Error.Description : " & adoConnect.Errors(0).Description & vbCrLf & vbCrLf & _
           "Error.SQL : " & SQL, 48, "AdoExecute Error"
    
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

'/--------------------------------------------------------------
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


Public Sub AdoCloseSet(ByRef sAdoset As ADODB.Recordset)

    On Error GoTo SetClose_Error
    sAdoset.Close
    If Not sAdoset Is Nothing Then Set sAdoset = Nothing
    Exit Sub
    
'/------------------------------------------------------------
SetClose_Error:
    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
           adoConnect.Errors(0).Description
    
End Sub

Public Function AdoGetString(ByRef adoS As ADODB.Recordset, ByVal adoCol As String, Optional ByVal AbsPos As Long = 1) As String
    
    On Error GoTo ReadError

    If AbsPos > -1 Then adoS.AbsolutePosition = AbsPos + 1
    AdoGetString = adoS.Fields(adoCol).Value & ""
    
    Exit Function
    
'/----------------------------------------------------------------

ReadError:
Dim aa As String
     MsgBox "Error.Number : " & adoConnect.Errors(0).Number & vbCrLf & _
           "Error.Description : " & adoConnect.Errors(0).Description & vbCrLf & vbCrLf & _
           "Error.SQL : " & strSQL, 48, "AdoGetString Error"
    
    AdoGetString = ""

End Function

Public Function AdoGetNumber(ByRef adoS As ADODB.Recordset, ByVal adoCol As String, Optional ByVal AbsPos As Long = 1) As Double

    On Error GoTo ReadError

    If AbsPos > -1 Then adoS.AbsolutePosition = AbsPos + 1
    AdoGetNumber = Val(adoS.Fields(adoCol).Value & "")
    
    Exit Function
    
'/--------------------------------------------------------------
ReadError:
    
    AdoGetNumber = 0
    
End Function

Public Sub CloseCursor(ByVal strCursor As String)
     Dim nCursorExistence                As Long

    strSQL = "     SELECT  cursorExistence"
    strSQL = strSQL & "  = Cursor_status ('global' ," & VarToStr(strCursor) & ")  "
    Result = adoSQL(strSQL)
    If rowindicator = 0 Then
        MsgBox "Cursor ��ȸ�� Error �߻�", 48, "�۾�����"
        Exit Sub
    End If
    
    nCursorExistence = AdoGetNumber(Rs, "cursorExistence", 0)
    Select Case nCursorExistence
        Case 1                          'Ŀ���� Open�ƴ�.
            Result = adoSQL("CLOSE       " & strCursor)
            Result = adoSQL("DEALLOCATE  " & strCursor)
        Case 0                          'Ŀ���� Open�ƴ�.
            Result = adoSQL("CLOSE       " & strCursor)
            Result = adoSQL("DEALLOCATE  " & strCursor)
        Case -1                         'Ŀ���� Declear�ƴ�.
            Result = adoSQL("DEALLOCATE  " & strCursor)
        Case -3                         'Ŀ���� ������� �ʾҴ�. �׳� �Ѿ���� �Ѵ�.
        Case Else
    End Select
    

End Sub


'/-------------------------------------------------------------------------------------------/
Public Function DupData_Chk(DupSqlQuery As String) As Boolean
      Dim DupRs                           As New ADODB.Recordset
      
      DupData_Chk = False
      If DupSqlQuery = "" Then Exit Function
      
      Call AdoConst.AdoOpenSet(DupRs, DupSqlQuery)
      
      If rowindicator = 0 Then
            DupData_Chk = False                 '�ߺ�����Ÿ ����
      Else
            DupData_Chk = True
      End If
      DupRs.Close
End Function

Public Function Quot(ByVal strString As String) As String

    Dim i       As Integer
    Dim nPos    As Integer
    
    nPos = 1
    Do
        For i = nPos To Len(strString)
            If Mid(strString, i, 1) = "'" Then
                strString = left(strString, i - 1) & "''" & Mid(strString, i + 1)
                Exit For
            End If
        Next i
        nPos = i + 2
        If nPos > Len(strString) Then Exit Do
    Loop While (True)
    
    Quot = strString
    
End Function

Public Sub ServerNameFetch()


    
    '���ϴ� �ųĸ� ������ �ߴµ� nt_server�̸� ��� server�� �Ѳ�����
    '�����ϵ��� �ϴ°Ŵ�.
    '�����  srvid ��  0 �̸� ó�� ������� serverDB�̴�. �׷��Ƿ�
    
    strSQL = "         SELECT srvname                   " & vbLf
    strSQL = strSQL & "  FROM master.dbo.sysservers     " & vbLf
    strSQL = strSQL & " WHERE srvid     = 0             " & vbLf
    Result = AdoOpenSet(Rs, strSQL)
    If rowindicator = 0 Then
        GstrPmpaServer = ""
    Else
        If UCase(AdoGetString(Rs, "srvname", 0)) = "NMC_PMPA" Then
            GstrPmpaServer = "NMC_PMPA."
        Else
            GstrPmpaServer = "NT_SERVER."
        End If
    End If
    
    GstrPmpaServer = ""
    
End Sub





'Option Explicit
'
'Public rowindicator         As Long
'Public SQL                  As String
'Public Result               As Integer
'Public strSQL               As String
'
'Public GnMousePointer       As Integer
'
'Public adoConnect           As ADODB.Connection
'Public ADORES               As ADODB.Recordset
'Public Rs                   As ADODB.Recordset
'
'Public GstrMsgList          As String
'Public GstrMsgTitle         As String
'Public GstrMsgOpt           As Integer
'Public GstrMsgRet           As Integer
'
'Public GnJobSabun           As Long
'Public GstrJobName          As String
'Public GstrJobPart          As String
'Public GstrJobGrade         As String
'
'Public GstrPassProgramID    As String * 8
'Public GstrPassWord         As String
'Public GstrPassGrade        As String
'Public GstrPassClass        As String
'Public GstrSubClass         As String * 1
'Public GstrPassRank         As String
'Public GstrPassName         As String
'Public GstrPassPart         As String * 1
'Public GstrSubPart          As String * 2
'
'Public GstrPassDept         As String
'Public GstrIdnumber         As String
'Public GstrPassIDnumber     As String
'
'Public GstrSysDate          As String
'Public GstrPmpaServer       As String
'Public GstrPassid           As String
'
'
''Public Function DbAdoConnect(ByVal ArgUser$, ByVal ArgPassword$, ByVal ArgSource$) As Integer
''
''    strSQL = "DSN=" & ArgSource & ";"
''    strSQL = strSQL & "UID=" & ArgUser & ";"
''    strSQL = strSQL & "PWD=" & ArgPassword & "; "
''
''    On Error GoTo DBConnect_Error
''
''    GnMousePointer = Screen.MousePointer
''    Screen.MousePointer = vbHourglass
''
''    Set adoConnect = New ADODB.Connection
''
''    adoConnect.CursorLocation = adUseClient
''
''    adoConnect.ConnectionString = strSQL
''    adoConnect.Open
''
''    Screen.MousePointer = GnMousePointer
''
''    Exit Function
''
'''/-----------------------------------------------------------------------------/
''DBConnect_Error:
''    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
''           adoConnect.Errors(0).Description & vbCrLf & vbCrLf & _
''           "ConnectString : " & strSQL & vbCrLf & _
''           "User          : " & ArgUser & vbCrLf & _
''           "Password      : " & ArgPassword, 48, "DbAdoConnect Error"
''    End
''
''
''End Function
'
'
'
'Public Function DbAdoConnect(ByVal ArgUser$, ByVal ArgPassword$, ByVal ArgSource$) As Integer
'
'    Dim sConString      As String
'
'    sConString = ""
'    sConString = sConString & "Provider=Microsoft OLE DB Provider for Oracle;"
'    sConString = sConString & "Data Source=" & ArgSource & ";"
'    sConString = sConString & "Persist Security Info=False"
'
'    On Error GoTo DBConnect_Error
'
'    GnMousePointer = Screen.MousePointer
'    Screen.MousePointer = vbHourglass
'
'    Set adoConnect = New ADODB.Connection
'
'    adoConnect.CursorLocation = adUseClient
'    adoConnect.Open sConString, ArgUser, ArgPassword
'
'    Screen.MousePointer = GnMousePointer
'
'    Exit Function
'
'
''/-----------------------------------------------------------------------------/
'
'DBConnect_Error:
'
'    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
'           adoConnect.Errors(0).Description & vbCrLf & vbCrLf & _
'           "ConnectString : " & sConString & vbCrLf & _
'           "User          : " & ArgUser & vbCrLf & _
'           "Password      : " & ArgPassword
'
'    End
'
'End Function
'Public Sub DbAdoDisConnect()
'
'    On Error Resume Next
'
'    If Not Rs Is Nothing Then Rs.Close
'
'    adoConnect.Close
'    If Not adoConnect Is Nothing Then
'        Set adoConnect = Nothing
'    End If
'
'End Sub
'
'Public Function VarToStr(ByVal sVariable As String) As String
'    VarToStr = "'" & sVariable & "'" & vbLf
'End Function
'
'Public Function VarToComma(ByVal sVariable As String) As String
'    VarToComma = "'" & sVariable & "'," & vbLf
'End Function
'
'Public Function NumToComma(ByVal sVariable As String) As String
'    NumToComma = sVariable & "," & vbLf
'End Function
'
'
'Public Function adoSQL(ByVal SQL As String) As Integer
'    Select Case UCase(Left(Trim(SQL), 6))
'            Case "SELECT", "FETCH "                              'select�� fetch�� ���� �Լ� ȣ���Ѵ�.
'                adoSQL = AdoOpenSet(Rs, SQL)
'            Case Else                                   'select �� �ƴѰ��� ���� ���Լ��� ����ϵ��� �Ѵ�.
'                adoSQL = AdoExecute(SQL)
'    End Select
'
'End Function
'
'Public Function AdoExecute(ByVal SQL As String) As Integer
'
'    GnMousePointer = Screen.MousePointer
'    Screen.MousePointer = vbHourglass
'
'    On Error GoTo ExecError:
'
'    AdoExecute = 0
'    rowindicator = 0
'
'    Call adoConnect.Execute(SQL, rowindicator, adCmdText + ADODB.adExecuteNoRecords)
'    Screen.MousePointer = GnMousePointer
'
'    Exit Function
'
''/----------------------------------------------------------------
'ExecError:
'    MsgBox "Error.Number : " & adoConnect.Errors(0).Number & vbCrLf & _
'           "Error.Description : " & adoConnect.Errors(0).Description & vbCrLf & vbCrLf & _
'           "Error.SQL : " & SQL, 48, "AdoExecute Error"
'
'    AdoExecute = -1
'    rowindicator = -1
'
'    Screen.MousePointer = GnMousePointer
'
'End Function
'
'Public Function AdoExecute1(ByVal SQL As String) As Integer
'
'    On Error GoTo ExecError:
'    AdoExecute1 = 0
'    rowindicator = 0
'
'    adoConnect.Execute SQL, rowindicator, adCmdText
'
'    Exit Function
'
''/--------------------------------------------------------------
'ExecError:
'    AdoExecute1 = -1
'    rowindicator = -1
'
'End Function
'
'Public Function AdoOpenSet(ByRef sAdoset As ADODB.Recordset, ByVal SQL As String, Optional ByVal nRowCnt As Boolean = True, Optional ByVal nMousePointer = True) As Integer
'
'    Set sAdoset = New ADODB.Recordset
'
'    If nMousePointer = True Then
'        GnMousePointer = Screen.MousePointer
'        Screen.MousePointer = vbHourglass
'    End If
'    On Error GoTo OpenError:
'
'    AdoOpenSet = 0
'    rowindicator = 0
'
'    If nRowCnt = True Then
'        adoConnect.CursorLocation = adUseClient
'    Else
'        adoConnect.CursorLocation = adUseServer
'    End If
'
'    'Set sAdoset = adoConnect.Execute(SQL, Rowindicator, adCmdText)
'    Call sAdoset.Open(SQL, adoConnect, adOpenStatic, adLockReadOnly, adCmdText)
'    If Not sAdoset.EOF Then
'        If nRowCnt = True Then
'            rowindicator = sAdoset.RecordCount
'        Else
'            rowindicator = -1
'        End If
'    End If
'
'    If nMousePointer = True Then
'        Screen.MousePointer = GnMousePointer
'    End If
'
'    Exit Function
'
'
'OpenError:
'
'    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
'           adoConnect.Errors(0).Description & vbCrLf & _
'           SQL
'
'    AdoOpenSet = -1
'
'    Screen.MousePointer = GnMousePointer
'
'End Function
'
'
'Public Sub AdoCloseSet(ByRef sAdoset As ADODB.Recordset)
'
'    On Error GoTo SetClose_Error
'    sAdoset.Close
'    If Not sAdoset Is Nothing Then Set sAdoset = Nothing
'    Exit Sub
'
''/------------------------------------------------------------
'SetClose_Error:
'    MsgBox adoConnect.Errors(0).Number & vbCrLf & _
'           adoConnect.Errors(0).Description
'
'End Sub
'
'Public Function AdoGetString(ByRef adoS As ADODB.Recordset, ByVal adoCol As String, Optional ByVal AbsPos As Long = 1) As String
'
'    On Error GoTo ReadError
'
'    If AbsPos > -1 Then adoS.AbsolutePosition = AbsPos + 1
'    AdoGetString = adoS.Fields(adoCol).Value & ""
'
'    Exit Function
'
''/----------------------------------------------------------------
'
'ReadError:
'Dim aa As String
'     MsgBox "Error.Number : " & adoConnect.Errors(0).Number & vbCrLf & _
'           "Error.Description : " & adoConnect.Errors(0).Description & vbCrLf & vbCrLf & _
'           "Error.SQL : " & strSQL, 48, "AdoGetString Error"
'
'    AdoGetString = ""
'
'End Function
'
'Public Function AdoGetNumber(ByRef adoS As ADODB.Recordset, ByVal adoCol As String, Optional ByVal AbsPos As Long = 1) As Double
'
'    On Error GoTo ReadError
'
'    If AbsPos > -1 Then adoS.AbsolutePosition = AbsPos + 1
'    AdoGetNumber = Val(adoS.Fields(adoCol).Value & "")
'
'    Exit Function
'
''/--------------------------------------------------------------
'ReadError:
'
'    AdoGetNumber = 0
'
'End Function
'
'Public Sub CloseCursor(ByVal strCursor As String)
'     Dim nCursorExistence                As Long
'
'    strSQL = "     SELECT  cursorExistence"
'    strSQL = strSQL & "  = Cursor_status ('global' ," & VarToStr(strCursor) & ")  "
'    Result = adoSQL(strSQL)
'    If rowindicator = 0 Then
'        MsgBox "Cursor ��ȸ�� Error �߻�", 48, "�۾�����"
'        Exit Sub
'    End If
'
'    nCursorExistence = AdoGetNumber(Rs, "cursorExistence", 0)
'    Select Case nCursorExistence
'        Case 1                          'Ŀ���� Open�ƴ�.
'            Result = adoSQL("CLOSE       " & strCursor)
'            Result = adoSQL("DEALLOCATE  " & strCursor)
'        Case 0                          'Ŀ���� Open�ƴ�.
'            Result = adoSQL("CLOSE       " & strCursor)
'            Result = adoSQL("DEALLOCATE  " & strCursor)
'        Case -1                         'Ŀ���� Declear�ƴ�.
'            Result = adoSQL("DEALLOCATE  " & strCursor)
'        Case -3                         'Ŀ���� ������� �ʾҴ�. �׳� �Ѿ���� �Ѵ�.
'        Case Else
'    End Select
'
'
'End Sub
'
'
''/-------------------------------------------------------------------------------------------/
'Public Function DupData_Chk(DupSqlQuery As String) As Boolean
'      Dim DupRs                           As New ADODB.Recordset
'
'      DupData_Chk = False
'      If DupSqlQuery = "" Then Exit Function
'
'      Call AdoConst.AdoOpenSet(DupRs, DupSqlQuery)
'
'      If rowindicator = 0 Then
'            DupData_Chk = False                 '�ߺ�����Ÿ ����
'      Else
'            DupData_Chk = True
'      End If
'      DupRs.Close
'End Function
'
'Public Function Quot(ByVal strString As String) As String
'
'    Dim i       As Integer
'    Dim nPos    As Integer
'
'    nPos = 1
'    Do
'        For i = nPos To Len(strString)
'            If Mid(strString, i, 1) = "'" Then
'                strString = Left(strString, i - 1) & "''" & Mid(strString, i + 1)
'                Exit For
'            End If
'        Next i
'        nPos = i + 2
'        If nPos > Len(strString) Then Exit Do
'    Loop While (True)
'
'    Quot = strString
'
'End Function
'
'Public Sub ServerNameFetch()
'
'
'
'    '���ϴ� �ųĸ� ������ �ߴµ� nt_server�̸� ��� server�� �Ѳ�����
'    '�����ϵ��� �ϴ°Ŵ�.
'    '�����  srvid ��  0 �̸� ó�� ������� serverDB�̴�. �׷��Ƿ�
'
'    strSQL = "         SELECT srvname                   " & vbLf
'    strSQL = strSQL & "  FROM master.dbo.sysservers     " & vbLf
'    strSQL = strSQL & " WHERE srvid     = 0             " & vbLf
'    Result = AdoOpenSet(Rs, strSQL)
'    If rowindicator = 0 Then
'        GstrPmpaServer = ""
'    Else
'        If UCase(AdoGetString(Rs, "srvname", 0)) = "NMC_PMPA" Then
'            GstrPmpaServer = "NMC_PMPA."
'        Else
'            GstrPmpaServer = "NT_SERVER."
'        End If
'    End If
'
'    GstrPmpaServer = ""
'
'End Sub
'
'
'
'
