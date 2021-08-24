Attribute VB_Name = "modSombrero"



'% DSN�� ����Ͽ� ODBC�� �����Ѵ�
Public Function DBConnect() As Long
    
   Dim iRetry As Integer
   
   On Error GoTo ConnectError
   
   iRetry = 0
   DBConnect = CONNECT_ERROR

DoAgain:
   iRetry = iRetry + 1
   Set DbConn = New DrSqlOcx.DrDatabase
   With DbConn
      .Server = SB_ServerNm    ' "GILBASE36"
      .dbname = SB_DatabaseNm    '"HIS_DB"
      .uID = SB_LoginId    '"hisbase"
      .PWD = SB_Password   '"hispass"
      .DbOpen
      If .DBConnect Then
        DBConnect = CONNECT_SUCCESS 'Oracle Connection Success !
        Date = Format(Get_SysDate, CS_DateMask)
        Time = Format(Get_SysTime, CS_TimeLMask)
      Else
        If iRetry < 3 Then GoTo DoAgain  '������ �ȵ� ��� 3������ ��õ�..
      End If
        
   End With
   
   Exit Function

ConnectError:
   
   'MsgBox DbConn.Errors.Item(0).Number & " : " & DbConn.Errors.Item(0).Description
   'MsgBox "Database ������ �ȉ���ϴ�. ����Ƿ� ���� �ٶ��ϴ�."
   DbConn.DbClose
   Set DbConn = Nothing
   DBConnect = CONNECT_ERROR

End Function


'% Query������ �޾Ƽ� ������ �� ������ RecordSet�� �Ѱ��ش�.
Public Function OpenRecordSet(ByVal SqlStmt As String, Optional ByVal ReadOnly As Variant, Optional ByVal CursorType As Variant, _
                                              Optional ByVal MyDb As Variant) As DrSqlOcx.DrRecordSet

   'Dim tmpRs As New ADODB.Recordset
   Dim tmpCursorType As Integer
   Dim tmpLockType As Integer
   
   On Error GoTo Err_Trap
   
   iRetry = 0
   Set OpenRecordSet = New DrSqlOcx.DrRecordSet
   With OpenRecordSet
      'DbConn : ���� ����, MyDb : ���ο� ����
      If IsMissing(MyDb) Then
         Set .ActiveConnection = DbConn
      Else
         Set .ActiveConnection = MyDb
      End If
      .SqlStatement = SqlStmt
      .RsOpen    ', , tmpCursorType  ', tmpLockType
      If .DBerror Then
         .RsClose
         DbConn.DbClose
         DBConnect
         .RsOpen
      End If
   End With
   
   'Set OpenRecordSet = tmpRs
   Exit Function

Err_Trap:
   MsgBox Err.Description
   'Set OpenRecordSet = Nothing

End Function


'% Query������ �����Ѵ�.
Public Function ExecuteSql(ByVal SqlStmt As String, Optional ByRef MyRs As Variant, Optional ByVal MyDb As Variant) As Boolean

   Dim MyCmd As New DrSqlOcx.DrCommand
   
   On Error GoTo Err_Trap
   
   'Command
   With MyCmd
      'DbConn : ���� ����, MyDb : ���ο� ����
      If IsMissing(MyDb) Then
         Set .ActiveConnection = DbConn
      Else
         Set .ActiveConnection = MyDb
      End If
      'Sql����
      .SqlStatement = SqlStmt
      'If IsMissing(MyRs) Then
         .Execute    '�ܼ� ����
      'Else
      '   Set MyRs = .Execute    'Record Set ��ȯ
      'End If
   End With
   ExecuteSql = True
   Set MyCmd = Nothing
   Exit Function
   
Err_Trap:
   Call Error_Routine
   Set MyCmd = Nothing
   ExecuteSql = False
   
End Function

Public Sub Error_Routine(Optional ByVal MyDb As Variant)
   
   Dim MyDb1 As DrSqlOcx.DrDatabase
   Dim I As Integer
   
   If IsMissing(MyDb) Then
      Set MyDb1 = DbConn
   Else
      Set MyDb1 = MyDb
   End If
   
   If MyDb1.Errors.Count > 0 Then
      Dim errLoop As DrSqlOcx.DrError
      Dim tmpError As String
      For I = 1 To MyDb1.Errors.Count
         tmpError = tmpError & _
                        "Error #" & MyDb1.Errors.Item(I).Number & vbCr & _
                        "   " & MyDb1.Errors.Item(I).Description & vbCr
      Next
      MyDb1.RollbackTrans
      MyDb1.Errors.Clear
      Set MyDb1 = Nothing
      MsgBox tmpError
   ElseIf Err.Number <> 0 Then
      MsgBox Err.Description
   End If

End Sub

Public Sub DbClose()

    DbConn.DbClose
    Set DbConn = Nothing
    
End Sub


Public Function Get_SysDate()

    Dim tmpRs As DrSqlOcx.DrRecordSet
    
    Set tmpRs = OpenRecordSet("select " & CS_SybaseDate & " as Today")
    If tmpRs.EOF Then
        Get_SysDate = Format(Now, CS_DateDbFormat)
    Else
        Get_SysDate = tmpRs.Fields("Today").Value
    End If
    
    tmpRs.RsClose
    Set tmpRs = Nothing

End Function

Public Function Get_SysTime()

    Dim tmpRs As DrSqlOcx.DrRecordSet
    
    Set tmpRs = OpenRecordSet("select " & CS_SybaseTime & " as Time")
    If tmpRs.EOF Then
        Get_SysDate = Format(Now, CS_TimeDbFormat)
    Else
        Get_SysTime = tmpRs.Fields("Time").Value
    End If
    
    tmpRs.RsClose
    Set tmpRs = Nothing

End Function

