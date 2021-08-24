Attribute VB_Name = "modServerConnection"
Option Explicit

Global objSysInfo As New clsS2DSO
Global objMyUser As New clsDSMLogOn

Global DbConn As DrDatabase
Global isDBOPEN As Boolean


Public Function DBConnect() As Long
    
    Dim iRetry As Integer
    
    On Error GoTo ConnectError
    
    iRetry = 0
    isDBOPEN = False

DoAgain:
    iRetry = iRetry + 1
    Set DbConn = New DrDatabase
    With DbConn
        .Whatsthis = .ThisIsSybase
        .Server = objSysInfo.ServerNm     '"SKY"
        .dbname = objSysInfo.DatabaseNm   '"LIS_DB"
        .uid = objSysInfo.DBLoginId       '"lisbase"
        .pwd = objSysInfo.DBPassword      '"lispass"
        .DbOpen
        If .DBConnect Then
            isDBOPEN = True
            Date = .GetSysDate
            Time = .GetSysDate
        Else
            If iRetry < 3 Then GoTo DoAgain  '연결이 안될 경우 3번까지 재시도..
        End If
         
    End With
    
    Exit Function

ConnectError:
   
    MsgBox DbConn.Errors(0).Number & " : " & DbConn.Errors(0).Description
    MsgBox "Database 연결이 안됬습니다. 전산실로 문의 바랍니다."
    DbConn.DbClose
    Set DbConn = Nothing
    isDBOPEN = False

End Function


'% Query문장을 받아서 실행한 후 생성된 RecordSet을 넘겨준다.
Public Function OpenRecordSet(ByVal SqlStmt As String, Optional ByVal ReadOnly As Variant, Optional ByVal CursorType As Variant, _
                                              Optional ByVal MyDb As Variant) As DrRecordSet

   Dim tmpCursorType As Integer
   Dim tmpLockType As Integer
   Dim iRetry As Long
   
   On Error GoTo Err_Trap
   
   iRetry = 0
   Set OpenRecordSet = New DrRecordSet
   With OpenRecordSet
      'DbConn : 기존 연결, MyDb : 새로운 연결
      If IsMissing(MyDb) Then
         Set .ActiveConnection = DbConn
      Else
         Set .ActiveConnection = MyDb
      End If
      .SqlStatement = SqlStmt
      .RsOpen , SqlStmt   ', , tmpCursorType  ', tmpLockType
      If .DBerror Then
         .RsClose
         DbConn.DbClose
         DBConnect
         .RsOpen , SqlStmt
      End If
   End With
   
   DbConn.Errors.Clear
   Exit Function

Err_Trap:
   MsgBox Err.Description
   Set OpenRecordSet = Nothing

End Function


'% Query문장을 실행한다.
Public Function ExecuteSql(ByVal SqlStmt As String, Optional ByRef MyRs As Variant, Optional ByVal MyDb As Variant) As Boolean

'   Dim MyCmd As New DrCommand
'
'   On Error GoTo Err_Trap
'
'   'Command
'   With MyCmd
'      'DbConn : 기존 연결, MyDb : 새로운 연결
'      If IsMissing(MyDb) Then
'         Set .ActiveConnection = DbConn
'      Else
'         Set .ActiveConnection = MyDb
'      End If
'      'Sql문장
'      .SqlStatement = SqlStmt
'      'If IsMissing(MyRs) Then
'         .Execute    '단순 실행
'      'Else
'      '   Set MyRs = .Execute    'Record Set 반환
'      'End If
'   End With
'   ExecuteSql = True
'   Set MyCmd = Nothing
'   Exit Function
'
'Err_Trap:
'   Call Error_Routine
'   Set MyCmd = Nothing
'   ExecuteSql = False
   
End Function

Public Sub Error_Routine(Optional ByVal MyDb As Variant)
   
   'Dim MyDb1 As DrSqlOcx.DrDatabase
   
   Dim i As Integer
   Dim tmpError As String
   
   'If IsMissing(MyDb) Then
   '   Set MyDb1 = DbConn
   'Else
   '   Set MyDb1 = MyDb
   'End If
   
   If Err.Number <> 0 Then tmpError = Err.Description & vbCr
   If DbConn.Errors.Count > 0 Then
      Dim errLoop As DrError
      For i = 1 To DbConn.Errors.Count
         tmpError = tmpError & _
                        "Error #" & DbConn.Errors.Item(i).Number & vbCr & _
                        "   " & DbConn.Errors.Item(i).Description & vbCr
      Next
      DbConn.RollbackTrans
      'Set MyDb1 = Nothing
   End If
   MsgBox tmpError, , "Error"
   DbConn.Errors.Clear

End Sub

Public Sub DbClose()

    On Error GoTo ErrDbClose
    
    If isDBOPEN Then DbConn.DbClose
    Set DbConn = Nothing

    Exit Sub
ErrDbClose:
    
End Sub

