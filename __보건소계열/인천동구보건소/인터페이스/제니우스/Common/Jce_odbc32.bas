Attribute VB_Name = "ODBC"
Option Explicit
Type connect
   henv    As Long
   hdbc    As Long
   hstmt   As Long
   useflag As Integer
End Type

Public ODBCConnect(1 To 10) As connect
Public Iniflag As String
Public rc As Integer
Public Const gSQL_STR_LEN = 4096
Public Const gnRET_BUF_MAX = 256             ' Default String Buffer Lengt


Function DescribeError(ByVal henv, ByVal hdbc As Long, ByVal hstmt As Long) As Long
   Const SbufferLen = SQL_MAX_MESSAGE_LENGTH
   Dim Cr$, L%
   Cr$ = Chr(13) & Chr(10)
   Dim rgbValue1 As String * 16
   Dim rgbValue3 As String * SbufferLen
   Dim OutLen As Integer
   Dim Native As Long
   Dim lrc As Integer
   Dim lhenv As Long, lhdbc As Long, lhstmt As Long
   DescribeError = 0
   rgbValue1 = String(16, 0)
   rgbValue3 = String(SbufferLen, 0)
   For L = 1 To 3

      Select Case L
         Case 1
            lhenv = SQL_NULL_HENV
            lhdbc = SQL_NULL_HDBC
            lhstmt = hstmt
         Case 2
            lhenv = SQL_NULL_HENV
            lhdbc = hdbc
            lhstmt = SQL_NULL_HSTMT
         Case 3
            lhenv = henv
            lhdbc = SQL_NULL_HDBC
            lhstmt = SQL_NULL_HSTMT
      End Select
      
      Do
         lrc = SQLError(lhenv, lhdbc, lhstmt, rgbValue1, Native, rgbValue3, SbufferLen, OutLen)
         If lrc = QSQL_SUCCESS Or lrc = QSQL_SUCCESS_WITH_INFO Then
            If OutLen = 0 Then
'                97.10.30 khj
'                MsgBox "Error -- No error information available"
                OdbcErrorMsg "Error -- No error information available", "ODBC Error"
            Else
                If lrc = SQL_ERROR Then
'                   97.10.30 khj
'                    MsgBox Left$(rgbValue3, OutLen)
                    OdbcErrorMsg Left$(rgbValue3, OutLen), "ODBC Error"
                Else
                    If Native <> 2601 Then
'                       97.10.30 khj
'                        Msgox Left$(rgbValue3, OutLen) & Cr & "Native error:" & Native
                        OdbcErrorMsg Left$(rgbValue3, OutLen) & "Native Error : " & Native, "ODBC Error"
                    Else
                        Exit Do
                    End If
                End If
            End If
         End If
      Loop Until lrc <> QSQL_SUCCESS
      If Native = 2601 Then
         DescribeError = Native
      End If
   Next L
End Function

Function QSqlBeginTrans() As Integer
   Dim i%
   For i = 1 To 10
      If ODBCConnect(i).useflag = True Then
         rc = SQLSetConnectOption(ODBCConnect(i).hdbc, SQL_AUTOCOMMIT, SQL_AUTOCOMMIT_OFF)
      End If
   Next i
End Function

Function Qsqlclose(Index As Long, finish As Integer) As Integer
   Dim i%
   i = Index
   If ODBCConnect(i).useflag = False Then
      'MsgBox "닫혀진 INDEX CLOSE" & Chr(13) & Chr(10) & "JCE에 연락주세요"
      Exit Function
   End If
   rc = SQLDisconnect(ODBCConnect(i).hdbc)
   rc = SQLFreeConnect(ODBCConnect(i).hdbc)
   rc = SQLFreeEnv(ODBCConnect(i).henv)
   ODBCConnect(i).useflag = False
End Function

Function QSqlCommitTrans() As Integer
   'Index As Integer
   'rc = SQLTransact(ODBCConnect(Index).henv, ODBCConnect(Index).hdbc, SQL_COMMIT)
   Dim i%
   For i = 1 To 10
      If ODBCConnect(i).useflag = True Then
         'rc = SQLSetConnectOption(ODBCConnect(i).hdbc, SQL_AUTOCOMMIT, SQL_AUTOCOMMIT_OFF)
         rc = SQLTransact(ODBCConnect(i).henv, ODBCConnect(i).hdbc, SQL_COMMIT)
         rc = SQLSetConnectOption(ODBCConnect(i).hdbc, SQL_AUTOCOMMIT, SQL_AUTOCOMMIT_ON)
      End If
   Next i

End Function


Function QSqlDBExec(SqlStr As String, Index As Long) As Integer
   Dim sSql As String * gSQL_STR_LEN
   Dim nRetCode As Integer
   Dim ErrorCode As Long
   
   Dim i%
   i = Index
   
'   On Error Resume Next
   If ODBCConnect(i).hstmt <> 0 Then
'        97.10.30 khj
'        MsgBox "hstmt is not 0 " & Chr(13) & Chr(10) & "JCE에 연락주세요"
        OdbcErrorMsg "hstmt is not 0 " & Chr(13) & Chr(10) & "에 연락주세요.", "ODBC Error"
'        Exit Function
   End If
   nRetCode = SQLAllocStmt(ODBCConnect(i).hdbc, ODBCConnect(i).hstmt)
   If nRetCode <> QSQL_SUCCESS Then GoTo roExit
   
   sSql = String$(gSQL_STR_LEN, 0)
   sSql = SqlStr
   
   nRetCode = SQLExecDirect(ODBCConnect(i).hstmt, sSql, Len(sSql))
   
   If nRetCode <> QSQL_SUCCESS Then
      ErrorCode = DescribeError(ODBCConnect(i).henv, ODBCConnect(i).hdbc, ODBCConnect(i).hstmt)
      QSqlDBExec = ErrorCode
      Select Case ErrorCode
         Case 2601   ' insert duplicate error
            QSqlDBExec = 1
      End Select
   End If

   Dim pos_update%, pos_insert%, pos_delete%, pcrow As Long
   
   pos_delete = InStr(1, SqlStr, "DELETE", 1)
   pos_update = InStr(1, SqlStr, "UPDATE", 1)
   pos_insert = InStr(1, SqlStr, "INSERT", 1)

   If pos_update > 0 Or pos_insert > 0 Or pos_delete > 0 Then
      nRetCode = SQLRowCount(ODBCConnect(i).hstmt, pcrow)
      If pcrow = 0 And ErrorCode <> 2601 Then                'No Record Count or Error
        QSqlDBExec = 2
      Else
        ''''''''''''''
      End If
      nRetCode = SQLFreeStmt(ODBCConnect(i).hstmt, SQL_DROP)
      ODBCConnect(i).hstmt = 0
   End If

   GoTo roExit
   
roExit:
   If nRetCode <> QSQL_SUCCESS Then
      'DescribeError ODBCConnect(i).henv, ODBCConnect(i).hdbc, 0
      Beep
      
      Exit Function
   End If
   
End Function

Function QSqlGetRow(record As String, Index As Long) As Integer
   Dim sSql As String
   Dim nRetCode As Integer
   Dim lRetLen As Long
   Dim lCount As Long
   Dim sRetBuf As String * gnRET_BUF_MAX
   Dim i%, RCols%, Tmp$

   Tmp = ""
  
'   On Error Resume Next
   nRetCode = SQLNumResultCols(ODBCConnect(Index).hstmt, RCols)

   lCount = 0
   nRetCode = SQLFetch(ODBCConnect(Index).hstmt)    'move cursor to next row
   If nRetCode = QSQL_SUCCESS Then
      For i = 1 To RCols
         nRetCode = SQLGetData(ODBCConnect(Index).hstmt, i, SQL_C_CHAR, sRetBuf, gnRET_BUF_MAX, lRetLen)
      
         If lRetLen > 0 Then
            Tmp = Tmp & Left(sRetBuf, lRetLen) & Chr(5)
         Else
            Tmp = Tmp & Chr(5)
         End If
      Next i
      lCount = lCount + 1
   Else
     QSqlGetRow = 100
   End If

   If lCount > 0 Then
      record = Tmp
   Else
      record = ""
      For i = 1 To RCols
         record = record & Chr(5)
      Next i
   End If
   If nRetCode <> QSQL_SUCCESS Then GoTo roExitErr1
   GoTo roExit1

roExitErr1:
roExit1:

End Function

Function QSqlOpen(Server As String, hWnd As Long, Index As Long) As Integer
   
   Dim connect       As String
   Dim connectout    As String * 255
   Dim connectoutlen As Integer
   Dim i%
   If Iniflag <> "Initialize" Then
      For i = 1 To 10
         ODBCConnect(i).useflag = False
      Next i
      Iniflag = "Initialize"
   End If
   For i = 1 To 10
      If ODBCConnect(i).useflag = False Then
         Index = i
         Exit For
      End If
   Next i
   If i = 11 Then
'        97.10.30 khj
'        MsgBox "DB Open Index Full"
        OdbcErrorMsg "DB Open Index Full", "ODBC Error"
        Exit Function
   End If
   connectoutlen = 0
   connectout = String(255, 0)
   rc = SQLAllocEnv(ODBCConnect(i).henv)
   If rc <> QSQL_SUCCESS Then Exit Function
   rc = SQLAllocConnect(ODBCConnect(i).henv, ODBCConnect(i).hdbc)
   If rc <> QSQL_SUCCESS Then Exit Function
   
   connect = Server
   Dim ErrorCode As Long
   rc = SQLDriverConnect(ODBCConnect(i).hdbc, hWnd, connect$, Len(connect$), connectout, 255, connectoutlen, SQL_DRIVER_COMPLETE)
   If rc = QSQL_SUCCESS Or rc = QSQL_SUCCESS_WITH_INFO Then
      rc = 0
   Else
      ErrorCode = DescribeError(ODBCConnect(i).henv, ODBCConnect(i).hdbc, 0)
   End If
   ODBCConnect(i).useflag = True
   QSqlOpen = rc
End Function

Function QSqlRollBack() As Integer
   'Index As Integer

   'rc = SQLTransact(ODBCConnect(Index).henv, ODBCConnect(Index).hdbc, SQL_ROLLBACK)
   Dim i%
   For i = 1 To 10
      If ODBCConnect(i).useflag = True Then
         'rc = SQLSetConnectOption(ODBCConnect(i).hdbc, SQL_AUTOCOMMIT, SQL_AUTOCOMMIT_OFF)
         rc = SQLTransact(ODBCConnect(i).henv, ODBCConnect(i).hdbc, SQL_ROLLBACK)
         rc = SQLSetConnectOption(ODBCConnect(i).hdbc, SQL_AUTOCOMMIT, SQL_AUTOCOMMIT_ON)
      End If
   Next i

End Function

Function QSqlSelectFree(Index As Long)
   Dim i%
   i = Index
   If ODBCConnect(i).hstmt > 0 Then
      rc = SQLFreeStmt(ODBCConnect(i).hstmt, SQL_DROP) ', "Unable to free statment handle"
   'Else
   '   ODBCConnect(i).hstmt = 0
   End If
   ODBCConnect(i).hstmt = 0
End Function


Public Sub OdbcErrorMsg(Msg As String, Title As String)
    
    Dim c As Integer
    Dim Tmp As String
    
    Open App.Path & "\ODBCErrorMsg.log" For Output As #9     'Dump 용
    Print #9, Format(Now, "yyyy-mm-dd hh:mm:ss") & Chr$(13) & Chr$(10) & Msg
    Close #9
    
    For c = 1 To Len(Msg)
        Tmp = Mid(Msg, c, 1)
        If Tmp = "." Or Tmp = "," Or Tmp = "]" Then
            If Mid(Msg, c + 1, 1) = "." Or Mid(Msg, c + 1, 1) = "[" Then
                c = c + 1
            Else
                If c <> Len(Msg) Then Msg = Left(Msg, c) & Chr(13) & Chr(10) & Mid(Msg, c + 1)
            End If
       End If
    Next
    
    With frmODBCErrorMsg
        .Height = 9000
        .Width = 12000
        .lblMsg = Msg
        .Height = .lblMsg.Height + 1300
        .Width = .lblMsg.Width + 1200
        .cmdConfirm.Left = (.Width - .cmdConfirm.Width - 100) / 2
        .cmdConfirm.Top = .Height - 500 - .cmdConfirm.Height
        .Caption = Title
        .Show 1
    End With
    
End Sub
