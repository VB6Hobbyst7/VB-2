Attribute VB_Name = "LOGON"
Option Explicit

Public Function Get_Record(ByVal SqlCode As Long _
                         , ByVal Field As String, Table As String _
                         , Key As String, KeyVal As String) As String

    Dim ReturnVal() As String
    Dim SqlStr      As String, ret As Integer
    
    SqlStr = " SELECT DISTINCT " & Field & " FROM " & Table _
            & " WHERE " & Key & " = '" & KeyVal & Chr$(39)
     
    ret = QSqlDBExec(SqlStr, SqlCode)
    If ret = QSQL_SUCCESS Then
        If QSqlGetRow(record, SqlCode) = QSQL_SUCCESS Then

            QSqlGetField 1, record, ReturnVal()
            Get_Record = Trim(ReturnVal(1))
        End If
    End If
    
    Call QSqlSelectFree(SqlCode)
    
    
End Function
'*------------------------------------------------------*
'*                                                      *
'*  Record의 존재하면 True, 아니면 False                *
'*  para : SQL 문                                       *
'*                                                      *
'*------------------------------------------------------*
Function G_EXIST_RECORD(ByVal SqlCode As Long, para As String) As Integer
    
    Dim status  As Integer
    Dim Row     As Integer
    Dim sStr    As String
    Dim tData() As String

    status = QSqlDBExec(para, SqlCode)
    If status <> QSQL_SUCCESS Then
        Call QSqlSelectFree(SqlCode)
        G_EXIST_RECORD = False
        Exit Function
    End If

    status = QSqlGetRow(sStr, SqlCode)
    If status <> QSQL_SUCCESS Then
        Call QSqlSelectFree(SqlCode)
        G_EXIST_RECORD = False
        Exit Function
    End If

    QSqlGetField 1, sStr, tData()

    If Val(tData(1)) = 0 Then
        G_EXIST_RECORD = False
    Else
        G_EXIST_RECORD = True
    End If
    
    Call QSqlSelectFree(SqlCode)

End Function


