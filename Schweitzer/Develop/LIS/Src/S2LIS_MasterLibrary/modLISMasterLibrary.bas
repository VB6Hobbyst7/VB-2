Attribute VB_Name = "modLISMasterLibrary"
Option Explicit

'Global gIsDeveloper As Boolean
Global gBuildingCd As String
'Global gEmpId As String
Global gParentWhnd As Long
Global lstItemList As New clsDictionary

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
'             (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
'              ByVal lpString As Any, ByVal lpFileName As String) As Long
'
'Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
'             (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
'              ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function InsertData(pSql() As String, Optional ByVal blnTrans As Boolean = True) As Boolean
'Coding By Legends
'쿼리문 실행
'blnTrans : True : 내부에서 트랜잭션을 제어하는 경우, False : 외부에서 트랜잭션을 제어하는 경우.

'쿼리문을 실행하는데 있어서 이 메서드를 사용하려면 쿼리는 필히 배열값으로 넘겨 주어야 한다.
'사용 예)
'   Dim arySQL() As String     배열 선언
'   redim ary(1)               배열 할당
'   arySQL(0) = objMySQL.SetTestItemMst(파라미터)
'   call objMySQL.InsertData(arySQL)
    Dim i As Long
    Dim lngCnt As Long
    On Error GoTo ErrInsertData
    
    lngCnt = UBound(pSql)
    
    With DBConn
        If blnTrans Then .BeginTrans
        
        For i = LBound(pSql) To UBound(pSql)
            'Debug.Print I & " : " & pSql(I)
            If pSql(i) <> "" Then .Execute pSql(i)
        Next
        
        If blnTrans Then .CommitTrans
        InsertData = True
        Exit Function
    End With

ErrInsertData:
    With DBConn
        If blnTrans Then
            .RollbackTrans
            MsgBox Err.Description, vbExclamation
        End If
        InsertData = False
    End With
End Function

'Public Function StripTerminator(ByVal strString As String) As String
'    Dim intZeroPos As Long
'
'    intZeroPos = InStr(strString, Chr$(0))
'    If intZeroPos > 0 Then
'        StripTerminator = VBA.Left$(strString, intZeroPos - 1)
'    Else
'        StripTerminator = strString
'    End If
'End Function


