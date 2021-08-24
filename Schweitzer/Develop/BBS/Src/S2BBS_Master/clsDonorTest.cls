VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDonorTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetGroup(ByRef pGroup() As String) As Long
    Dim SSQL As String
    Dim RS   As Recordset
    Dim i      As Long
    
    SSQL = "SELECT distinct cdval1 " & _
             "FROM " & T_COM002 & " " & _
             "WHERE " & DBW("cdindex=", "B003")
    Set RS = New Recordset
    Call RS.Open(SSQL, DBConn)
    
    GetGroup = RS.RecordCount
    
    For i = 1 To RS.RecordCount
        ReDim Preserve pGroup(i - 1)
        pGroup(i - 1) = RS.Fields("cdval1").Value & ""
        RS.MoveNext
    Next i
    
    Set RS = Nothing
End Function

Public Function GetTestOfGroup(ByVal pGroup As String, ByRef pTest() As String) As Long
    Dim SSQL As String
    Dim RS   As Recordset
    Dim i      As Long
    
    'pColTest : 검사코드,검사명,검체코드
    
    SSQL = " SELECT a.cdval1 groupcd,a.cdval2 testcd,b.testnm,a.field1 speccd " & _
           " FROM " & T_COM002 & " a, " & T_LAB001 & " b " & _
           " WHERE " & DBW("a.cdindex=", "B003") & " " & _
           " AND   " & DBW("a.cdval1=", pGroup) & " " & _
           " AND   a.cdval2=b.testcd " & _
           " AND   b.applydt=(SELECT max(c.applydt) " & _
                            " FROM  " & T_LAB001 & " c " & _
                            " WHERE c.testcd=a.cdval2) "

    Set RS = New Recordset
    RS.Open SSQL, DBConn
    GetTestOfGroup = RS.RecordCount
    
    For i = 1 To RS.RecordCount
        ReDim Preserve pTest(i - 1)
        pTest(i - 1) = RS.Fields("testcd").Value & "" & vbTab & _
                       RS.Fields("testnm").Value & "" & vbTab & _
                       RS.Fields("speccd").Value & ""
        RS.MoveNext
    Next i
    
    Set RS = Nothing
    
End Function

Public Function Delete(ByVal pGroup As String, ByVal pTestCd As String) As Boolean
    Dim SSQL As String
    
On Error GoTo Delete_error

    SSQL = " DELETE " & _
           " FROM " & T_COM002 & " " & _
           " WHERE " & DBW("cdindex=", "B003") & " " & _
           " AND   " & DBW("cdval1=", pGroup) & " " & _
           " AND   " & DBW("cdval2=", pTestCd)
              
    DBConn.Execute SSQL
    Delete = True
    Exit Function
    
Delete_error:
    Delete = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function Save(ByVal pGroup As String, ByVal pTestCd As String, ByVal pSpecCd As String, ByVal pSpecNm As String) As Boolean
    If IsExist(pGroup, pTestCd) = True Then
        Save = Update(pGroup, pTestCd, pSpecCd, pSpecNm)
    Else
        Save = Insert(pGroup, pTestCd, pSpecCd, pSpecNm)
    End If
End Function

Private Function IsExist(ByVal pGroup As String, ByVal pTestCd As String) As Boolean
    Dim SSQL As String
    Dim RS   As Recordset
    
    SSQL = " SELECT * " & _
           " FROM " & T_COM002 & " " & _
           " WHERE " & DBW("cdindex=", "B003") & " " & _
           " AND   " & DBW("cdval1=", pGroup) & " " & _
           " AND   " & DBW("cdval2=", pTestCd)
    Set RS = New Recordset
    
    RS.Open SSQL, DBConn
    
    If RS.RecordCount < 1 Then
        IsExist = False
    Else
        IsExist = True
    End If
    Set RS = Nothing
End Function

Private Function Update(ByVal pGroup As String, ByVal pTestCd As String, ByVal pSpecCd As String, ByVal pSpecNm As String) As Boolean
    Dim SSQL As String
    
    SSQL = " UPDATE " & T_COM002 & " " & _
           " SET " & DBW("field1=", pSpecCd) & "," & _
                     DBW("field2=", pSpecNm) & " " & _
           " WHERE " & DBW("cdindex=", "B003") & " " & _
           " AND   " & DBW("cdval1=", pGroup) & " " & _
           " AND   " & DBW("cdval2=", pTestCd)
On Error GoTo Update_error
    DBConn.Execute (SSQL)
    Update = True
    Exit Function
    
Update_error:
    Update = False
    MsgBox Err.Description, vbExclamation
End Function

Private Function Insert(ByVal pGroup As String, ByVal pTestCd As String, ByVal pSpecCd As String, ByVal pSpecNm As String) As Boolean
    Dim SSQL As String
    
    SSQL = "insert into " & T_COM002 & "(cdindex,cdval1,cdval2,field1,field2) " & _
             "values(" & DBV("cdindex", "B003") & "," & _
                         DBV("cdval1", pGroup) & "," & _
                         DBV("cdval2", pTestCd) & "," & _
                         DBV("field1", pSpecCd) & "," & _
                         DBV("field2", pSpecNm) & ")"
On Error GoTo Insert_error
    DBConn.Execute (SSQL)
    Insert = True
    Exit Function
    
Insert_error:
    Insert = False
    MsgBox Err.Description, vbExclamation
End Function
