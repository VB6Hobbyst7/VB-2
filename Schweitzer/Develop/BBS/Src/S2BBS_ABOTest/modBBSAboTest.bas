Attribute VB_Name = "modBBSAboTest"
Option Explicit

Public Const BB_WORKAREA$ = "05"

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public CODE_ABOFRONT As String
Public CODE_ABOBACK As String
Public CODE_RH As String
Public CODE_ABOSUB As String
Public CODE_RHSUB As String
Public CHECK_READ_CODE As Boolean

Public Sub LoadTestCd()
    Dim objcom003 As clsCom003
    Dim Rs As Recordset
    
    If CHECK_READ_CODE Then Exit Sub
    
    Set objcom003 = New clsCom003
    Set Rs = objcom003.OpenRecordSetDay(BC2_ABO_TEST)
    Set objcom003 = Nothing
    
    If Rs.EOF = False Then
        CODE_ABOFRONT = Rs.Fields("field1").Value & ""
        CODE_ABOBACK = Rs.Fields("field2").Value & ""
        CODE_RH = Rs.Fields("field3").Value & ""
        CODE_ABOSUB = Rs.Fields("field4").Value & ""
        CODE_RHSUB = Rs.Fields("text1").Value & ""
        
        CHECK_READ_CODE = True
    End If
    
    Set Rs = Nothing
    Set objcom003 = Nothing
End Sub

Public Sub LoadRemark(ByRef vCboRemark As Object)
    Dim objABOSql As clsABOSql
    Dim Rs As Recordset
    
    Set objABOSql = New clsABOSql
    Set Rs = objABOSql.LoadRemark
        
    vCboRemark.Clear
    vCboRemark.AddItem "(없음)"
    Do Until Rs.EOF
        vCboRemark.AddItem Rs.Fields("cdval1").Value & "" & vbTab & Rs.Fields("text1").Value & ""
        
        Rs.MoveNext
    Loop
    
    Set objABOSql = Nothing
End Sub

Public Function GetLenOfAccDt(ByVal vWorkarea As String) As Long
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select * from " & T_LAB032 & _
             " where " & DBW("cdindex=", LC3_WorkArea) & _
             " and " & DBW("cdval1=", vWorkarea)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    If Rs.EOF Then
        GetLenOfAccDt = 6
    Else
        Select Case Rs.Fields("field2").Value & ""
            Case enLabDiv.LabDiv_ByDay:       '일단위
                GetLenOfAccDt = 6
            Case enLabDiv.LabDiv_ByMonth:       '월단위
                GetLenOfAccDt = 4
            Case enLabDiv.LabDiv_ByYear:       '년단위
                GetLenOfAccDt = 2
            Case enLabDiv.LabDiv_BySpc:       '검체단위
                GetLenOfAccDt = 4
            Case Else:
                GetLenOfAccDt = 6
        End Select
    End If
    
    Set Rs = Nothing
End Function

Public Function GetSqlABOFront(ByVal vRstCd As String) As String
    Dim strSQL As String
    
    strSQL = " select cdval2,field1 from " & T_LAB031 & _
             " where cdindex='C110' " & _
             " and " & DBW("cdval1=", CODE_ABOFRONT)
    strSQL = strSQL & " and " & DBW("cdval2=", vRstCd)
    
    GetSqlABOFront = strSQL
End Function

Public Function GetSqlABOFrontList() As String
    Dim strSQL As String
    
    strSQL = " select cdval2,field1 from " & T_LAB031 & _
             " where cdindex='C110' " & _
             " and " & DBW("cdval1=", CODE_ABOFRONT)
    
    GetSqlABOFrontList = strSQL
End Function

Public Function GetSqlABOBack(ByVal vRstCd As String) As String
    Dim strSQL As String
    
    strSQL = " select cdval2,field1 from " & T_LAB031 & _
             " where cdindex='C110' " & _
             " and " & DBW("cdval1=", CODE_ABOBACK)

    strSQL = strSQL & " and " & DBW("cdval2=", vRstCd)
    
    GetSqlABOBack = strSQL
End Function

Public Function GetSqlABOBackList() As String
    Dim strSQL As String
    
    strSQL = " select cdval2,field1 from " & T_LAB031 & _
             " where cdindex='C110' " & _
             " and " & DBW("cdval1=", CODE_ABOBACK)
    
    GetSqlABOBackList = strSQL
End Function

Public Function GetSqlRh(ByVal vRstCd As String) As String
    GetSqlRh = " select cdval2,field1 from " & T_LAB031 & _
                " where cdindex='C110' " & _
                " and " & DBW("cdval1=", CODE_RH) & _
                " and " & DBW("cdval2=", vRstCd)
End Function

Public Function GetSqlRhList() As String
    GetSqlRhList = " select cdval2,field1 from " & T_LAB031 & _
                " where cdindex='C110' " & _
                " and " & DBW("cdval1=", CODE_RH)
End Function

Public Function GetSqlABOSub(Optional ByVal vRstCd As String = "") As String
    GetSqlABOSub = " select cdval2,field1 from " & T_LAB031 & _
                " where cdindex='C110' " & _
                " and " & DBW("cdval1=", CODE_ABOSUB) & _
                " and " & DBW("cdval2=", vRstCd)
End Function

Public Function GetSqlABOSubList() As String
    GetSqlABOSubList = " select cdval2,field1 from " & T_LAB031 & _
                    " where cdindex='C110' " & _
                    " and " & DBW("cdval1=", CODE_ABOSUB)
End Function

Public Function GetSqlRhSub(ByVal vRstCd As String) As String
    GetSqlRhSub = " select cdval2,field1 from " & T_LAB031 & _
                " where cdindex='C110' " & _
                " and " & DBW("cdval1=", CODE_RHSUB) & _
                " and " & DBW("cdval2=", vRstCd)
End Function

Public Function GetSqlRhSubList() As String
    GetSqlRhSubList = " select cdval2,field1 from " & T_LAB031 & _
                    " where cdindex='C110' " & _
                    " and " & DBW("cdval1=", CODE_RHSUB)
End Function

Public Function GetABOFrontNm(ByVal vRstCd As String) As String
    Dim Rs As Recordset
    
    'ABO결과---------------------
    Set Rs = New Recordset
    
    Rs.Open GetSqlABOFront(vRstCd), DBConn
    
    If Rs.EOF = False Then
        GetABOFrontNm = Rs.Fields("field1").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Public Function GetABOBackNm(ByVal vRstCd As String) As String
    Dim Rs As Recordset
    
    'ABO결과---------------------
    Set Rs = New Recordset
    
    Rs.Open GetSqlABOBack(vRstCd), DBConn
    
    If Rs.EOF = False Then
        GetABOBackNm = Rs.Fields("field1").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Public Function GetABOSUBNM(ByVal vRstCd As String) As String
    Dim Rs As Recordset
    
    Set Rs = New Recordset
    Rs.Open GetSqlABOSub(vRstCd), DBConn
    
    If Rs.EOF = False Then
        GetABOSUBNM = Rs.Fields("field1").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Public Function GetRHNM(ByVal vRstCd As String) As String
    Dim Rs As Recordset
    
    Set Rs = New Recordset
    Rs.Open GetSqlRh(vRstCd), DBConn
    
    If Rs.EOF = False Then
        GetRHNM = Rs.Fields("field1").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Public Function GetRHSUBNM(ByVal vRstCd As String) As String
    Dim Rs As Recordset
    
    Set Rs = New Recordset
    Rs.Open GetSqlRhSub(vRstCd), DBConn
    
    If Rs.EOF = False Then
        GetRHSUBNM = Rs.Fields("field1").Value & ""
    End If
    
    Set Rs = Nothing
End Function

