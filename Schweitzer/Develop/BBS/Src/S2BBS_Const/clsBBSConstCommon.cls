VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBBSConstCommon"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function InsertData(pSQL() As String, Optional ByVal blnTrans As Boolean = True) As Boolean

    Dim i As Long
    
    On Error GoTo ErrInsertData

    With DBConn
        If blnTrans Then .BeginTrans

        For i = LBound(pSQL) To UBound(pSQL)
'            Debug.Print i & " : " & pSQL(i)
            If pSQL(i) <> "" Then .Execute pSQL(i)
        Next

        If blnTrans Then .CommitTrans
        InsertData = True
        Exit Function
    End With

ErrInsertData:
    With DBConn
        If blnTrans Then .RollbackTrans
        InsertData = False
    End With
End Function

Public Function Get_TestNm(ByVal TestCd As String) As String
    Dim SSQL    As String
    Dim RS      As Recordset
    
    SSQL = " SELECT testnm " & _
             " FROM " & T_BBS001 & _
             " WHERE " & DBW("testcd", TestCd, 2)
    Set RS = New Recordset
    
    RS.Open SSQL, DBConn
    
    If Not RS.EOF = True Then
        Get_TestNm = RS.Fields("testnm").Value
    Else
        Get_TestNm = ""
    End If
    
    Set RS = Nothing

End Function

Public Function Get_CompNm(ByVal COMPOCD As String) As String
'------------------------------
'혈액제제명을 가져온다
'------------------------------
    Dim RS   As Recordset
    Dim SSQL As String
    
    SSQL = " SELECT abbrnm,keepday " & _
           " FROM " & T_BBS006 & _
           " WHERE " & _
                     "     " & DBW("compocd", COMPOCD, 2) & _
                     " AND (expdt='' or expdt is null)"
             
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If RS.EOF = True Then
        Get_CompNm = "" & COL_DIV & ""
    Else
        Get_CompNm = RS.Fields("abbrnm").Value & COL_DIV & RS.Fields("keepday").Value
    End If
    
    Set RS = Nothing
End Function

Public Function GetCom003(ByVal CDINDEX As String, Optional cdval1 As String = "") As String
    Dim SSQL As String
    
    SSQL = "SELECT * FROM " & T_COM003 & " WHERE " & DBW("cdindex=", CDINDEX) & " "
    If cdval1 <> "" Then SSQL = SSQL & " AND " & DBW("cdval1=", cdval1) & " "
    
    SSQL = SSQL & " AND (field5 is null or field5='')"
    
    GetCom003 = SSQL
    
End Function

Public Function GetCenterNm(ByVal Centercd As String) As String
    Dim RS As Recordset
    
    If Centercd = "" Then GetCenterNm = "": Exit Function
    'BC2_CENTER="C230"
    Set RS = New Recordset
    RS.Open GetCom003("C230", Centercd), DBConn
    
    If RS.EOF Then
        GetCenterNm = ""
    Else
'        If RS.RecordCount < 1 Then
'            GetCenterNm = ""
'        Else
            GetCenterNm = RS.Fields("field1").Value & ""
'        End If
    End If
    
    Set RS = Nothing
End Function

Public Function GetBranchNm(ByVal BranchCd As String) As String
    Dim RS As Recordset
    
    If BranchCd = "" Then GetBranchNm = "": Exit Function
    'BC2_BRANCH="B005"
    Set RS = New Recordset
    RS.Open GetCom003("B005", BranchCd), DBConn
    
    If RS.EOF Then
        GetBranchNm = ""
    Else
'        If RS.RecordCount < 1 Then
'            GetBranchNm = ""
'        Else
            GetBranchNm = RS.Fields("field1").Value & ""
'        End If
'        RS.RsClose
    End If
    
    Set RS = Nothing
End Function

Public Function GetLocalNm(ByVal LocalCd As String) As String
    Dim RS      As Recordset
    Dim SSQL    As String
    
    If LocalCd = "" Then GetLocalNm = "": Exit Function
    'BC2_LOCAL="B006"
    SSQL = " SELECT cdval1,field1 " & _
           " FROM " & T_COM003 & " " & _
           " WHERE " & DBW("cdindex=", "B006") & _
           " AND " & DBW("cdval1=", LocalCd)
           
    Set RS = New Recordset
    
    RS.Open SSQL, DBConn
    If RS.EOF Then
        GetLocalNm = ""
    Else
'        If RS.RecordCount < 1 Then
'            GetLocalNm = ""
'        Else
            GetLocalNm = RS.Fields("field1").Value & ""
'        End If
'        RS.RsClose
    End If
    
    Set RS = Nothing
End Function


Public Function GetLabTestNm(ByVal TestCd As String, Optional ByVal ApplyDt As String = "", Optional ByVal TestDiv As String = "") As String
    Dim SSQL As String
    Dim RS As Recordset
    
    If TestCd = "" Then GetLabTestNm = "": Exit Function
    
    If ApplyDt = "" Then ApplyDt = Format(GetSystemDate, "YYYYMMDD")
    
    SSQL = "SELECT testcd,applydt,testnm,abbrnm5,abbrnm10,testdiv " & _
           "FROM " & T_LAB001 & " " & _
           "WHERE " & DBW("testcd=", TestCd) & " " & _
           "AND applydt=(" & _
                         "SELECT max(applydt) " & _
                         "FROM " & T_LAB001 & " " & _
                         "WHERE testcd='" & TestCd & "' " & _
                         "AND " & DBW("applydt<=", ApplyDt) & " " & _
                       ")"
                       
    Set RS = New Recordset
    Call RS.Open(SSQL, DBConn)
    If RS.EOF Then
'        dbconn.DisplayErrors
        GetLabTestNm = ""
        Set RS = Nothing
        Exit Function
    End If
    
    With RS
        If .RecordCount > 0 Then
            If TestDiv = "" Then
                GetLabTestNm = .Fields("testnm").Value & ""
            Else
                If TestDiv = .Fields("testdiv").Value & "" Then
                    GetLabTestNm = .Fields("testnm").Value & ""
                Else
                    GetLabTestNm = ""
                End If
            End If
'        Else
'            GetLabTestNm = ""
        End If
'        .RsClose
    End With
    Set RS = Nothing
End Function

Public Function ReadCom003(ByVal CDINDEX As String, Optional cdval1 As String = "") As Recordset
    Dim SSQL As String
    
    SSQL = GetCom003(CDINDEX, cdval1)
    Set ReadCom003 = New Recordset
    Call ReadCom003.Open(SSQL, DBConn)
End Function
