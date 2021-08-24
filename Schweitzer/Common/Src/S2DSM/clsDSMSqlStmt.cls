VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDSMSqlStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+--------------------------------------------------------------------------------------+
'|  1. cls ��   : clsDSMSqlStmt
'|  2. ��  ��   : �� List ���,����,����
'|  3. �ۼ���   : �� ����
'|  4. �ۼ���   : 2000.10.23
'|
'|  CopyRight(C) 2000 ��ÿ�Ƽ����
'+--------------------------------------------------------------------------------------+
Option Explicit
'�Ӽ����� ���� ���� ����
Private mvarGroupID As String
Private mvarGroupNm As String
Private mvarGroupDesc As String
Private mvarUserFg As String
Private mvarAPSFg As String
Private mvarBBSFg As String
Private mvarLISFg As String

Public Property Let LISFg(ByVal vData As String)
    mvarLISFg = vData
End Property

Public Property Get LISFg() As String
    LISFg = mvarLISFg
End Property

Public Property Let BBSFg(ByVal vData As String)
    mvarBBSFg = vData
End Property

Public Property Get BBSFg() As String
    BBSFg = mvarBBSFg
End Property

Public Property Let APSFg(ByVal vData As String)
    mvarAPSFg = vData
End Property

Public Property Get APSFg() As String
    APSFg = mvarAPSFg
End Property

Public Property Let UserFg(ByVal vData As String)
    mvarUserFg = vData
End Property

Public Property Get UserFg() As String
    UserFg = mvarUserFg
End Property

Public Property Let GroupDesc(ByVal vData As String)
    mvarGroupDesc = vData
End Property

Public Property Get GroupDesc() As String
    GroupDesc = mvarGroupDesc
End Property

Public Property Let GroupNm(ByVal vData As String)
    mvarGroupNm = vData
End Property

Public Property Get GroupNm() As String
    GroupNm = mvarGroupNm
End Property

Public Property Let GroupID(ByVal vData As String)
    mvarGroupID = vData
End Property

Public Property Get GroupID() As String
    GroupID = mvarGroupID
End Property

Public Function Set_COM007(ByVal blnChk As Boolean, DFg As String, ByVal FID As String, ByVal FNm As String, _
                        ByVal FDesc As String, ByVal RFg As String, ByVal WFg As String, ByVal PFg As String) As Boolean
    Dim strSQL() As String
    Dim strFlag As String
    Dim i As Long
    
    '��� ��Ű��.
    If blnChk = True Then
        ReDim strSQL(1)
        strSQL(1) = "INSERT INTO " & T_COM007 & "(deptfg,formid,formnm,formdesc,readfg,writefg,printfg) " & _
                 " VALUES(" & DBV("deptfg", DFg) & " , " & DBV("formid", FID) & " , " & DBV("formnm", FNm) & " , " & _
                 DBV("formdesc", FDesc) & " , " & DBV("readfg", RFg) & " , " & DBV("writefg", WFg) & " , " & _
                 DBV("printfg", PFg) & " )"
    Else
    '���� ��Ű��.
        ReDim strSQL(1)
        strSQL(1) = "UPDATE " & T_COM007 & " SET formnm='" & FNm & "', formdesc='" & FDesc & "', " & _
                " " & DBW("readfg=", RFg) & " , " & DBW("writefg=", WFg) & " , " & DBW("printfg=", PFg) & " " & _
                " WHERE " & DBW("deptfg=", DFg) & _
                " and " & DBW("formid=", FID)
                
'Modify By legends 2003/09/29
'�� �� ������ �� X�� �ھ��ִ��� �𸣰ڳ�.. �׳� �� �ھ��ָ� �Ǵµ�..

'������ ����-------------------------------------------------------------------------------------
        ReDim Preserve strSQL(UBound(strSQL) + 1)
        
        strSQL(UBound(strSQL)) = " update " & T_COM009 & " set " & DBW("readfg=", "", 1) & _
                                 DBW("writefg=", "", 1) & DBW("printfg=", "") & _
                                 " where  " & DBW("deptfg =", DFg) & _
                                 " and    " & DBW("formid =", FID)
'������ ��-----------------------------------------------------------------------------------------

'������ ���� ---------------------------------------------------------------------------------------
'        'Read Flag
'        strFlag = IIf(RFg = 0, "X", "0")
'        ReDim Preserve strSQL(UBound(strSQL) + 1)
'        strSQL(UBound(strSQL)) = "update " & T_COM009 & " set readfg = '" & strFlag & "' " & _
'                                 "where  " & DBW("deptfg =", DFg) & _
'                                 "and    " & DBW("formid =", FID)
'        If strFlag = "0" Then strSQL(UBound(strSQL)) = strSQL(UBound(strSQL)) & "and readfg = 'X'"
'        'Write Flag
'        strFlag = IIf(WFg = 0, "X", "0")
'        ReDim Preserve strSQL(UBound(strSQL) + 1)
'        strSQL(UBound(strSQL)) = "update " & T_COM009 & " set writefg = '" & strFlag & "' " & _
'                                 "where  " & DBW("deptfg =", DFg) & _
'                                 "and    " & DBW("formid =", FID)
'        If strFlag = "0" Then strSQL(UBound(strSQL)) = strSQL(UBound(strSQL)) & "and writefg = 'X'"
'        'Print Flag
'        strFlag = IIf(PFg = 0, "X", "0")
'        ReDim Preserve strSQL(UBound(strSQL) + 1)
'        strSQL(UBound(strSQL)) = "update " & T_COM009 & " set  " & DBW("printfg =", strFlag) & " " & _
'                                 " where  " & DBW("deptfg =", DFg) & _
'                                 " and    " & DBW("formid =", FID)
'        If strFlag = "0" Then strSQL(UBound(strSQL)) = strSQL(UBound(strSQL)) & "and printfg = 'X'"
'������ ��---------------------------------------------------------------------------------------------
    End If
             
On Error GoTo Set_COM007_Error
    DBConn.BeginTrans
    For i = 1 To UBound(strSQL)
        DBConn.Execute strSQL(i)
    Next
    DBConn.CommitTrans
    Set_COM007 = True
    Exit Function

Set_COM007_Error:
    DBConn.RollbackTrans
    Set_COM007 = False
    MsgBox Err.Description
End Function

Public Function Del_COM007(ByVal DFg As String, ByVal FID As String) As Boolean
    Dim strSQL As String
    
    '������Ű��.
    strSQL = " DELETE FROM " & T_COM007 & " WHERE " & DBW("deptfg=", DFg) & _
             " AND " & DBW("formid=", FID)
             
On Error GoTo Del_COM007_Error
    DBConn.BeginTrans
    DBConn.Execute strSQL
    DBConn.CommitTrans
    Del_COM007 = True
    Exit Function

Del_COM007_Error:
    DBConn.RollbackTrans
    Del_COM007 = False
    MsgBox Err.Description, "�������"
End Function

'Public Function Get_COM007(ByVal DFg As String, ByVal FID As String) As Recordset
'    Dim strSQL As String
'
'    '��ϵ� ������ ��������.
'    strSQL = "SELECT * FROM " & T_COM007 & _
'             " WHERE " & DBW("deptfg=", DFg) & _
'             " AND " & DBW("formid=", FID)
'
'    Set Get_COM007 = New Recordset
'    Get_COM007.Open strSQL, dbconn
'
''    If Get_COM007.DBerror = True Then
''        MsgBox dbconn.Errors.Item(1).Description
''        Exit Function
''    End If
'End Function

Public Function GetSQLCOM007(ByVal DFg As String, ByVal FID As String) As String
    Dim strSQL As String
    
    '��ϵ� ������ ��������.
    strSQL = "SELECT * FROM " & T_COM007 & _
             " WHERE " & DBW("deptfg=", DFg) & _
             " AND " & DBW("formid=", FID)
    GetSQLCOM007 = strSQL
End Function

'Public Function Listview_COM007(ByVal DFg As String) As Recordset
'    Dim strSQL As String
'
'    '����Ʈ �� ������ ������.
'    strSQL = "SELECT * FROM " & T_COM007 & _
'             " WHERE " & DBW("deptfg=", DFg)
'
'    Set Listview_COM007 = New Recordset
'    Listview_COM007.Open strSQL, dbconn
'
''    If Listview_COM007.DBerror = True Then
''        MsgBox dbconn.Errors.Item(1).Description
''        Exit Function
''    End If
'End Function

Public Function ShowListView(ByRef pListView As Object, ByVal DFg As String)
    '����Ʈ display
    Dim Rs As New Recordset
    Dim strSQL As String
    Dim itmX As Object
    Dim strDeptFg As String
    Dim strRead As String
    Dim strWrite As String
    Dim strPrint As String
    
    strSQL = "SELECT * FROM " & T_COM007 & " WHERE " & DBW("deptfg=", DFg)
    Rs.Open strSQL, DBConn
    If Rs.EOF = False Then
        With pListView
            .ListItems.clear
            Do Until Rs.EOF
                Select Case "" & Rs.Fields("deptfg").Value
                    Case "L": strDeptFg = "LIS"
                    Case "A": strDeptFg = "APS"
                    Case "B": strDeptFg = "BBS"
                End Select
        
                Select Case "" & Rs.Fields("readfg").Value
                    Case "0": strRead = "����"
                    Case "1": strRead = "����"
                End Select
        
                Select Case "" & Rs.Fields("writefg").Value
                    Case "0": strWrite = "����"
                    Case "1": strWrite = "����"
                End Select
        
                Select Case "" & Rs.Fields("printfg").Value
                    Case "0": strPrint = "����"
                    Case "1": strPrint = "����"
                End Select
                Set itmX = .ListItems.Add(, , strDeptFg)
                itmX.SubItems(1) = "" & Rs.Fields("formid").Value
                itmX.SubItems(2) = "" & Rs.Fields("formnm").Value
                itmX.SubItems(3) = "" & Rs.Fields("formdesc").Value
                itmX.SubItems(4) = strRead
                itmX.SubItems(5) = strWrite
                itmX.SubItems(6) = strPrint
                Rs.MoveNext
            Loop
        End With
    End If
    Set Rs = Nothing
End Function

Public Function ShowListView_all(ByRef pListView As Object)
    '����Ʈ display
    Dim Rs As New Recordset
    Dim strSQL As String
    Dim itmX As Object
    Dim strDeptFg As String
    Dim strRead As String
    Dim strWrite As String
    Dim strPrint As String
    
    strSQL = "SELECT * FROM " & T_COM007
    Rs.Open strSQL, DBConn
    If Rs.EOF = False Then
        With pListView
            .ListItems.clear
            Do Until Rs.EOF
                Select Case "" & Rs.Fields("deptfg").Value
                    Case "L": strDeptFg = "LIS"
                    Case "A": strDeptFg = "APS"
                    Case "B": strDeptFg = "BBS"
                End Select
                Select Case "" & Rs.Fields("readfg").Value
                    Case "0": strRead = "����"
                    Case "1": strRead = "����"
                End Select
        
                Select Case "" & Rs.Fields("writefg").Value
                    Case "0": strWrite = "����"
                    Case "1": strWrite = "����"
                End Select
        
                Select Case "" & Rs.Fields("printfg").Value
                     Case "0": strPrint = "����"
                     Case "1": strPrint = "����"
                End Select
                
                Set itmX = .ListItems.Add(, , strDeptFg)
                itmX.SubItems(1) = "" & Rs.Fields("formid").Value
                itmX.SubItems(2) = "" & Rs.Fields("formnm").Value
                itmX.SubItems(3) = "" & Rs.Fields("formdesc").Value
                itmX.SubItems(4) = strRead
                itmX.SubItems(5) = strWrite
                itmX.SubItems(6) = strPrint
                Rs.MoveNext
            Loop
        End With
    End If
    Set Rs = Nothing
End Function

'
'Here is Legends's Start
'
'Coding By Legends
'Conding Date 2k/10
'���������� ���, �׷��Ͽ��� ���Ǵ� Ŭ����
'frmDSM002, frmDSM003 ������ ���Ǵ� Ŭ����
'��������, ����Ʈ��, COM006, COM008, COM009 �������, �������� ���


Public Function SetCOM006(ByVal updateFg As Boolean, ByVal EmpId As String, _
                          ByVal EmpLngNm As String, ByVal EmpShtNm As String, _
                          ByVal Sex As String, ByVal DOB As String, _
                          ByVal SSN As String, ByVal ShiftCd As String, _
                          ByVal DeptCd As String, ByVal HireDt As String, _
                          ByVal EntDt As String, ByVal FireDt As String, _
                          ByVal Degree As String, ByVal TechFg As String, _
                          ByVal TelNo As String, ByVal CellNo As String, _
                          ByVal Remark As String, ByVal FireCheck As Boolean _
                          )
'COM006�� ����

    Dim strSQL As String
        
    If FireCheck Then
        strSQL = " Update " & T_COM006 & " Set " & DBW("empid =", EmpId) & " , " & DBW("empnm =", EmpLngNm) & " , " & _
                 DBW("empsnm =", EmpShtNm) & " ,  " & DBW("sex =", Sex) & " , " & DBW("dob =", Format(DOB, "yyyyMMdd")) & " , " & _
                 DBW("ssn =", SSN) & " , " & DBW("shiftcd =", ShiftCd) & " , " & DBW("deptcd =", DeptCd) & " , " & _
                 DBW("hiredt =", Format(HireDt, "yyyyMMdd")) & " , " & DBW("entdt =", Format(EntDt, "yyyyMMdd")) & " , " & _
                 DBW("firedt =", Format(FireDt, "yyyyMMdd")) & " , " & DBW("degree =", Degree) & " , " & _
                 DBW("techfg =", TechFg) & " , " & DBW("telno =", TelNo) & " , " & DBW("cellno =", CellNo) & " , " & _
                 DBW("note =", Remark) & _
                 "Where  " & DBW("empid =", EmpId)
    Else
        If updateFg = False Then
            strSQL = "Insert Into " & T_COM006 & " (empid, empnm, empsnm, sex, dob, ssn, shiftcd, deptcd, hiredt, " & _
                     " entdt, firedt, degree, techfg, telno, cellno, note) Values(" & _
                     DBV("empid", EmpId) & " , " & DBV("empnm", EmpLngNm) & " ,  " & DBV("empsnm", EmpShtNm) & " , " & _
                     DBV("sex", Sex) & " , " & DBV("dob", Format(DOB, "yyyyMMdd")) & " , " & DBV("ssn", SSN) & " , " & _
                     DBV("shiftcd", ShiftCd) & " , " & DBV("deptcd", DeptCd) & " , " & DBV("hiredt", Format(HireDt, "yyyyMMdd")) & " , " & _
                     DBV("entdt", Format(EntDt, "yyyyMMdd")) & " , " & DBV("firedt", "") & " ,  " & DBV("degree", Degree) & ", " & _
                     DBV("techfg", TechFg) & " , " & DBV("telno", TelNo) & " , " & DBV("cellno", CellNo) & " , " & DBV("remark", Remark) & " )"
        Else
            strSQL = "Update " & T_COM006 & " Set  " & DBW("empid =", EmpId) & " , " & DBW("empnm =", EmpLngNm) & " , " & DBW("empsnm =", EmpShtNm) & _
                 " , " & DBW("sex =", Sex) & " , " & DBW("dob =", Format(DOB, "yyyyMMdd")) & " , " & DBW("ssn =", SSN) & " , " & DBW("shiftcd =", ShiftCd) & " , " & _
                 DBW("deptcd =", DeptCd) & " , " & _
                 DBW("hiredt =", Format(HireDt, "yyyyMMdd")) & " , " & DBW("entdt =", Format(EntDt, "yyyyMMdd")) & _
                 " , " & DBW("firedt =", "") & " , " & DBW("degree =", Degree) & ", " & DBW("techfg =", TechFg) & _
                 " , " & DBW("telno =", TelNo) & " , " & DBW("cellno =", CellNo) & " , " & DBW("note =", Remark) & _
                 " Where  " & DBW("empid =", EmpId)
        End If
    End If
    
On Error GoTo errSetCOM006
    DBConn.BeginTrans
    DBConn.Execute strSQL
    DBConn.CommitTrans
    MsgBox "����Ǿ����ϴ�.", vbInformation, "��������"
    Exit Function
    
errSetCOM006:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation

End Function

Public Function SetCOM098(ByVal updateFg As Boolean, ByVal EmpId As String, _
                          ByVal EmpLngNm As String, ByVal DeptCd As String, ByVal DeptNm As String, _
                          ByVal EntDt As String, ByVal TelNo As String, ByVal CellNo As String, _
                          ByVal Remark As String, ByVal FireCheck As Boolean _
                          )
'COM098�� ����

    Dim strSQL As String
        
    If FireCheck Then
            strSQL = ""
            strSQL = strSQL & "UPDATE S2COM098"
            strSQL = strSQL & "   SET EMPNM = '" & EmpLngNm & "' , "
            strSQL = strSQL & "       DEPTCD = '" & DeptCd & "' , "
            strSQL = strSQL & "       DEPTNM = '" & DeptNm & "' , "
            strSQL = strSQL & "       EDITDT = '" & Format(Now, "YYYY-MM-DD") & "' , "
            strSQL = strSQL & "       TELNO = '" & TelNo & "' , "
            strSQL = strSQL & "       REMARK = '" & Remark & "'"
            strSQL = strSQL & " WHERE EMPNO = '" & EmpId & "'"
    Else
        If updateFg = False Then
            strSQL = "Insert Into S2COM098 (empno, empnm, deptcd, deptnm, editdt, telno, remark) Values(" & _
                     DBV("empid", EmpId) & " , " & DBV("empnm", EmpLngNm) & " , " & _
                     DBV("deptcd", DeptCd) & " , " & DBV("deptnm", DeptNm) & " , " & _
                     DBV("entdt", Format(EntDt, "yyyyMMdd")) & " , " & _
                     DBV("telno", TelNo) & " , " & DBV("remark", Remark) & " )"
        Else
            strSQL = ""
            strSQL = strSQL & "UPDATE S2COM098"
            strSQL = strSQL & "   SET EMPNM = '" & EmpLngNm & "' , "
            strSQL = strSQL & "       DEPTCD = '" & DeptCd & "' , "
            strSQL = strSQL & "       DEPTNM = '" & DeptNm & "' , "
            strSQL = strSQL & "       EDITDT = '" & Format(Now, "YYYY-MM-DD") & "' , "
            strSQL = strSQL & "       TELNO = '" & TelNo & "' , "
            strSQL = strSQL & "       REMARK = '" & Remark & "'"
            strSQL = strSQL & " WHERE EMPNO = '" & EmpId & "'"
        End If
    End If
    
'    DBConn.BeginTrans
    DBConn.Execute strSQL
'    DBConn.CommitTrans
    MsgBox "����Ǿ����ϴ�.", vbInformation, "��������"
    Exit Function
    
errSetCOM006:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation

End Function

Public Function UpdateCheck(Optional ByVal pKey1 As Variant, Optional ByVal pKey2 As Variant, _
                            Optional ByVal pKey3 As Variant, Optional ByVal strSQL As String) As Boolean
'������Ʈ üũ

    Dim Rs As New Recordset
    
    If Not IsMissing(pKey1) Then
        Rs.Open strSQL, DBConn
        UpdateCheck = Not Rs.EOF
    End If
    Set Rs = Nothing
    
End Function

Public Function DelCOM006(ByVal EmpId As String)
'COM006 ����

    Dim strSQL As String
    
    strSQL = "Delete " & T_COM006 & " Where " & DBW("empid =", EmpId)
    
On Error GoTo ErrDelCOM006
    DBConn.BeginTrans
    DBConn.Execute strSQL
    DBConn.CommitTrans
    MsgBox "�����Ǿ����ϴ�.", vbInformation, "��������"
    Exit Function
    
ErrDelCOM006:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
    
End Function

Public Function DelCOM098(ByVal EmpId As String)
'COM098 ����

    Dim strSQL As String
    
    strSQL = "Delete S2COM098 Where EMPNO = " & Trim(EmpId)
On Error GoTo ErrDelCOM098
    DBConn.BeginTrans
    DBConn.Execute strSQL
    DBConn.CommitTrans
    MsgBox "�����Ǿ����ϴ�.", vbInformation, "��������"
    Exit Function
    
ErrDelCOM098:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
    
End Function

Public Function ShowListViewValue(ByRef pListView As Object)
'����Ʈ �信 �����ִ�

    Dim Rs As New Recordset
    Dim strSQL As String
    Dim itmX As Object
    
    strSQL = Query(2)
    
    Rs.Open strSQL, DBConn
'    If rs.DBerror Then
'        DisplayErrors
'        Exit Function
'    End If
    With pListView
        .ListItems.clear
        Do Until Rs.EOF
            Set itmX = .ListItems.Add(, , "" & Rs.Fields("EmpID").Value)
            itmX.SubItems(1) = Rs.Fields("EmpNm").Value & ""
            itmX.SubItems(2) = Rs.Fields("EmpSNm").Value & ""
            itmX.SubItems(3) = IIf(Rs.Fields("Sex").Value = "0", "��", "��")
            itmX.SubItems(4) = Format(Rs.Fields("DOB").Value & "", "####-##-##")
            itmX.SubItems(5) = Rs.Fields("SSN").Value & ""
            itmX.SubItems(6) = IIf(Rs.Fields("ShiftCd").Value = "0", "Day", "Night")
            itmX.SubItems(7) = Rs.Fields("DeptCd").Value & ""
            itmX.SubItems(8) = Format(Rs.Fields("HireDt").Value & "", "####-##-##")
            itmX.SubItems(9) = Format(Rs.Fields("EntDt").Value & "", "####-##-##")
            itmX.SubItems(10) = IIf(Rs.Fields("FireDt").Value & "" = "" Or Rs.Fields("FireDt").Value = Null, "", Format(Rs.Fields("FireDt").Value & "", "####-##-##"))
            If itmX.SubItems(10) <> "" Then
                itmX.ForeColor = DCM_Gray
                'itmX.SubItems(1) = itmX.SubItems(1) & "**"
            End If
            Select Case Rs.Fields("Degree").Value & ""
                Case "0": itmX.SubItems(11) = "Master"
                Case "1": itmX.SubItems(11) = "Engineer"
                Case "2": itmX.SubItems(11) = "Employee"
            End Select
            itmX.SubItems(12) = IIf("" & Rs.Fields("TechFg").Value = "0", "", "���")
            itmX.SubItems(13) = Rs.Fields("TelNo").Value & ""
            itmX.SubItems(14) = Rs.Fields("CellNo").Value & ""
            itmX.SubItems(15) = Rs.Fields("note").Value & ""
            Rs.MoveNext
        Loop
    End With

    Set Rs = Nothing
End Function

Public Function ShowEmpListView(ByRef pListView As Object, Optional ByVal pExpFg As Long = 0)
'����Ʈ �信 �����ִ�

    Dim Rs As New Recordset
    Dim strSQL As String
    Dim itmX As Object
    
    strSQL = Query(2)
    
    Rs.Open strSQL, DBConn
'    If rs.DBerror Then
'        DisplayErrors
'        Exit Function
'    End If
    With pListView
        .ListItems.clear
        Do Until Rs.EOF
            If (pExpFg = 0) Or _
               (pExpFg = 1 And (Rs.Fields("FireDt").Value & "") = "") Then
                Set itmX = .ListItems.Add(, , "" & Rs.Fields("EmpID").Value)
                itmX.SubItems(1) = Rs.Fields("EmpNm").Value & ""
                itmX.SubItems(2) = Rs.Fields("EmpSNm").Value & ""
                itmX.SubItems(3) = IIf(Rs.Fields("Sex").Value = "0", "��", "��")
                itmX.SubItems(4) = Format(Rs.Fields("DOB").Value & "", "####-##-##")
                itmX.SubItems(5) = Rs.Fields("SSN").Value & ""
                itmX.SubItems(6) = IIf(Rs.Fields("ShiftCd").Value = "0", "Day", "Night")
                itmX.SubItems(7) = Rs.Fields("DeptCd").Value & ""
                itmX.SubItems(8) = Format(Rs.Fields("HireDt").Value & "", "####-##-##")
                itmX.SubItems(9) = Format(Rs.Fields("EntDt").Value & "", "####-##-##")
                itmX.SubItems(10) = IIf(Rs.Fields("FireDt").Value & "" = "" Or Rs.Fields("FireDt").Value = Null, "", Format(Rs.Fields("FireDt").Value & "", "####-##-##"))
                If itmX.SubItems(10) <> "" Then
                    itmX.ForeColor = DCM_Gray
                    'itmX.SubItems(1) = itmX.SubItems(1) & "**"
                End If
                Select Case Rs.Fields("Degree").Value & ""
                    Case "0": itmX.SubItems(11) = "Master"
                    Case "1": itmX.SubItems(11) = "Engineer"
                    Case "2": itmX.SubItems(11) = "Employee"
                End Select
                itmX.SubItems(12) = IIf("" & Rs.Fields("TechFg").Value = "0", "", "���")
                itmX.SubItems(13) = Rs.Fields("TelNo").Value & ""
                itmX.SubItems(14) = Rs.Fields("CellNo").Value & ""
                itmX.SubItems(15) = Rs.Fields("note").Value & ""
                itmX.SubItems(16) = Format(.ListItems.Count, "0000")
            End If
            Rs.MoveNext
        Loop
    End With
    Set Rs = Nothing
End Function

Public Function ShowDoctListView(ByRef pListView As Object, Optional ByVal pExpFg As Long = 0)
'����Ʈ �信 �����ִ�

    Dim Rs As New Recordset
    Dim strSQL As String
    Dim itmX As Object
    
    strSQL = Query(7)
    
    Rs.Open strSQL, DBConn
'    If rs.DBerror Then
'        DisplayErrors
'        Exit Function
'    End If
    With pListView
        .ListItems.clear
        Do Until Rs.EOF
            Set itmX = .ListItems.Add(, , "" & Rs.Fields("EmpNo").Value)
            itmX.SubItems(1) = Rs.Fields("EmpNm").Value & ""
            itmX.SubItems(2) = Rs.Fields("DeptCd").Value & ""
            itmX.SubItems(3) = Rs.Fields("DeptNm").Value & ""
            itmX.SubItems(4) = Rs.Fields("EditDt").Value & ""
            itmX.SubItems(5) = Rs.Fields("TelNo").Value & ""
            'itmX.SubItems(14) = Rs.Fields("CellNo").Value & ""
            itmX.SubItems(6) = Rs.Fields("Remark").Value & ""
       
            Rs.MoveNext
        Loop
    End With
    Set Rs = Nothing
End Function

Public Function FireCheck(ByRef pCheck As Object) As Boolean
'��� üũ

    If pCheck = "0" Then
        FireCheck = False
    Else
        FireCheck = True
    End If

End Function

Private Sub clear()
'���� �ʱ�ȭ
    mvarGroupID = ""
    mvarGroupNm = ""
    mvarGroupDesc = ""
    mvarUserFg = ""
    mvarAPSFg = ""
    mvarBBSFg = ""
    mvarLISFg = ""
End Sub

Public Function Query(ByVal pOption As Long) As String
'COM007,COM008,COM009���� ���ϴ� ���� ������
'pOption = '1':HIS ����, '2':COM006 ����, '3':COM007 ����, '4':COM007  �μ����� ����, '5':COM008 ����,'6':COM007,COM008,COM009 ����

    Select Case pOption
        Case 1
            'Con_hos�� �ִ� ���� ����ϱ� ���� �ٲ�
'            Dim objDept As New clsBasisData
            Query = GetSQLDeptList
'            Set objDept = Nothing
            
'            Query = "SELECT " & F_DEPTCD & " as deptcd, " & F_DEPTNM & " as deptnm  FROM " & T_HIS003
        Case 2
            Query = "SELECT * From " & T_COM006     '& " order by empid" ' com006"
        Case 3
            Query = "SELECT * FROM " & T_COM007 '"
        Case 4
            Query = "SELECT * FROM " & T_COM007 '"
        Case 5
            Query = "SELECT * FROM " & T_COM008 & " order by groupid "
        Case 6
            Query = "Select a.groupid, a.groupnm, a.groupdesc, a.userfg, a.apsfg, a.bbsfg, a.lisfg, " & _
                         "b.deptfg , b.formid, c.formnm, b.readfg, b.writefg, b.printfg  " & _
                         "from " & T_COM008 & " a , " & T_COM009 & " b, " & T_COM007 & " c " & _
                         "Where a.groupid = b.groupid And b.deptfg=c.deptfg And b.formid = c.formid And a.groupid = '" ' & Trim(pText) & "'"
        Case 7
            Query = "SELECT * FROM S2COM098"
    End Select
End Function
'Legends End