VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'
'' HIS001 : ȯ�ڱ⺻������
'Public PTID         As String       '/* ȯ�� ID         */
'Public PTNM         As String       '/* ȯ�ڼ���        */
'Public SSN          As String       '/* �ֹι�ȣ        */
'Public DOB          As String       '/* �������        */
'Public SEX          As String       '/* ����            */
'Public EntDt        As String       '/* ��������        */
'Public ZIPCD        As String       '/* ������ȣ        */
'Public Addr1        As String       '/* �ּ�(1)         */
'Public Addr2        As String       '/* �ּ�(2)         */
'Public OrgAddr      As String       '/* ����            */
'Public TelHome      As String       '/* ��ȭ��ȣ(����)  */
'Public TelOffice    As String       '/* ��ȭ��ȣ(ȸ��)  */
'Public PTDIV        As String       '/* ȯ�ڱ���        */
'Public FirstDt      As String       '/* ���ʳ�����      */
'Public EMPDIV       As String       '/* ��������        */
'Public QlfyCd       As String       '/* �����ڰ�        */
'Public MomId        As String       '/* ���� ID         */
'Public RaceCd       As String       '/* �����ڵ�        */
'Public NationCd     As String       '/* �����ڵ�        */
'Public CivilFg      As String       '/* �ùαǿ���('0':��,'1':��) */
'Public GuardNm      As String       '/* ��������        */
'Public FamDiv       As String       '/* �������ּ�      */
'Public GuardAddr    As String       '/* ��ȣ���ּ�      */
'Public TmpDiv       As String       '/* ���� �ӽ�ID ���� */
'Public AgeYY        As Integer      '/* ���� ---> ������Ϸ� ��� */
'Public AGE          As Integer      '/* ���� ---> ������Ϸ� ��� */
'Public AgeDiv       As String       '/* ���̴��� (Y,M,D) */
'Public SexNm        As String       '/* ���� �ѱ۸� (M:����,F:����,U:�߼�) */
'Public SexAge       As String       '/* ���� �ѱ۸� (M:����,F:����,U:�߼�) */
'
'' HIS002 : ���ȯ�ڸ�����
'Public InPatientFg  As Boolean      '/* ����ȯ�ڱ��� */
'Public HIS002EXIST  As Boolean      '/* �������Ÿ���翩�� */
'Public BEDINDT      As String       '/* �Կ���          */
'Public BEDINTM      As String       '/* �Կ��ú�        */
'Public BEDOUTDT     As String       '/* �����          */
'Public BEDOUTTM     As String       '/* ����ú�        */
'Public DEPTCD       As String       '/* ������ڵ�      */
'Public MAJDOCT      As String       '/* ��ġ�� �ڵ�     */
'Public WARDID       As String       '/* �����ڵ�        */
'Public ROOMID       As String       '/* �����ڵ�        */
'Public BEDID        As String       '/* �����ڵ�        */
'Public BEDPATH      As String       '/* �Կ����        */
'Public InDiseaCd    As String       '/* �Կ����ڵ�    */
'Public PtntOrderDiv As String       '/* ó�汸��
'
'Public DEPTNM       As String       '/* �������
'Public DOCTNM       As String       '/* ��ġ�Ǹ�
'Public EMPNM        As String       '/* ������
'
''Public objSE        As Object       'Session
''Public objDb        As Connection   'Database
'
'Private objTables   As New clsTables
'Private objFields   As New clsFields
'
'Public Function PtntQuery(ByVal strPtId As String) As Boolean
''-- Parameter�� ���� Sql�� �����ϰ�, �� �ʵ��� ����
''-- Ŭ���� clsPatient�� Data Attribute�� �����Ѵ�.
'
'    Dim SqlStmt As String
'     'FROM HIS001 - �ܷ�ȯ�� ����
'    SqlStmt = " SELECT " & objFields.F_PTID & " as ptid, " & objFields.F_PTNM & " as ptnm, " & _
'                           objFields.F_SSN2 & " as ssn,  " & objFields.F_DOB & " as dbo, " & _
'                           objFields.F_SEX2 & " as sex,  " & objFields.F_ZIPCODE & " as zip_code1, " & _
'                           objFields.F_ADDRESS & " as address, " & objFields.F_TEL & " as tel, " & objFields.F_TMPDIV & " as tmpdiv " & _
'              " FROM   " & objTables.T_HIS001 & " " & _
'              " WHERE  " & DBW(objFields.F_PTID, strPtId, 2)
'    PtntQuery = HIS001READ(SqlStmt)
'    'FROM HIS002 - ����ȯ�� ����  --> ������̺��� �ִ� ��� ����.
'    '(h7lab501 ����ϱ�� �� : 2001.1.17 kmk)
'    SqlStmt = " SELECT a." & objFields.F_PTID & " as ptid," & objFields.F_BEDINDT & " as bedindt, '' as bedintm, " & _
'              "    " & objFields.F_BEDOUTDT & " as bedoutdt, '' as bedouttm, " & _
'              "        a." & objFields.F_PTDEPTCD & " as deptcd, a." & objFields.F_MAJDOCT & " as majdoct, " & _
'              "        a." & objFields.F_PTWARDID & " as wardid, a." & objFields.F_PTROOMID & " as hosilid, '' as bedid, " & _
'              "        '' as bedpath, '' as indiseacd, " & _
'              "        b." & objFields.F_DEPTNM & " as deptnm, c." & objFields.F_DOCTNM & " as doctnm " & _
'              " FROM " & objTables.T_HIS002 & " a, " & _
'                        objTables.T_HIS003 & " b, " & _
'                        objTables.T_HIS005 & " c  " & _
'              " WHERE  " & DBW(" a." & objFields.F_PTID, strPtId, 2) & _
'              " AND    b." & objFields.F_DEPTCD & " = a." & objFields.F_PTDEPTCD & _
'              " AND    c." & objFields.F_DOCTID & " = a." & objFields.F_MAJDOCT & _
'              " ORDER BY bedindt DESC"
'    HIS002EXIST = HIS002READ(SqlStmt)
'    If HIS002EXIST Then
'        InPatientFg = IIf(BEDOUTDT = "", True, False)
'    End If
'End Function
'
'Private Sub Class_Initialize()
'    Call ClearData
'End Sub
'
'Private Sub Class_Terminate()
'    Set objTables = Nothing
'    Set objFields = Nothing
'End Sub
'
'
'Private Function HIS001READ(ByVal tmpSql As String) As Boolean
''-- HIS001 : �ܷ�ȯ�� �⺻����
'    Dim objRS As Recordset
'
'    Call ClearData
'
'On Error GoTo Error_Trap
'    Set objRS = New Recordset
'    objRS.Open tmpSql, DbConn
'    If objRS.EOF Then
'        HIS001READ = False
'        GoTo NoData
'    Else
'        HIS001READ = True
'    End If
'    With objRS
'        PTID = Trim("" & .Fields("ptid").Value)
'        PTNM = Trim("" & .Fields("ptnm").Value)
'        SSN = Trim("" & .Fields("ssn").Value)    '& .Fields("sujumin2").Value)
'        DOB = Trim("" & .Fields("dbo").Value)
'        If Not IsDate(Format(DOB, "####-##-##")) Then
'            If IsDate(Format(SSN, "####-##-##")) Then
'                DOB = SSN
'            Else
'                If Len(Mid(DOB, 1, 4)) = 4 Then
'                    DOB = Mid(DOB, 1, 4) & "0101"
'                ElseIf Len(Mid(SSN, 1, 4)) = 4 Then
'                    DOB = Mid(SSN, 1, 4) & "0101"
'                Else
'                    DOB = Format(Now, CS_DateDbFormat)
'                End If
'            End If
'        End If
'        SEX = Trim("" & .Fields("sex").Value)
'        If IsNumeric(SEX) Then
'            SEX = Choose((Val(SEX) Mod 2) + 1, "F", "M")
'        End If
'        ZIPCD = Trim("" & .Fields("zip_code1").Value)
'        Addr1 = Trim("" & .Fields("address").Value)
'        TelHome = Trim("" & .Fields("tel").Value)
'        TmpDiv = Trim(.Fields("tmpdiv").Value & "")
'
'        Call GetAge(DOB, AGE, AgeDiv)
'        Call GetSex(SEX, SexNm)
'        SexAge = SexNm & " / " & AGE & AgeDiv
'
'        If Len(DOB) = 4 Then DOB = DOB & "0101"
'        If Len(DOB) = 6 Then DOB = DOB & "01"
'    End With
'NoData:
''   objRS.RsClose
'   Set objRS = Nothing
'   Exit Function
'Error_Trap:
'   If Err.Number <> 94 Then
'      MsgBox Err.Number & "  " & Err.Description
'      Exit Function
'   Else
'      Resume Next
'   End If
'
'End Function
'
'
'Private Function HIS002READ(ByVal tmpSql As String) As Boolean
''-- HIS002 : �Կ�ȯ�� �⺻����
'    Dim objRS As Recordset
'
'On Error GoTo Error_Trap
'    Set objRS = New Recordset
'    objRS.Open tmpSql, DbConn
'    If objRS.EOF Then
'        HIS002READ = False
'        GoTo NoData
'    Else
'        HIS002READ = True
'    End If
'
'    With objRS
'        BEDINDT = "" & .Fields("bedindt").Value
'        BEDINTM = "" & .Fields("bedintm").Value
'        BEDOUTDT = "" & .Fields("bedoutdt").Value
'        BEDOUTTM = "" & .Fields("bedouttm").Value
'        DEPTCD = "" & .Fields("deptcd").Value
'        MAJDOCT = "" & .Fields("majdoct").Value
'        WARDID = "" & .Fields("wardid").Value
'        ROOMID = "" & .Fields("hosilid").Value
'        BEDID = "" & .Fields("bedid").Value
'        InDiseaCd = "" & .Fields("indiseacd").Value
'        DEPTNM = "" & .Fields("deptnm").Value
'        DOCTNM = "" & .Fields("doctnm").Value
'    End With
'
'NoData:
' '   objRS.RsClose
'    Set objRS = Nothing
'    Exit Function
'Error_Trap:
'    If Err.Number <> 94 Then
'        MsgBox Err.Number & "  " & Err.Description
'        Exit Function
'    Else
'        Resume Next
'    End If
'
'End Function
'
'Public Sub GetAge(ByVal strDOB As String, ByRef intAge As Integer, ByRef strAgeDiv As String)
''% ������ϰ� �������� �������� ����(����,�Ϸ�)�� ���Ѵ�.
'
'    Dim tmpAge As Integer
'
'On Error GoTo Err_Trap
'
'    If Len(strDOB) = 6 Then strDOB = strDOB & "01"
'    If Not IsDate(Format(strDOB, CS_DateMask)) Then strDOB = Mid(strDOB, 1, 4) & "0101"
'
'    strAgeDiv = "Y"
'    intAge = DateDiff("YYYY", Format(strDOB, "####-##-##"), DbConn.GetSysDate)
'    AgeYY = intAge  '����
'    If intAge = 0 Then
'       strAgeDiv = "M"
'       intAge = DateDiff("M", Format(strDOB, "####-##-##"), DbConn.GetSysDate)   '����
'       If intAge < 6 Then
'          strAgeDiv = "D"
'          intAge = DateDiff("D", Format(strDOB, "####-##-##"), DbConn.GetSysDate)   '�Ϸ�
'       End If
'    End If
'    Exit Sub
'
'Err_Trap:
'    intAge = 1
'    strAgeDiv = "Y"
'
'End Sub
'
'Public Sub GetSex(ByVal strSex As String, ByRef strSexNm As String)
'      Select Case strSex
'         Case "M": strSexNm = "����"
'         Case "F": strSexNm = "����"
'         Case "U": strSexNm = "�߼�"
'         Case Else: strSexNm = strSex
'      End Select
'End Sub
'
'Public Function GetBedinDt(ByVal pPtid As String) As String
'
'    Dim sSql   As String
'    Dim objRS    As Recordset
'
'   sSql = " select " & objFields.F_BEDINDT & " as bedindt from " & objTables.T_HIS002 & _
'            " where " & objFields.F_PTID & " = '" & pPtid & "'"
'
'    Set objRS = New Recordset
'    objRS.Open sSql, DbConn
'    If objRS.EOF Then
'    GetBedinDt = ""
'    Else
'    GetBedinDt = objRS.Fields("bedindt").Value
'    End If
' '   objRS.RsClose
'    Set objRS = Nothing
'End Function
'
'
'Public Sub PatientSearch(ByRef lstPtList As Object, ByVal SearchKey As String, _
'                         ByVal SortOption As Integer, Optional ByVal strWardId As String = "")
''% ȯ��Id �Ǵ� �������� ȯ�ڸ����͸� �˻��Ѵ�.
'
'    Dim i        As Integer
'    Dim SqlStmt  As String
'    Dim tmpPtId  As String
'    Dim tmpPtNm  As String
'    Dim tmpSSN   As String
'    Dim tmpRs    As Recordset
'    Dim objSQL   As clsHosComSQLStmt
'
'    On Error GoTo Err_Trap
'    Set objSQL = New clsHosComSQLStmt
'
'    lstPtList.Clear
'    SqlStmt = objSQL.SqlPtntSearch(SortOption, SearchKey, strWardId)   '1:ȯ��ID ��, 2:ȯ�ڸ���
'    Set tmpRs = New Recordset
'    tmpRs.Open SqlStmt, DbConn
''    While (Not tmpRs.EOF)
'    Do Until tmpRs.EOF
'        tmpPtId = tmpRs.Fields("PtId").Value
'        tmpPtNm = tmpRs.Fields("PtNm").Value
'        tmpSSN = tmpRs.Fields("SSN").Value
'        lstPtList.AddItem tmpPtId & Space(11 - Len(tmpPtId)) & tmpPtNm & " (" & tmpSSN & ")"
'        tmpRs.MoveNext
'    Loop
'
''    tmpRs.RsClose
'    Set tmpRs = Nothing
'    Set objSQL = Nothing
'    Exit Sub
'Err_Trap:
'    Set tmpRs = Nothing
'    Set objSQL = Nothing
'End Sub
'
'
'Public Sub ClearData()
'
'      PTID = "":      PTNM = "":      SSN = "":       DOB = ""
'      SEX = "":       EntDt = "":     ZIPCD = "":     Addr1 = ""
'      Addr2 = "":     OrgAddr = "":   TelHome = "":   TelOffice = ""
'      PTDIV = "":     FirstDt = "":   EMPDIV = "":    QlfyCd = ""
'      MomId = "":     RaceCd = "":    NationCd = "":  CivilFg = ""
'      GuardNm = "":   FamDiv = "":    GuardAddr = "": AGE = 0
'      AgeDiv = "":    SexNm = "":     BEDINDT = "":   BEDINTM = ""
'      BEDOUTDT = "":  BEDOUTTM = "":  DEPTCD = "":    MAJDOCT = ""
'      WARDID = "":    ROOMID = "":    BEDID = "":     BEDPATH = ""
'      InDiseaCd = "": DEPTNM = "":    DOCTNM = ""
'
'      HIS002EXIST = False: InPatientFg = False
'
'End Sub
'
'Public Function GetDeptNm(ByVal pDeptCd As String)
''-- �μ��� �о����
'
'    Dim sSql    As String
'    Dim objRS   As Recordset
'    Dim objSQL  As clsHosComSQLStmt
'
'    Set objSQL = New clsHosComSQLStmt
'
'    GetDeptNm = ""
'    If pDeptCd = "" Then GoTo ExitPos
'
'    sSql = objSQL.SqlDeptList(pDeptCd)
'    Set objRS = New Recordset
'    objRS.Open sSql, DbConn
'    If objRS.EOF Then
'       GetDeptNm = ""
'    Else
'       GetDeptNm = Trim("" & objRS.Fields("DeptNm").Value)
'    End If
' '   objRS.RsClose
'
'ExitPos:
'    Set objRS = Nothing
'    Set objSQL = Nothing
'
'End Function
'
'Public Function GetDoctNm(ByVal pDoctId As String) As Boolean
'    Dim objRS   As Recordset
'    Dim objSQL  As clsHosComSQLStmt
'
'    Set objSQL = New clsHosComSQLStmt
'    Set objRS = New Recordset
'    objRS.Open objSQL.SqlDoctList(pDoctId), DbConn
'
'    If objRS.EOF = False Then
'        DOCTNM = objRS.Fields("EmpNm").Value & ""
'        GetDoctNm = True
'    Else
'        GetDoctNm = False
'    End If
' '   objRS.RsClose
'
'    Set objSQL = Nothing
'    Set objRS = Nothing
'End Function
'
'Public Function GetEmpNm(ByVal pEmpId As String, ByVal pDeptCd As String) As Boolean
'    Dim objRS   As Recordset
'    Dim objSQL  As clsHosComSQLStmt
'
'    Set objSQL = New clsHosComSQLStmt
'    Set objRS = New Recordset
'    objRS.Open objSQL.SqlEmpList(pEmpId, pDeptCd), DbConn
'    If objRS.EOF = False Then
'        EMPNM = objRS.Fields("empnm").Value & ""
'        GetEmpNm = True
'    Else
'        GetEmpNm = False
'    End If
' '   objRS.RsClose
'    Set objRS = Nothing
'    Set objSQL = Nothing
'End Function
'
'