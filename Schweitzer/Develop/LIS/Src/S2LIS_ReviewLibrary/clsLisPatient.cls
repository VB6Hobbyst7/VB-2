VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
''+--------------------------------------------------------------------------------------+
''|  1.  Class   ��  : clsLisPatient
''|  2.  �� ��         : ȯ�� �⺻���� ������ �����Ѵ�.
''|  3. �ۼ���        : ��̰�
''|  4. �ۼ���        : 1999.06.01
''|
''|  CopyRight(C) 1999 ��ÿ�Ƽ����
''+--------------------------------------------------------------------------------------+
'
'
''// HIS001
'Public PtId  As String '/* ȯ�� ID         */
'Public PtNm As String '/* ȯ�ڼ���        */
'Public SSN As String '/* �ֹι�ȣ        */
'Public DOB As String '/* �������        */
'Public Sex As String '/* ����            */
'Public EntDt As String '/* ��������        */
'Public ZipCd As String '/* ������ȣ        */
'Public Addr1 As String '/* �ּ�(1)         */
'Public Addr2 As String '/* �ּ�(2)         */
'Public OrgAddr As String '/* ����            */
'Public TelHome As String '/* ��ȭ��ȣ(����)  */
'Public TelOffice As String '/* ��ȭ��ȣ(ȸ��)  */
'Public PtDiv As String '/* ȯ�ڱ���        */
'Public FirstDt As String '/* ���ʳ�����      */
'Public EmpDiv As String '/* ��������        */
'Public QlfyCd As String '/* �����ڰ�        */
'Public MomId As String '/* ���� ID         */
'Public RaceCd As String '/* �����ڵ�        */
'Public NationCd As String '/* �����ڵ�        */
'Public CivilFg As String '/* �ùαǿ���('0':��,'1':��) */
'Public GuardNm As String '/* ��������        */
'Public FamDiv As String '/* �������ּ�      */
'Public GuardAddr As String '/* ��ȣ���ּ�      */
'
'Public Age As Integer '/* ���� ---> ������Ϸ� ��� */
'Public AgeDiv As String '/* ���̴��� (Y,M,D) */
'Public SexNm As String '/* ���� �ѱ۸� (M:����,F:����,U:�߼�) */
'
'Public DoneFg As String
'
''// HIS002
'Public InPatientFg As Boolean '/* ����ȯ�ڱ��� */
'Public BedinDt As String '/* �Կ���          */
'Public BedInTm As String '/* �Կ��ú�        */
'Public BedOutDt As String '/* �����          */
'Public BedOutTm As String '/* ����ú�        */
'Public DeptCd As String '/* ������ڵ�      */
'Public MajDoct As String '/* ��ġ�� �ڵ�     */
'Public WardId As String '/* �����ڵ�        */
'Public RoomId As String '/* �����ڵ�        */
'Public BedId As String '/* �����ڵ�        */
'Public BedPath As String '/* �Կ����        */
'Public InDiseaCd As String '/* �Կ����ڵ�    */
'
'Public DeptNm As String '/* �������
'Public DoctNm As String '/* ��ġ�Ǹ�
'
'
'
''% Method 1 : PtntQuery
''%                 Parameter�� ���� Sql�� �����ϰ�, �� �ʵ��� ����
''%                 Ŭ���� clsLisPatient�� Data Attribute�� �����Ѵ�.
'
'Public Function PtntQuery(ByVal strPtID As String, Optional ByVal strBedinDt As String = "") As Boolean
'
'    Dim SqlStmt As String
'
'    SqlStmt = "Select distinct a.ptid, b." & F_PTNM & " as ptnm, " & F_DOB2("b") & " as dob, " & F_SSN2("b") & " as ssn, " & _
'              "       a.bedindt, a.donefg, a.deptcd, a.wardid, a.hosilid, b." & F_ADDRESS & " addr, " & _
'              "       c." & F_IKNM & " as DiagNm, d." & F_DEPTNM & " as DeptNm " & _
'              " From  " & T_HIS001 & " b, " & T_HIS006 & " c, " & T_HIS003 & " d, " & T_LAB501 & " a " & _
'              " Where " & DBW("a.ptid = ", strPtID) & _
'              " and   b." & F_PTID & " = a.ptid" & _
'              " and   " & DBJ("c." & F_ICD & "=* a.disease") & _
'              " and   " & DBJ("d." & F_DEPTCD & "=* a.deptcd")
'
'    If strBedinDt <> "" Then SqlStmt = SqlStmt & " and " & DBW("a.bedindt=", strBedinDt)
'
'    PtntQuery = HIS001READ(SqlStmt)
'
'    InPatientFg = False
'
'End Function
'
'
'Private Function HIS001READ(ByVal tmpSQL As String) As Boolean
'
'   Dim MyRs As Recordset
'
'   On Error GoTo Error_Trap
'
'   Set MyRs = New Recordset
'   MyRs.Open tmpSQL, DBConn 'Sql ����
'
'   If MyRs.EOF Then
'      HIS001READ = False
'      GoTo NoData
'   Else
'      HIS001READ = True
'   End If
'
'   With MyRs
'
'      '.MoveFirst
'
'      PtId = Trim("" & .Fields("ptid").Value)
'      PtNm = Trim("" & .Fields("ptnm").Value)
'      SSN = Trim("" & .Fields("ssn").Value)
'      Addr1 = Trim("" & .Fields("addr").Value)
'      DOB = Trim("" & .Fields("dob").Value)
'      If Not IsDate(Format(DOB, CS_DateLongMask)) Then DOB = Mid(DOB, 1, 4) & "0101"
'        Sex = Choose((Val(Mid("" & .Fields("ssn").Value, 7, 1)) Mod 2) + 1, "F", "M")
'
'      DeptCd = "" & .Fields("DeptCd").Value
'      WardId = "" & .Fields("WardId").Value
'      RoomId = "" & .Fields("HosilId").Value
'
'      DeptNm = "" & .Fields("DeptNm").Value
'      InDiseaCd = "" & .Fields("DiagNm").Value
'      BedinDt = "" & .Fields("BedinDt").Value
'
'      DoneFg = "" & .Fields("DoneFg").Value
'
'      Call GetAge(DOB, Age, AgeDiv)
'      Call GetSex(Sex, SexNm)
'
'
'        If Len(DOB) = 4 Then DOB = DOB & "01"
'
'   End With
'
'NoData:
'   Set MyRs = Nothing
'   Exit Function
'
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
'
'Private Function HIS001READ_Back(ByVal tmpSQL As String) As Boolean
'
'   Dim MyRs As Recordset
'
'   On Error GoTo Error_Trap
'
'   Set MyRs = New Recordset
'   MyRs.Open tmpSQL, DBConn 'Sql ����
'
'   If MyRs.EOF Then
'      HIS001READ_Back = False
'      GoTo NoData
'   Else
'      HIS001READ_Back = True
'   End If
'
'   With MyRs
'
'      '.MoveFirst
'
'      PtId = "" & .Fields("PtId").Value
'      PtNm = "" & .Fields("PtNm").Value
'      SSN = "" & .Fields("SSN").Value
'      DOB = "" & .Fields("DOB").Value
'      Sex = "" & .Fields("Sex").Value
'      EntDt = "" & .Fields("EntDt").Value
'      ZipCd = "" & .Fields("ZipCd").Value
'      Addr1 = "" & .Fields("Addr1").Value
'      Addr2 = "" & .Fields("Addr2").Value
'      OrgAddr = "" & .Fields("OrgAddr").Value
'      TelHome = "" & .Fields("TelHome").Value
'      TelOffice = "" & .Fields("TelOffice").Value
'      PtDiv = "" & .Fields("PtDiv").Value
'      FirstDt = "" & .Fields("FirstDt").Value
'      EmpDiv = "" & .Fields("EmpDiv").Value
'      QlfyCd = "" & .Fields("QlfyCd").Value
'      MomId = "" & .Fields("MomId").Value
'      RaceCd = "" & .Fields("RaceCd").Value
'      NationCd = "" & .Fields("NationCd").Value
'      CivilFg = "" & .Fields("CivilFg").Value
'      GuardNm = "" & .Fields("GuardNm").Value
'      FamDiv = "" & .Fields("FamDiv").Value
'      GuardAddr = "" & .Fields("GuardAddr").Value
'
'      Call GetAge(DOB, Age, AgeDiv)
'      Call GetSex(Sex, SexNm)
'
'   End With
'
'NoData:
'   Set MyRs = Nothing
'   Exit Function
'
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
''% HIS002 - �Կ�ȯ�� �⺻����
'
'Private Function HIS002READ(ByVal tmpSQL As String) As Boolean
'
'   Dim MyRs As Recordset
'
'   On Error GoTo Error_Trap
'
'   Set MyRs = New Recordset   'Sql ����
'   MyRs.Open tmpSQL, DBConn
'
'   If MyRs.EOF Then
'      HIS002READ = False
'      GoTo NoData
'   Else
'      HIS002READ = True
'   End If
'
'   With MyRs
'
'      '.MoveFirst
'
'      BedinDt = "" & .Fields("BedInDt").Value
'      BedInTm = "" & .Fields("BedInTm").Value
'      BedOutDt = "" & .Fields("BedOutDt").Value
'      BedOutTm = "" & .Fields("BedOutTm").Value
'      DeptCd = "" & .Fields("DeptCd").Value
'      MajDoct = "" & .Fields("MajDoct").Value
'      WardId = "" & .Fields("WardId").Value
'      RoomId = "" & .Fields("RoomId").Value
'      BedId = "" & .Fields("BedId").Value
'      BedPath = "" & .Fields("BedPath").Value
'      InDiseaCd = "" & .Fields("InDiseaCd").Value
'
'      DeptNm = "" & .Fields("DeptNm").Value
'      DoctNm = "" & .Fields("DoctNm").Value
'
'   End With
'
'NoData:
'   MyRs.Close
'   Set MyRs = Nothing
'   Exit Function
'
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
''% ������ϰ� �������� �������� ����(����,�Ϸ�)�� ���Ѵ�.
'
'Public Sub GetAge(ByVal strDOB As String, ByRef intAge As Integer, ByRef strAgeDiv As String)
'
'   Dim tmpAge As Integer
'
'   On Error GoTo Err_Trap
'
'   If Len(strDOB) = 4 Then strDOB = DOB & "0101"
'   If Len(strDOB) = 6 Then strDOB = DOB & "01"
'
'   If Not IsDate(Format(strDOB, CS_DateLongMask)) Then strDOB = Mid(strDOB, 1, 4) & "0101"
'
'   strAgeDiv = "Y"
'   intAge = DateDiff("YYYY", Format(strDOB, CS_DateLongMask), Now)
'   If intAge = 0 Then
'      strAgeDiv = "M"
'      intAge = DateDiff("M", Format(strDOB, CS_DateLongMask), Now)
'      If intAge < 6 Then
'         strAgeDiv = "D"
'         intAge = DateDiff("D", Format(strDOB, CS_DateLongMask), Now)
'      End If
'   End If
'   Exit Sub
'
'Err_Trap:
'    intAge = 1
'    strAgeDiv = "Y"
'
'End Sub
'
'
'
'Public Sub GetSex(ByVal strSex As String, ByRef strSexNm As String)
'
'      Select Case strSex
'         Case "M": strSexNm = "����"
'         Case "F": strSexNm = "����"
'         Case "U": strSexNm = "�߼�"
'         Case Else: strSexNm = ""
'      End Select
'
'End Sub
'
''Public Function GetBedinDt(ByVal ReceptNo As String) As String
''
''   Dim MyRs As Recordset
''
''   Set MyRs = OpenRecordSet("select " & F_BEDINDT & " from " & T_HIS002 & " where hosp_gb = '" & HosptGb & "' and recept_no = " & ReceptNo)
''
''   If MyRs.EOF Then
''        GetBedinDt = ""
''   Else
''        GetBedinDt = "" & MyRs.Fields("Adm_Ymd").Value
''   End If
''   MyRs.Close
''   Set MyRs = Nothing
'
''End Function
'
''% ȯ��Id �Ǵ� �������� ȯ�ڸ����͸� �˻��Ѵ�.
'
'Public Sub PatientSearch(ByRef lstPtList As Object, ByVal SearchKey As String, ByVal SortOption As Integer)
'
'   Dim i As Integer
'   Dim SqlStmt As String
'   Dim tmpRs As Recordset
'   Dim tmpPtId As String
'   Dim tmpPtNm As String
'   Dim tmpSSN As String
'   Dim MySql As New clsLISSqlStatement
'
'   On Error GoTo Err_Trap
'
'   lstPtList.Clear
'   If SortOption = 1 Then
'      SqlStmt = MySql.SqlPtntSearch(SearchKey, "ptid")  'ȯ��ID ������ ����
'   Else
'      SqlStmt = MySql.SqlPtntSearch(SearchKey, "ptnm") 'ȯ�ڸ������� ����
'   End If
'   Set tmpRs = New Recordset
'   tmpRs.Open SqlStmt, DBConn
'
'   medLockWindowUpdate (lstPtList.hWnd)
'   While (Not tmpRs.EOF)
'      tmpPtId = "" & tmpRs.Fields("PtId").Value
'      tmpPtNm = "" & tmpRs.Fields("PtNm").Value
'      tmpSSN = "" & tmpRs.Fields("SSN").Value
'      lstPtList.AddItem tmpPtId & Space(11 - Len(tmpPtId)) & tmpPtNm & " (" & tmpSSN & ")"
'      tmpRs.MoveNext
'   Wend
'   medLockWindowUpdate (0&)
'   Call medHorScrol(lstPtList)
'
'   tmpRs.Close
'   Set tmpRs = Nothing
'   Set MySql = Nothing
'
'   Exit Sub
'
'Err_Trap:
'   MsgBox DBConn.Errors.Item(0).Number & "   " & DBConn.Errors.Item(0).Description
'   'Resume Next
'
'End Sub
'
'
'Public Sub ClearData()
'
'      PtId = ""
'      PtNm = ""
'      SSN = ""
'      DOB = ""
'      Sex = ""
'      EntDt = ""
'      ZipCd = ""
'      Addr1 = ""
'      Addr2 = ""
'      OrgAddr = ""
'      TelHome = ""
'      TelOffice = ""
'      PtDiv = ""
'      FirstDt = ""
'      EmpDiv = ""
'      QlfyCd = ""
'      MomId = ""
'      RaceCd = ""
'      NationCd = ""
'      CivilFg = ""
'      GuardNm = ""
'      FamDiv = ""
'      GuardAddr = ""
'      Age = 0
'      AgeDiv = ""
'      SexNm = ""
'      BedinDt = ""
'      BedInTm = ""
'      BedOutDt = ""
'      BedOutTm = ""
'      DeptCd = ""
'      MajDoct = ""
'      WardId = ""
'      RoomId = ""
'      BedId = ""
'      BedPath = ""
'      InDiseaCd = ""
'      DeptNm = ""
'      DoctNm = ""
'
'End Sub
'
''Public Function GetDeptNm(ByVal DeptCd As String)
''
''   Dim tmpRs As Recordset
''   Dim MySql As New clsLISSqlStatement
''   Dim SqlStmt As String
''
''   GetDeptNm = ""
''   If DeptCd = "" Then Exit Function
''
''   SqlStmt = MySql.SqlHIS003CodeList(DeptCd)
''   Set tmpRs = New Recordset
''   tmpRs.Open SqlStmt, DBConn
''   If tmpRs.EOF Then
''      GetDeptNm = ""
''   Else
''      GetDeptNm = Trim("" & tmpRs.Fields("DeptNm").Value)
''   End If
''   tmpRs.Close
''   Set tmpRs = Nothing
''
''End Function
'
'