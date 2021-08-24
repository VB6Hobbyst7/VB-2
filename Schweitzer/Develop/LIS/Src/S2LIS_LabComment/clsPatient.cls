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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+--------------------------------------------------------------------------------------+
'|  1.  Class   명  : clsPatient
'|  2.  기 능         : 환자 기본정보 내역을 생성한다.
'|  3. 작성자        : 김미경
'|  4. 작성일        : 1999.06.01
'|
'|  CopyRight(C) 1999 대련엠티에스
'+--------------------------------------------------------------------------------------+


'// HIS001
Public PtId  As String '/* 환자 ID         */
Public PtNm As String '/* 환자성명        */
Public SSN As String '/* 주민번호        */
Public DOB As String '/* 생년월일        */
Public Sex As String '/* 성별            */
Public EntDt As String '/* 생성일자        */
Public ZipCd As String '/* 우편번호        */
Public Addr1 As String '/* 주소(1)         */
Public Addr2 As String '/* 주소(2)         */
Public OrgAddr As String '/* 본적            */
Public TelHome As String '/* 전화번호(자택)  */
Public TelOffice As String '/* 전화번호(회사)  */
Public PtDiv As String '/* 환자구분        */
Public FirstDt As String '/* 최초내원일      */
Public EmpDiv As String '/* 직원구분        */
Public QlfyCd As String '/* 진료자격        */
Public MomId As String '/* 엄마 ID         */
Public RaceCd As String '/* 인종코드        */
Public NationCd As String '/* 국적코드        */
Public CivilFg As String '/* 시민권여부('0':부,'1':여) */
Public GuardNm As String '/* 가족관계        */
Public FamDiv As String '/* 연고자주소      */
Public GuardAddr As String '/* 보호자주소      */

Public Age As Integer '/* 나이 ---> 생년월일로 계산 */
Public AgeDiv As String '/* 나이단위 (Y,M,D) */
Public SexNm As String '/* 성별 한글명 (M:남자,F:여자,U:중성) */

Public DoneFg As String

'// HIS002
Public InPatientFg As Boolean '/* 병동환자구분 */
Public BedinDt As String '/* 입원일          */
Public BedInTm As String '/* 입원시분        */
Public BedOutDt As String '/* 퇴원일          */
Public BedOutTm As String '/* 퇴원시분        */
Public DeptCd As String '/* 진료과코드      */
Public MajDoct As String '/* 주치의 코드     */
Public WardId As String '/* 병동코드        */
Public RoomId As String '/* 병실코드        */
Public BedId As String '/* 병상코드        */
Public BedPath As String '/* 입원경로        */
Public InDiseaCd As String '/* 입원상병코드    */

Public DeptNm As String '/* 진료과명
Public DoctNm As String '/* 주치의명

'공용변수 - 참조하는 프로젝트로부터 설정됨
'Public MyOraSE As Object     'OraSession
Public MyDb As Object     'OraDatabase

'Private MyOraDS As Object    'ADODB.recordset



'% Method 1 : PtntQuery
'%                 Parameter로 받은 Sql을 실행하고, 각 필드의 값을
'%                 클래스 clsPatient의 Data Attribute에 저장한다.

Public Function PtntQuery(ByVal strPtId As String, Optional ByVal strBedinDt As String = "") As Boolean

   Dim SqlStmt As String
   
   SqlStmt = "Select distinct a.*, b." & F_PTNM & " as ptnm, " & F_SSN2("b") & " as ssn, b." & F_ADDRESS & ", " & _
             "      c." & F_KNM & " as DiagNm, d." & F_DEPTNM & " as DeptNm " & _
             "From " & T_LAB501 & " a, " & T_HIS001 & " b, " & T_HIS008 & " c, " & T_HIS003 & " d " & _
             "Where a.ptid = " & strPtId & " " & _
             "and   b.ptnt_no = a.ptid " & _
             "and   c.diag_cd =* a.disease " & _
             "and   d.hosp_gb =  '" & HosptGb & "' " & _
             "and   d.dept_cd =* a.deptcd "
   If strBedinDt <> "" Then SqlStmt = SqlStmt & " and a.bedindt = '" & strBedinDt & "' "
   
   PtntQuery = HIS001READ(SqlStmt)
   
   InPatientFg = False
   
End Function


Private Function HIS001READ(ByVal tmpSql As String) As Boolean
   
   Dim MyRs As Object
   
   On Error GoTo Error_Trap
   
   Set MyRs = OpenRecordSet(tmpSql)  'Sql 실행
   
   If MyRs.EOF Then
      HIS001READ = False
      GoTo NoData
   Else
      HIS001READ = True
   End If
   
   With MyRs
      
      '.MoveFirst
         
      PtId = Trim("" & .Fields("ptid").Value)
      PtNm = Trim("" & .Fields("ptnm").Value)
      SSN = Trim("" & .Fields("ssn").Value)
      Addr1 = Trim("" & .Fields("addr").Value)
      DOB = Trim(Mid("" & .Fields("ssn").Value, 1, 8))
      If Not IsDate(Format(DOB, CS_DateMask)) Then DOB = Mid(DOB, 1, 4) & "0101"
      Sex = Choose((Val(Mid("" & .Fields("ssn").Value, 9, 1)) Mod 2) + 1, "F", "M")
      
      DeptCd = "" & .Fields("DeptCd").Value
      WardId = "" & .Fields("WardId").Value
      RoomId = "" & .Fields("HosilId").Value
      
      DeptNm = "" & .Fields("DeptNm").Value
      InDiseaCd = "" & .Fields("DiagNm").Value
      BedinDt = "" & .Fields("BedinDt").Value
      
      DoneFg = "" & .Fields("DoneFg").Value
      
      Call GetAge(DOB, Age, AgeDiv)
      Call GetSex(Sex, SexNm)
      
      If Len(DOB) = 4 Then DOB = DOB & "0000"
      If Len(DOB) = 6 Then DOB = DOB & "00"
   
   End With
   
NoData:
   MyRs.RsClose
   Set MyRs = Nothing
   Exit Function
   
Error_Trap:
   If Err.Number <> 94 Then
      MsgBox Err.Number & "  " & Err.Description
      Exit Function
   Else
      Resume Next
   End If

End Function



Private Function HIS001READ_Back(ByVal tmpSql As String) As Boolean
   
   Dim MyRs As Object
   
   On Error GoTo Error_Trap
   
   Set MyRs = OpenRecordSet(tmpSql)  'Sql 실행
   
   If MyRs.EOF Then
      HIS001READ_Back = False
      GoTo NoData
   Else
      HIS001READ_Back = True
   End If
   
   With MyRs
      
      '.MoveFirst
         
      PtId = "" & .Fields("PtId").Value
      PtNm = "" & .Fields("PtNm").Value
      SSN = "" & .Fields("SSN").Value
      DOB = "" & .Fields("DOB").Value
      Sex = "" & .Fields("Sex").Value
      EntDt = "" & .Fields("EntDt").Value
      ZipCd = "" & .Fields("ZipCd").Value
      Addr1 = "" & .Fields("Addr1").Value
      Addr2 = "" & .Fields("Addr2").Value
      OrgAddr = "" & .Fields("OrgAddr").Value
      TelHome = "" & .Fields("TelHome").Value
      TelOffice = "" & .Fields("TelOffice").Value
      PtDiv = "" & .Fields("PtDiv").Value
      FirstDt = "" & .Fields("FirstDt").Value
      EmpDiv = "" & .Fields("EmpDiv").Value
      QlfyCd = "" & .Fields("QlfyCd").Value
      MomId = "" & .Fields("MomId").Value
      RaceCd = "" & .Fields("RaceCd").Value
      NationCd = "" & .Fields("NationCd").Value
      CivilFg = "" & .Fields("CivilFg").Value
      GuardNm = "" & .Fields("GuardNm").Value
      FamDiv = "" & .Fields("FamDiv").Value
      GuardAddr = "" & .Fields("GuardAddr").Value
      
      Call GetAge(DOB, Age, AgeDiv)
      Call GetSex(Sex, SexNm)
   
   End With
   
NoData:
   MyRs.RsClose
   Set MyRs = Nothing
   Exit Function
   
Error_Trap:
   If Err.Number <> 94 Then
      MsgBox Err.Number & "  " & Err.Description
      Exit Function
   Else
      Resume Next
   End If

End Function


'% HIS002 - 입원환자 기본정보

Private Function HIS002READ(ByVal tmpSql As String) As Boolean

   Dim MyRs As Object
   
   On Error GoTo Error_Trap
   
   Set MyRs = OpenRecordSet(tmpSql)  'Sql 실행
   
   If MyRs.EOF Then
      HIS002READ = False
      GoTo NoData
   Else
      HIS002READ = True
   End If
   
   With MyRs
      
      '.MoveFirst
         
      BedinDt = "" & .Fields("BedInDt").Value
      BedInTm = "" & .Fields("BedInTm").Value
      BedOutDt = "" & .Fields("BedOutDt").Value
      BedOutTm = "" & .Fields("BedOutTm").Value
      DeptCd = "" & .Fields("DeptCd").Value
      MajDoct = "" & .Fields("MajDoct").Value
      WardId = "" & .Fields("WardId").Value
      RoomId = "" & .Fields("RoomId").Value
      BedId = "" & .Fields("BedId").Value
      BedPath = "" & .Fields("BedPath").Value
      InDiseaCd = "" & .Fields("InDiseaCd").Value
      
      DeptNm = "" & .Fields("DeptNm").Value
      DoctNm = "" & .Fields("DoctNm").Value
      
   End With
   
NoData:
   MyRs.RsClose
   Set MyRs = Nothing
   Exit Function

Error_Trap:
   If Err.Number <> 94 Then
      MsgBox Err.Number & "  " & Err.Description
      Exit Function
   Else
      Resume Next
   End If

End Function


'% 생년월일과 현재일을 기준으로 연령(월령,일령)을 구한다.

Public Sub GetAge(ByVal strDOB As String, ByRef intAge As Integer, ByRef strAgeDiv As String)
   
   Dim tmpAge As Integer
   
   On Error GoTo Err_Trap
   
   If Len(strDOB) = 4 Then strDOB = DOB & "0101"
   If Len(strDOB) = 6 Then strDOB = DOB & "01"
   
   If Not IsDate(Format(strDOB, CS_DateMask)) Then strDOB = Mid(strDOB, 1, 4) & "0101"
   
   strAgeDiv = "Y"
   intAge = DateDiff("YYYY", Format(strDOB, CS_DateMask), Now)
   If intAge = 0 Then
      strAgeDiv = "M"
      intAge = DateDiff("M", Format(strDOB, CS_DateMask), Now)
      If intAge < 6 Then
         strAgeDiv = "D"
         intAge = DateDiff("D", Format(strDOB, CS_DateMask), Now)
      End If
   End If
   Exit Sub

Err_Trap:
    intAge = 1
    strAgeDiv = "Y"
   
End Sub

      

Public Sub GetSex(ByVal strSex As String, ByRef strSexNm As String)
      
      Select Case strSex
         Case "M": strSexNm = "남자"
         Case "F": strSexNm = "여자"
         Case "U": strSexNm = "중성"
         Case Else: strSexNm = ""
      End Select

End Sub

Public Function GetBedinDt(ByVal ReceptNo As String) As String
      
   Dim MyRs As DrRecordSet
   
   Set MyRs = OpenRecordSet("select * from " & T_HIS002 & " where hosp_gb = '" & HosptGb & "' and recept_no = " & ReceptNo)
   If MyRs.EOF Then
        GetBedinDt = ""
   Else
        GetBedinDt = MyRs.Fields("Adm_Ymd").Value
   End If
   MyRs.RsClose
   Set MyRs = Nothing
   
End Function

'% 환자Id 또는 성명으로 환자마스터를 검색한다.

Public Sub PatientSearch(ByRef lstPtList As Object, ByVal SearchKey As String, ByVal SortOption As Integer)
   
   Dim I As Integer
   Dim SqlStmt As String
   Dim tmpRs As Object
   Dim tmpPtId As String
   Dim tmpPtNm As String
   Dim tmpSSN As String
   Dim MySql As New clsLISSqlStatement
   
   On Error GoTo Err_Trap
   
   lstPtList.Clear
   If SortOption = 1 Then
      SqlStmt = MySql.SqlPtntSearch(SearchKey, "ptnt_no")  '환자ID 순으로 정렬
   Else
      SqlStmt = MySql.SqlPtntSearch(SearchKey, "ptnt_nm") '환자명순으로 정렬
   End If
   Set tmpRs = OpenRecordSet(SqlStmt)
   
   medLockWindowUpdate (lstPtList.hWnd)
   While (Not tmpRs.EOF)
      tmpPtId = tmpRs.Fields("PtId").Value
      tmpPtNm = tmpRs.Fields("PtNm").Value
      tmpSSN = tmpRs.Fields("SSN").Value
      lstPtList.AddItem tmpPtId & Space(11 - Len(tmpPtId)) & tmpPtNm & " (" & tmpSSN & ")"
      tmpRs.MoveNext
   Wend
   medLockWindowUpdate (0&)
   Call medHorScrol(lstPtList)
   
   tmpRs.RsClose
   Set tmpRs = Nothing
   Set MySql = Nothing
      
   Exit Sub
   
Err_Trap:
   MsgBox DBConn.Errors.Item(0).Number & "   " & DBConn.Errors.Item(0).Description
   'Resume Next

End Sub


Public Sub ClearData()

      PtId = ""
      PtNm = ""
      SSN = ""
      DOB = ""
      Sex = ""
      EntDt = ""
      ZipCd = ""
      Addr1 = ""
      Addr2 = ""
      OrgAddr = ""
      TelHome = ""
      TelOffice = ""
      PtDiv = ""
      FirstDt = ""
      EmpDiv = ""
      QlfyCd = ""
      MomId = ""
      RaceCd = ""
      NationCd = ""
      CivilFg = ""
      GuardNm = ""
      FamDiv = ""
      GuardAddr = ""
      Age = 0
      AgeDiv = ""
      SexNm = ""
      BedinDt = ""
      BedInTm = ""
      BedOutDt = ""
      BedOutTm = ""
      DeptCd = ""
      MajDoct = ""
      WardId = ""
      RoomId = ""
      BedId = ""
      BedPath = ""
      InDiseaCd = ""
      DeptNm = ""
      DoctNm = ""

End Sub

Public Function GetDeptNm(ByVal DeptCd As String)

   Dim tmpRs As Object
   Dim MySql As New clsLISSqlStatement
   
   GetDeptNm = ""
   If DeptCd = "" Then Exit Function
   
   SqlStmt = MySql.SqlHIS003CodeList(DeptCd)
   Set tmpRs = OpenRecordSet(SqlStmt)
   If tmpRs.EOF Then
      GetDeptNm = ""
   Else
      GetDeptNm = Trim("" & tmpRs.Fields("DeptNm").Value)
   End If
   tmpRs.RsClose
   Set tmpRs = Nothing

End Function


