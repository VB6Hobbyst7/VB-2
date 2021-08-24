VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPtInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
'
''Public dbconn As New DrDatabase
'
'Private mvarPtID As String
'Private mvarPtNm As String
'Private mvarDob As String
'Private mvarBedIndt As String
'Private mvarDeptCd As String
'Private mvarDeptNm As String
'Private mvarMajdoct As String
'Private mvarDoctNm As String
'Private mvarWardId As String
'Private mvarSex As String
'Private mvarAge As String
'Private mvarPtDiv As String
'Private mvarHosilID As String
'Private mvarBedID As String
'
'
'Public Property Get BedID() As String
'    BedID = mvarBedID
'End Property
'
'Public Property Get HosilId() As String
'    HosilId = mvarHosilID
'End Property
'Public Property Get PtDiv() As String
'    PtDiv = mvarPtDiv
'End Property
'
'Public Property Get PtNm() As String
'    PtNm = mvarPtNm
'End Property
'
'Public Property Get Ptid() As String
'    Ptid = mvarPtID
'End Property
'
'Public Property Get Dob() As String
'    Dob = mvarDob
'End Property
'
'Public Property Get BedInDt() As String
'    BedInDt = mvarBedIndt
'End Property
'
'Public Property Get DeptCd() As String
'    DeptCd = mvarDeptCd
'End Property
'
'Public Property Get DeptNm() As String
'    DeptNm = mvarDeptNm
'End Property
'
'Public Property Get MajDoct() As String
'    MajDoct = mvarMajdoct
'End Property
'
'Public Property Get DoctNm() As String
'    DoctNm = mvarDoctNm
'End Property
'
'Public Property Get WardId() As String
'    WardId = mvarWardId
'End Property
'
'Public Property Get Sex() As String
'    Sex = mvarSex
'End Property
'
'Public Property Get Age() As String
'    Age = mvarAge
'End Property
'
''Public Sub setDbConn(ByRef pDbConn As DrDatabase)
''    Set dbconn = pDbConn
''End Sub
'
'Public Function GetPtInfo(ByVal Searchkey As String, ByVal blnSearch As Boolean, Optional ByVal OrdDt As String, Optional BussDiv As String = "", Optional DoneFg As String = "") As String
'    '모든환자의 정보(donefg=""),채혈대상자의 정보(donefg="0")
'    Dim strSQL As String
'
'
'    If DoneFg = "" Then                 '처방등록에서 모든환자 조회
'        GetPtInfo = " SELECT " & F_PTID & " as ptid," & F_PTNM & " as ptnm," & F_SSN & " as ssn, '' as dob" & _
'                    " FROM " & T_HIS001 & " a" & _
'                    " WHERE "
'    Else
'
'    End If
'   '환자ID는 존재여부만 확인한다.
'    If blnSearch = True Then
'        GetPtInfo = GetPtInfo & " a." & F_PTID & " >= " & DBV(F_PTID, Searchkey)
'    Else
'        GetPtInfo = GetPtInfo & " a." & F_PTNM & " like " & "'" & Searchkey & "%" & "'"
'    End If
'
'End Function
'Public Function Get_Ptid(ByVal Ptid As String) As String
'    Get_Ptid = " SELECT " & F_PTID & " FROM " & T_HIS001 & _
'               " WHERE " & dbw(F_PTID & " =", Ptid)
'End Function
'Public Function BedPt_Chk(ByVal strPtid As String, ByVal strDate As String) As Boolean
'    '입원환자 조회를 한다.(전에 입원했던 환자도 모조리)
'    Dim strSQL As String
'    Dim DrRs As Recordset
'
'    Set DrRs = New Recordset
'
'    strSQL = " SELECT " & F_INPTID & " as ptid FROM " & T_HIS002 & _
'             " WHERE " & dbw(F_INPTID & " =", strPtid)
'
'    DrRs.Open strSQL, dbconn
'    If Not DrRs.EOF = True Then
'        '입원환자이다.(퇴원한 환자일수도있다.)
'
'        BedPt_Info strPtid, strDate
'    Else
'        '입원환자아니다.
'
'        GetBedPtInfo strPtid, False
'    End If
'    Set DrRs = Nothing
'
'End Function
'Private Function BedPt_Info(ByVal strPtid As String, ByVal strBedInDt As String) As Boolean
'    '입원환자를 조회(지금 현재 입원한 환자만 조회)
'    Dim Rs      As Recordset
'    Dim strSQL  As String
'
'    strSQL = " SELECT " & F_INPTID & " as ptid FROM " & T_HIS002 & _
'             " WHERE " & _
'                               dbw(F_INPTID & "=", strPtid) & _
'                     " AND in_date <= to_date('" & strBedInDt & "','yyyymmdd') "
'
'    Set Rs = New Recordset
'    Rs.Open strSQL, dbconn
'
''    If RS.DBerror = True Then
''        dbconn.DisplayErrors
''        Set RS = Nothing
''        Exit Function
''    End If
'
'    If Not Rs.EOF = True Then
'        '입원환자(지금 현재도 입원했는지는 미지수)
'        BedOut_chk strPtid, strBedInDt
'    End If
'    Set Rs = Nothing
'
'End Function
'
'Private Function BedOut_chk(ByVal strPtid As String, ByVal BadOut As String) As Boolean
'    Dim DrRs As Recordset
'    Dim strSQL As String
'
'    Set DrRs = New Recordset
'    strSQL = " SELECT " & F_BEDOUTDT & " as bedoutdt FROM " & T_HIS002 & _
'             " WHERE " & _
'                               dbw(F_INPTID & "=", strPtid) & _
'                     " AND in_date <= to_date('" & BadOut & "','yyyymmdd')"
'
'    DrRs.Open strSQL, dbconn
'
''    If DrRs.DBerror = True Then
''        dbconn.DisplayErrors
''        DrRs.RsClose
''        Set DrRs = Nothing
''        Exit Function
''    End If
'    If Not DrRs.EOF = True Then
'
'       '과거에 입원했던 환자.
'        GetBedPtInfo strPtid, False
'    Else
'        '확실한 입원환자
'
'        GetBedPtInfo strPtid, True
'    End If
'    Set DrRs = Nothing
'End Function
'
'Public Function GetBedPtInfo(ByVal strPtid As String, blnBedIn As Boolean) As Boolean
'    '현재환자의 정보를 가지고 온다.
'    'True=입원환자 : False=외래환자
'    Dim strSQL As String
'    Dim DrRs1 As New Recordset
'
'    mvarPtID = strPtid
'
'    If blnBedIn = True Then
'        strSQL = "SELECT " & F_SSN2("a") & " as ssn," & F_DOB2("a") & " as dob, " & _
'                " a." & F_PTNM & " as ptnm,b." & F_BEDINDT & " as bedindt," & _
'                " b." & F_PTDEPTCD & " as deptcd, c." & F_DEPTNM & " as deptnm," & _
'                " b." & F_MAJDOCT & " as majdoct , e." & F_DOCTNM & " as doctnm, " & _
'                " b." & F_PTWARDID & " as wardid ,b." & F_PTROOMID & " as roomid ," & F_PTBEDID & " as bedid  " & _
'                " FROM " & T_HIS001 & " a," & T_HIS002 & " b," & T_HIS003 & " c," & T_HIS005 & " e " & _
'                " WHERE " & _
'                                    dbw(" b." & F_INPTID & " =", strPtid) & " " & _
'                        " AND b." & F_INPTID & " =a." & F_PTID & _
'                        " AND b." & F_PTDEPTCD & " *= c." & F_DEPTCD & _
'                        " AND b." & F_MAJDOCT & " *= e." & F_DOCTID
'    Else
'        strSQL = " SELECT " & F_PTID & " as ptid," & F_SSN2 & " as ssn," & F_DOB2 & " as dob," & _
'                              F_PTNM & " as ptnm," & F_TMPDIV & " as tmpdiv" & _
'                 " FROM " & T_HIS001 & _
'                 " WHERE " & dbw(F_PTID & " =", strPtid) & ""
'    End If
'
'    DrRs1.Open strSQL, dbconn
''    If DrRs1.DBerror = True Then
''        dbconn.DisplayErrors
''        DrRs1.RsClose
''        Set DrRs1 = Nothing
''        GetBedPtInfo = False
''        Exit Function
''    End If
'
'    If blnBedIn = True Then
'        If Not DrRs1.EOF = True Then
'            mvarPtDiv = "BED" '병동환자
'            mvarDob = Mid(DrRs1.Fields("dob").Value, 1, 4) & "/" & Mid(DrRs1.Fields("dob").Value, 5, 2) & "/" & Mid(DrRs1.Fields("dob").Value, 7, 2)
'            mvarPtNm = DrRs1.Fields("ptnm")
'            mvarBedIndt = Mid(DrRs1.Fields("BedIndt").Value, 1, 4) & "-" & Mid(DrRs1.Fields("BedIndt").Value, 5, 2) & "-" & Mid(DrRs1.Fields("BedIndt").Value, 7)
'            mvarDeptCd = DrRs1.Fields("Deptcd").Value
'            mvarDeptNm = DrRs1.Fields("DeptNm").Value
'            mvarMajdoct = DrRs1.Fields("Majdoct").Value
'            mvarDoctNm = DrRs1.Fields("DoctNm").Value
'            mvarWardId = Trim(DrRs1.Fields("WardId").Value)
'            mvarHosilID = Trim(DrRs1.Fields("roomid").Value)
'            mvarBedID = Trim(DrRs1.Fields("bedid").Value)
'            mvarAge = medFindAge(Mid(DrRs1.Fields("dob").Value, 1, 8), "Y")
''            mvarSex = IIf(Mid(DrRs1.Fields("ssn").Value, 9, 1) = "1" Or "3", "남", "여")   'Modify By Legends 2001/03/02
'            Select Case Mid(DrRs1.Fields("ssn").Value, 7, 1)
'                Case "1", "3"
'                    mvarSex = "남"
'                Case "2", "4"
'                    mvarSex = "여"
'            End Select
'        Else
'            mvarPtDiv = "NOBED" '외래환자
'            mvarPtNm = DrRs1.Fields("ptnm").Value
'            mvarDob = Mid(DrRs1.Fields("dob").Value, 1, 4) & "/" & Mid(DrRs1.Fields("dob").Value, 5, 2) & "/" & Mid(DrRs1.Fields("dob").Value, 7)
'            mvarAge = medFindAge(Mid(DrRs1.Fields("dob").Value, 1, 8), "Y")
'                Select Case Mid(DrRs1.Fields("ssn").Value, 7, 1)
'                    Case "1", "3"
'                        mvarSex = "남"
'                    Case "2", "4"
'                        mvarSex = "여"
'                End Select
'
'            mvarBedIndt = ""
'        End If
'    Else
'        mvarPtDiv = "NOBED" '외래환자
'        mvarPtNm = DrRs1.Fields("ptnm").Value
'        mvarDob = Mid("" & DrRs1.Fields("dob").Value, 1, 4) & "/" & Mid(DrRs1.Fields("dob").Value, 5, 2) & "/" & Mid(DrRs1.Fields("dob").Value, 7, 2)
'        If IsDate(mvarDob) Then
'            mvarAge = medFindAge(Mid("" & DrRs1.Fields("dob").Value, 1, 8), "Y")
'        Else
'            mvarAge = 0
'        End If
'            Select Case Mid(DrRs1.Fields("ssn").Value, 7, 1)
'                Case "1", "3"
'                    mvarSex = "남"
'                Case "2", "4"
'                    mvarSex = "여"
'            End Select
'
'        mvarBedIndt = ""
'    End If
'    Set DrRs1 = Nothing: GetBedPtInfo = True
'
'End Function
'
