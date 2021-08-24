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

'환자와 관련된 정보 수집(병원별로 이 클래스를 고치면 된다.)
'Coding By Legends 2004/08/12

'환자 검색 이벤트용
Private WithEvents objSearchForm As frmSearchPt
Attribute objSearchForm.VB_VarHelpID = -1

Public Event SelectedId(ByVal vPtID As String)

'환자 마스터
Private mvarPtID As String   '환자아뒤
Private mvarPtNm As String   '환자명
Private mvarSSN As String    '주민번호
Private mvarAge As String    '나이
Private mvarAGEDIV As String '나이구분
Private mvarSex As String    '성별
Private mvarSEXNM As String '성별명
Private mvarSEXAGE As String '성별나이
Private mvarDob As String    '생년월일
Private mvarZIPCD As String  '우편번호
Private mvarADDR As String   '주소
Private mvarTEL As String    '전화번호
Private mvarPtDiv As String  '환자유형
Private mvarTMPDIV As String '검진 임시ID 구분

'재원 마스터
Private mvarINPTID As String '환자아뒤
Private mvarBedIndt As String    '입원일자
Private mvarBEDINTM As String    '입원시간
Private mvarBEDOUTDT As String   '퇴원일자
Private mvarBEDOUTTM As String   '퇴원시간
Private mvarDeptCd As String   '진료과
Private mvarBEDPATH As String    '입원경로
Private mvarINDISEACD As String  '입원상병
Private mvarWardId As String   '병동아뒤
Private mvarROOMID As String   '병실아뒤
Private mvarBedID As String    '베드아뒤
Private mvarMajdoct As String    '주치의아뒤

''부서 마스터
'private mvarDEPTCD As String '부서코드
Private mvarDeptNm As String '부서명
'private mvarDEPTDIV As String    '부서구분
'
''병상 마스터
'private mvarWARDID As String '병동아뒤
'private mvarWARDNM As String '병동명
'private mvarROOMID As String '병실아뒤
'private mvarBEDID As String  '베드아뒤
'
''의사 마스터 및 직원마스터
'private mvarDOCTID As String '의사아뒤
Private mvarDoctNm As String '의사명
'private mvarEMPID As String  '직원아뒤
'private mvarEMPNM As String  '직원명
'private mvarEMPDIV  As String    '직종 구분
'private mvarEXPDT As String  '퇴사일
'
''상병 마스터
'private mvarICD As String    '상병코드
'private mvarIENM As String   '영문상병명
'private mvarIKNM As String   '국문상병명
'
''수가 마스터
'private mvarAMTCD As String  '수가코드
'private mvarAMTNM As String  '코드명
'
''수술코드 마스터
'private mvarOCD As String    '수술코드
'private mvarOENM As String   '영문수술명
'private mvarOKNM As String   '국문수술명

'기타 정보
Private mvarINADMISSION As Boolean '입원 여부
Private mvarEXISTADM As Boolean '재원마스터 존재여부

'--------------------------------------------------------------------------------
'환자 마스터
Public Property Let Ptid(ByVal vData As String)
    mvarPtID = vData
End Property

Public Property Get Ptid() As String
    Ptid = mvarPtID
End Property

Public Property Let PtNm(ByVal vData As String)
    mvarPtNm = vData
End Property

Public Property Get PtNm() As String
    PtNm = mvarPtNm
End Property

Public Property Let SSN(ByVal vData As String)
    mvarSSN = vData
End Property

Public Property Get SSN() As String
    SSN = mvarSSN
End Property

Public Property Get Age() As String
    Age = mvarAge
End Property

Public Property Get AGEDIV() As String
    AGEDIV = mvarAGEDIV
End Property

Public Property Get Sex() As String
    Sex = mvarSex
End Property

Public Property Get SEXNM() As String
    SEXNM = mvarSEXNM
End Property

Public Property Get SEXAGE() As String
    SEXAGE = mvarSEXAGE
End Property

Public Property Get Dob() As String
    Dob = mvarDob
End Property

Public Property Get ZIPCD() As String
    ZIPCD = mvarZIPCD
End Property

Public Property Get ADDR() As String
    ADDR = mvarADDR
End Property

Public Property Get TEL() As String
    TEL = mvarTEL
End Property

Public Property Get PtDiv() As String
    PtDiv = mvarPtDiv
End Property

Public Property Get TmpDiv() As String
    TmpDiv = mvarTMPDIV
End Property
'--------------------------------------------------------------------------------
'재원 마스터
Public Property Get INPTID() As String
    INPTID = mvarINPTID
End Property

Public Property Get BedInDt() As String
    BedInDt = mvarBedIndt
End Property

Public Property Get BEDINTM() As String
    BEDINTM = mvarBEDINTM
End Property

Public Property Get BEDOUTDT() As String
    BEDOUTDT = mvarBEDOUTDT
End Property

Public Property Get BEDOUTTM() As String
    BEDOUTTM = mvarBEDOUTTM
End Property

Public Property Get DeptCd() As String
    DeptCd = mvarDeptCd
End Property

Public Property Let DeptCd(ByVal vData As String)
    mvarDeptCd = vData
End Property

Public Property Get BEDPATH() As String
    BEDPATH = mvarBEDPATH
End Property

Public Property Get InDiseaCd() As String
    InDiseaCd = mvarINDISEACD
End Property

Public Property Get WardId() As String
    WardId = mvarWardId
End Property

Public Property Get ROOMID() As String
    ROOMID = mvarROOMID
End Property

Public Property Get BedID() As String
    BedID = mvarBedID
End Property

Public Property Get MajDoct() As String
    MajDoct = mvarMajdoct
End Property

Public Property Let MajDoct(ByVal vData As String)
    mvarMajdoct = mvarMajdoct
End Property

'--------------------------------------------------------------------------------
'부서 마스터
'Public Property Get DEPTCD() As String
'    DEPTCD = mvarDEPTCD
'End Property

Public Property Get DeptNm() As String
    DeptNm = mvarDeptNm
End Property

'Public Property Get DEPTDIV() As String
'    DEPTDIV = mvarDEPTDIV
'End Property
'--------------------------------------------------------------------------------
'병상 마스터
'Public Property Get WARDID() As String
'    WARDID = mvarWARDID
'End Property
'
'Public Property Get WARDNM() As String
'    WARDNM = mvarWARDNM
'End Property
'
'Public Property Get ROOMID() As String
'    ROOMID = mvarROOMID
'End Property
'
'Public Property Get BEDID() As String
'    BEDID = mvarBEDID
'End Property
'--------------------------------------------------------------------------------
'의사 마스터 및 직원 마스터
'Public Property Get DOCTID() As String
'    DOCTID = mvarDOCTID
'End Property
'
Public Property Get DoctNm() As String
    DoctNm = mvarDoctNm
End Property

Public Property Let DoctNm(ByVal vData As String)
    mvarDoctNm = vData
End Property

'
'Public Property Get EMPID() As String
'    EMPID = mvarEMPID
'End Property
'
'Public Property Get EMPNM() As String
'    EMPNM = mvarEMPNM
'End Property
'
'Public Property Get EMPDIV() As String
'    EMPDIV = mvarEMPDIV
'End Property
'
'Public Property Get EXPDT() As String
'    EXPDT = mvarEXPDT
'End Property
'--------------------------------------------------------------------------------
'상병 마스터
'Public Property Get ICD() As String
'    ICD = mvarICD
'End Property
'
'Public Property Get IENM() As String
'    IENM = mvarIENM
'End Property
'
'Public Property Get IKNM() As String
'    IKNM = mvarIKNM
'End Property
'--------------------------------------------------------------------------------
'수가 마스터
'Public Property Get AMTCD() As String
'    AMTCD = mvarAMTCD
'End Property
'
'Public Property Get AMTNM() As String
'    AMTNM = mvarAMTNM
'End Property
'--------------------------------------------------------------------------------
'수술코드 마스터
'Public Property Get OCD() As String
'    OCD = mvarOCD
'End Property
'
'Public Property Get OENM() As String
'    OENM = mvarOENM
'End Property
'
'Public Property Get OKNM() As String
'    OKNM = mvarOKNM
'End Property
'--------------------------------------------------------------------------------
'기타 정보
Public Property Get INADMISSION() As Boolean
    INADMISSION = mvarINADMISSION
End Property

Public Property Get EXISTADM() As Boolean
    EXISTADM = mvarEXISTADM
End Property

Public Function FUNC_CONVERT(ByVal vType As String, ByVal vField As String, _
                             Optional ByVal vFormat As String = "") As String
'to_char num, date변환
'to_num char 변환
'to_date char 변환,포맷 필수
    
    Select Case UCase(vType)
        Case "CHAR"
            If vFormat = "" Then
                FUNC_CONVERT = " to_char(" & vField & ") "
            Else
                FUNC_CONVERT = " to_char(" & vField & ", '" & vFormat & "')"
            End If
        Case "NUM"
            If vFormat = "" Then
                FUNC_CONVERT = " to_number(" & vField & ") "
            Else
                FUNC_CONVERT = " to_number(" & vField & ", '" & vFormat & "')"
            End If
        Case "DATE"
            FUNC_CONVERT = " to_date(" & vField & ", '" & vFormat & "')"
    End Select
End Function

' 2017-06-28 신규암호화
' crypto.dec('cpattern1',resno2)
Private Function F_SSN2(Optional ByVal tA As String) As String
    If IsMissing(tA) Then
        F_SSN2 = "crypto.dec('cpattern1',resno1) || crypto.dec('cpattern1',resno1)"
    Else
        If tA = "" Then
            F_SSN2 = "crypto.dec('cpattern1',resno1) || crypto.dec('cpattern1',resno1)"
        Else
            F_SSN2 = "crypto.dec('cpattern1'," & tA & ".resno1) || " & "crypto.dec('cpattern1'," & tA & ".resno2)"
        End If
    End If
End Function

' 구암호화
'Private Function F_SSN2(Optional ByVal tA As String) As String
'    If IsMissing(tA) Then
'        F_SSN2 = "CryptIT.decrypt(resno1,'pmc1898') || CryptIT.decrypt(resno2,'pmc1898')"
'    Else
'        If tA = "" Then
'            F_SSN2 = "CryptIT.decrypt(resno1,'pmc1898') || CryptIT.decrypt(resno2,'pmc1898')"
'        Else
'            F_SSN2 = tA & ".resno1 || " & tA & ".resno2"
'        End If
'    End If
'End Function

' 2017.05.23 암호화 해제
'Private Function F_SSN2(Optional ByVal tA As String) As String
'    If IsMissing(tA) Then
'        F_SSN2 = "resno1 || resno2"
'    Else
'        If tA = "" Then
'            F_SSN2 = "resno1 || resno2"
'        Else
'            F_SSN2 = tA & ".resno1 || " & tA & ".resno2"
'        End If
'    End If
'End Function

Private Function F_DOB2(Optional ByVal tA As String) As String
    If IsMissing(tA) Then
        F_DOB2 = FUNC_CONVERT("char", "birtdate", "yyyymmdd")
    Else
        If tA = "" Then
            F_DOB2 = FUNC_CONVERT("char", "birtdate", "yyyymmdd")
        Else
            F_DOB2 = FUNC_CONVERT("char", tA & ".birtdate", "yyyymmdd")
        End If
    End If
End Function

' 2017-06-28 신규암호화
' crypto.dec('cpattern1',resno2)
Private Function F_SEX2(Optional ByVal tA As String) As String

    If IsMissing(tA) Then
        F_SEX2 = "substr(crypto.dec('cpattern1',resno2), 1, 1)"
    Else
        If tA = "" Then
            F_SEX2 = "substr(crypto.dec('cpattern1',resno2), 1, 1)"
        Else
            F_SEX2 = "substr(crypto.dec('cpattern1'," & tA & ".resno2),1,1)" 'tA & ".resno2, 1, 1)"
        End If
    End If
End Function

' 구암호화
'Private Function F_SEX2(Optional ByVal tA As String) As String
'    If IsMissing(tA) Then
'        F_SEX2 = "substr(CryptIT.decrypt(resno2,'pmc1898'), 1, 1)"
'    Else
'        If tA = "" Then
'            F_SEX2 = "substr(CryptIT.decrypt(resno2,'pmc1898'), 1, 1)"
'        Else
'            F_SEX2 = "substr(" & tA & ".resno2, 1, 1)"
'        End If
'    End If
'End Function

' 2017.05.23 암호화 해제
'Private Function F_SEX2(Optional ByVal tA As String) As String
'    If IsMissing(tA) Then
'        F_SEX2 = "substr(resno2, 1, 1)"
'    Else
'        If tA = "" Then
'            F_SEX2 = "substr(resno2, 1, 1)"
'        Else
'            F_SEX2 = "substr(" & tA & ".resno2, 1, 1)"
'        End If
'    End If
'End Function

Private Function F_BEDOUTDT2(Optional ByVal tA As String) As String
    If IsMissing(tA) Then
        F_BEDOUTDT2 = FUNC_CONVERT("char", "dschdate", "yyyymmdd")
    Else
        If tA = "" Then
            F_BEDOUTDT2 = FUNC_CONVERT("char", "dschdate", "yyyymmdd")
        Else
            F_BEDOUTDT2 = FUNC_CONVERT("char", tA & ".dschdate", "yyyymmdd")
        End If
    End If
End Function

Private Function F_BEDINDT2(Optional ByVal tA As String) As String
    If IsMissing(tA) Then
        F_BEDINDT2 = FUNC_CONVERT("char", "admdate", "yyyymmdd")
    Else
        If tA = "" Then
            F_BEDINDT2 = FUNC_CONVERT("char", "admdate", "yyyymmdd")
        Else
            F_BEDINDT2 = FUNC_CONVERT("char", tA & ".admdate", "yyyymmdd")
        End If
    End If
End Function

Public Function F_BEDOUTTM2(Optional ByVal tA As String) As String
    If IsMissing(tA) Then
        F_BEDOUTTM2 = FUNC_CONVERT("char", "dschtime", "hh24:mi:ss")
    Else
        If tA = "" Then
            F_BEDOUTTM2 = FUNC_CONVERT("char", "dschtime", "hh24:mi:ss")
        Else
            F_BEDOUTTM2 = FUNC_CONVERT("char", tA & ".dschtime", "hh24:mi:ss")
        End If
    End If
End Function

Public Function F_BEDINTM2(Optional ByVal tA As String) As String
    If IsMissing(tA) Then
        F_BEDINTM2 = FUNC_CONVERT("char", "admtime", "hh24:mi:ss")
    Else
        If tA = "" Then
            F_BEDINTM2 = FUNC_CONVERT("char", "admtime", "hh24:mi:ss")
        Else
            F_BEDINTM2 = FUNC_CONVERT("char", tA & ".admtime", "hh24:mi:ss")
        End If
    End If
End Function


Public Function GETPatient(ByVal vPtID As String) As Boolean
'환자조회용(환자가 있는 경우 True, 없는 경우 또는 오류 난 경우 False
'첫번째 외래환자 또는 병동환자 판단
'병동환자의 경우 병동아뒤가 조건에 포함되는지 판단
    
    GETPatient = False
    
    '환자 마스터에서 검색
    If CHKForeign(vPtID) Then
        GETPatient = True
        
        '재원 마스터에서 재검색
        mvarEXISTADM = CHKAdmission(vPtID)
    End If
End Function

Private Function CHKForeign(ByVal vPtID As String) As Boolean
'외래환자 마스터에 있는지 체크
    Dim Rs As Recordset
'    Dim strSQL As String
'
'    strSQL = " SELECT " & mF_PTID & " as ptid, " & mF_PTNM & " as ptnm, " & _
'                           F_SSN2 & " as ssn,  " & mF_DOB & " as dbo, " & _
'                           F_SEX2 & " as sex,  " & mF_ZIPCODE & " as zipcd, " & _
'                           mF_ADDRESS & " as addr, " & mF_TEL & " as tel, " & mF_TMPDIV & " as tmpdiv " & _
'             " FROM   " & mT_HIS001 & " " & _
'             " WHERE  " & dbw(mF_PTID, vPtId, 2)

    Set Rs = New Recordset
    
    On Error GoTo NoData
    Rs.Open GetSQLPt(vPtID), DBConn
    If Rs.EOF Then GoTo NoData
    
    With Rs
        mvarPtID = .Fields("ptid").Value & ""
        mvarPtNm = .Fields("ptnm").Value & ""
        mvarSSN = .Fields("ssn").Value & ""
'        mvarAGE = .Fields("age").Value & ""
        mvarSex = .Fields("sex").Value & ""
        If IsNumeric(mvarSex) Then
            mvarSex = Choose((Val(mvarSex) Mod 2) + 1, "F", "M")
        End If
        mvarDob = .Fields("dob").Value & ""
        If Not IsDate(Format(mvarDob, "####-##-##")) Then
            If IsDate(Format(mvarSSN, "####-##-##")) Then
                mvarDob = mvarSSN
            Else
                If Len(Mid(mvarDob, 1, 4)) = 4 Then
                    mvarDob = Mid(mvarDob, 1, 4) & "0101"
                ElseIf Len(Mid(mvarSSN, 1, 4)) = 4 Then
                    mvarDob = Mid(mvarSSN, 1, 4) & "0101"
                Else
                    mvarDob = Format(GetSystemDate, CS_DateDbFormat)
                End If
            End If
        End If
        If Len(mvarDob) = 4 Then mvarDob = mvarDob & "0101"
        If Len(mvarDob) = 6 Then mvarDob = mvarDob & "01"
        
        mvarZIPCD = .Fields("zipcd").Value & ""
        mvarADDR = .Fields("addr").Value & ""
        mvarTEL = .Fields("tel").Value & ""
        mvarTMPDIV = .Fields("tmpdiv").Value & ""
        
        Call GetAge '(DOB, AGE, AGEDIV)
        Call GetSexNm '(SEX, SEXNM)
        mvarSEXAGE = mvarSEXNM & " / " & mvarAge & mvarAGEDIV
        
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
        
    End With
    
    CHKForeign = True
    Set Rs = Nothing
    Exit Function
    
NoData:
    CHKForeign = False
    Set Rs = Nothing
End Function

Private Sub GetAge() '(ByVal strDOB As String, ByRef intAge As Integer, ByRef strAgeDiv As String)
'% 생년월일과 현재일을 기준으로 연령(월령,일령)을 구한다.

    Dim tmpAge As Integer

On Error GoTo Err_Trap

    If Len(mvarDob) = 6 Then mvarDob = mvarDob & "01"
    If Not IsDate(Format(mvarDob, CS_DateMask)) Then mvarDob = Mid(mvarDob, 1, 4) & "0101"

    mvarAGEDIV = "Y"
'    mvarAge = DateDiff("YYYY", Format(mvarDob, "####-##-##"), GetSystemDate)
'    AgeYY = mvarAGE  '연령

    If Val(Format(GetSystemDate, "mmdd")) > Val(Mid(mvarDob, 5, 4)) Then
        mvarAge = DateDiff("YYYY", Format(mvarDob, "####-##-##"), GetSystemDate)
    Else
        mvarAge = DateDiff("YYYY", Format(mvarDob, "####-##-##"), GetSystemDate) - 1
    End If

    If mvarAge = 0 Or mvarAge = 1 Then
       mvarAGEDIV = "M"
       mvarAge = DateDiff("M", Format(mvarDob, "####-##-##"), GetSystemDate)   '월령
       If mvarAge < 6 Then
          mvarAGEDIV = "D"
          mvarAge = DateDiff("D", Format(mvarDob, "####-##-##"), GetSystemDate)    '일령
       End If
    End If
    Exit Sub

Err_Trap:
    mvarAge = 1
    mvarAGEDIV = "Y"
End Sub

Private Sub GetSexNm()
      Select Case mvarSex
         Case "M": mvarSEXNM = "남자"
         Case "F": mvarSEXNM = "여자"
         Case "U": mvarSEXNM = "중성"
         Case Else: mvarSEXNM = mvarSex
      End Select
End Sub

Private Function CHKAdmission(ByVal vPtID As String) As Boolean
'재원 마스터에 있는지 체크
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " SELECT a." & mF_PTID & " as ptid," & F_BEDINDT2 & " as bedindt, '' as bedintm, " & _
              "      " & F_BEDOUTDT2 & " as bedoutdt, '' as bedouttm, " & _
              "        a." & mF_PTDEPTCD & " as deptcd, a." & mF_MAJDOCT & " as majdoct, " & _
              "        a." & mF_PTWARDID & " as wardid, a." & mF_PTROOMID & " as hosilid, '' as bedid, " & _
              "        '' as bedpath, '' as indiseacd, " & _
              "        b." & mF_DEPTNM & " as deptnm, c." & mF_DOCTNM & " as doctnm " & _
              " FROM " & mT_HIS002 & " a, " & _
                        mT_HIS003 & " b, " & _
                        mT_HIS005 & " c  " & _
              " WHERE  " & DBW(" a." & mF_PTID, vPtID, 2) & _
              " AND    b." & mF_DEPTCD & " = a." & mF_PTDEPTCD & _
              " AND    c." & mF_DOCTID & " = a." & mF_MAJDOCT & _
              " AND (a.dschdate is null or a.dschdate='') " & _
              " AND (a.REJTDATE is null or a.REJTDATE = '') " & _
              " ORDER BY bedindt DESC"
              
    Set Rs = New Recordset
    
    On Error GoTo NoData
    Rs.Open strSQL, DBConn
    If Rs.EOF Then GoTo NoData
    
    With Rs
        mvarINPTID = .Fields("ptid").Value & ""
        mvarBedIndt = .Fields("bedindt").Value & ""
        mvarBEDINTM = .Fields("bedintm").Value & ""
        mvarBEDOUTDT = .Fields("bedoutdt").Value & ""
        
        If mvarBEDOUTDT = "" Then mvarINADMISSION = True
        
        mvarBEDOUTTM = .Fields("bedouttm").Value & ""
        mvarDeptCd = .Fields("deptcd").Value & ""
        mvarBEDPATH = ""
        mvarINDISEACD = ""
        mvarWardId = .Fields("wardid").Value & ""
        mvarROOMID = .Fields("hosilid").Value & ""
        mvarBedID = .Fields("bedid").Value & ""
        mvarMajdoct = .Fields("majdoct").Value & ""
        mvarDeptNm = .Fields("deptnm").Value & ""
        mvarDoctNm = .Fields("doctnm").Value & ""
    End With
    
    CHKAdmission = True
    Set Rs = Nothing
    
    Exit Function
    
NoData:
    Set Rs = Nothing
    CHKAdmission = False
End Function

Public Function GetSQLPtNt(ByVal vCase As String, ByVal vSearchKey As String, _
                           Optional ByVal vWardId As String = "", Optional ByVal vVfyDt As String = "") As String
'환자 검색용 sql
'환자ID, 환자명, 채혈대상, 결과대상, 외래, 병동으로 검색
' vCase : 1-환자ID, 2-환자명,
'           3-환자ID(외래채혈대상), 4-환자명(외래채혈대상),
'           5-환자ID(병동채혈대상), 6-환자명(병동채혈대상)
    
    Dim strSQL As String
    Dim strBussDiv As String
        
'    If vCase = 3 Or vCase = 4 Then
'        strBussDiv = "1"
'    ElseIf vCase = 5 Or vCase = 6 Then
'        strBussDiv = "2"
'    End If
' 2017.05.24 vcase=7 추가
    If vCase = 3 Or vCase = 4 Or vCase = 7 Then
        strBussDiv = "1"
    ElseIf vCase = 5 Or vCase = 6 Then
        strBussDiv = "2"
    End If

' 2015.05.23 암호화 해제
    strSQL = " select distinct a." & mF_PTID & " as ptid, " & " a." & mF_PTNM & " as ptnm, " & _
             F_SSN2("a") & " as ssn, " & F_DOB2("a") & " as dob, " & _
             " a.zpaddress || ' ' || a." & mF_ADDRESS & _
             " as address, " & "a." & mF_TEL & " as telno, " & "a." & mF_HPTEL & " as hptelno " & _
             " from " & mT_HIS001 & " a "
'
'    strSQL = " select distinct a." & mF_PTID & " as ptid, " & " a." & mF_PTNM & " as ptnm, " & _
'             F_SSN2 & " as ssn, " & F_DOB2 & " as dob, " & _
'             " a.zpaddress || ' ' || a." & mF_ADDRESS & _
'             " as address, " & "a." & mF_TEL & " as telno, " & "a." & mF_HPTEL & " as hptelno " & _
'             " from " & mT_HIS001 & " a "
             
    If vVfyDt <> "" Then '결과보고 대상 검색
        If Val(vCase) Mod 2 = 1 Then '아디로 검색
            strSQL = strSQL & " where  a." & mF_PTID & " = " & DBV(mF_PTID, vSearchKey) & " " & " and rownum <=100 "
        Else '환자명으로 검색
            strSQL = strSQL & " where  a." & mF_PTNM & " like '" & vSearchKey & "%' "
        End If
        
        If vWardId <> "" Then
            strSQL = strSQL & " and exists (select * from " & mT_LAB202 & " b " & _
                                           " where " & DBW("b.deptcd=", vWardId) & _
                                           " and " & DBW("b.vfydt=", vVfyDt) & _
                                           " and b.ptid = a." & mF_PTID & " ) "
        End If
    Else
        If Val(vCase) > 2 Then    '채혈대상
            strSQL = strSQL & ", " & mT_LAB101 & " b "
            strSQL = strSQL & _
                    " WHERE  " & DBW("b.donefg =", "0") & _
                    " and " & DBW("b.bussdiv =", strBussDiv) & _
                    " and " & DBW("b.reqdt>=", Format(DateAdd("m", -6, Now), "YYYYMMDD")) & _
                    " and a." & mF_PTID & " = b.ptid  and "
        Else
            strSQL = strSQL & " WHERE  "
        End If

'''        If Val(vCase) Mod 2 = 1 Then
'''            If vCase = "99" Then
'''                strSQL = strSQL & " a." & mF_PTID & " = " & DBV(mF_PTID, vSearchKey) & " "
'''            Else
'''                strSQL = strSQL & " a." & mF_PTID & " = " & DBV(mF_PTID, vSearchKey) & " " & " and rownum <= 100 "
'''            End If
'''        Else
'''            strSQL = strSQL & " a." & mF_PTNM & " like '" & vSearchKey & "%' "
'''        End If
' 2017.05.24 생년월일로 조회 추가
        If Val(vCase) Mod 2 = 1 Then
            If vCase = "99" Then
                strSQL = strSQL & " a." & mF_PTID & " = " & DBV(mF_PTID, vSearchKey) & " "
            ElseIf Val(vCase) = 7 Then
                strSQL = strSQL & " a.resno1 like '" & vSearchKey & "%' "
            Else
                strSQL = strSQL & " a." & mF_PTID & " = " & DBV(mF_PTID, vSearchKey) & " " & " and rownum <= 100 "
            End If
        ElseIf Val(vCase) = 0 Then
            strSQL = strSQL & " a.resno1 like '" & vSearchKey & "%' "
        Else
            strSQL = strSQL & " a." & mF_PTNM & " like '" & vSearchKey & "%' "
        End If
        
        If vWardId <> "" Then
            strSQL = strSQL & " and b.wardid = '" & vWardId & "' "
        End If
    End If
    
    If Val(vCase) Mod 2 = 1 Then  '아디로 검색
        strSQL = strSQL & " order  by a." & mF_PTID & ""
    Else '환자명으로 검색
        strSQL = strSQL & " order  by a." & mF_PTNM & ", ssn"
    End If
    
    GetSQLPtNt = strSQL
End Function

Public Function GetSQLCol_Wm(ByVal vCase As String, ByVal vSearchKey As String, _
                        Optional ByVal vWardId As String = "") As String
'병동용 채혈대상 검색
    Dim strSQL As String
    Dim SQL    As String
    
'    strSQL = " select distinct a." & mF_PTID & " as ptid, a." & mF_PTNM & " as ptnm," & _
'                                  F_SSN2("a") & " as SSN, " & _
'                                  F_DOB2("a") & " as DOB " & _
'            " from " & mT_HIS001 & " a ," & mT_LAB101 & " b," & mT_LAB102 & " c " & _
'            " where " & DBW("b.donefg =", "0") & _
'            " and   " & DBW("b.bussdiv =", "2") & _
'            " and    a." & mF_PTID & " = b.ptid"
'
'    strSQL = strSQL & " AND b.orddt between '" & vCase & "' and '" & vSearchKey & "' "
'
'    strSQL = strSQL & " AND b.wardid = '" & vWardId & "'" & _
'                " AND b.ptid=c.ptid" & _
'                " AND b.orddt=c.orddt" & _
'                " AND b.ordno=c.ordno" & _
'                " AND (c.dcfg='' or c.dcfg is null)" & _
'                " AND c.ordcd in ('C3710-1','C3710','C37101','C37102','C3710B','C3710C','C3710D','C3710E') "
'
'    strSQL = strSQL & " ORDER BY a." & mF_PTNM & ", SSN"


    SQL = " select DISTINCT b.patno AS ptid,                                                                             "
    SQL = SQL + "       b.patname AS ptnm,                                                                                     "
    SQL = SQL + "       crypto.dec('cpattern1',b.resno1) || crypto.dec('cpattern1',b.resno2) AS SSN,                           "
    SQL = SQL + "       to_char(b.birtdate, 'yyyymmdd') AS DOB                                                                 "
    SQL = SQL + "   from ( select patno,orddate,ordseqno,decode(discyn, 'X', '1', 'Y', '1', null) dcfg,ordcd,wardno            "
    SQL = SQL + "            from oram1.mdexmort                                                                               "
    SQL = SQL + "           where ordgrp = 'C1'                                                                                "
    SQL = SQL + "             and substr(slipcd, 1, 1) in ('L', 'N', 'P')                                                      "
    SQL = SQL + "           union all                                                                                          "
    SQL = SQL + "          select patno,orddate,ordseqno,decode(discyn, 'X', '1', 'Y', '1', null) dcfg,ordcd,wardno            "
    SQL = SQL + "            from oras1.mdbldort                                                                               "
    SQL = SQL + "           union all                                                                                          "
    SQL = SQL + "          select patno,orddate,ordseqno,decode(null, 'X', '1', 'Y', '1', null) dcfg,examcode ordcd,wardno     "
    SQL = SQL + "            from oras1.su2examt                                                                               "
    SQL = SQL + "           where ordgrp = 'C1'                                                                                "
    SQL = SQL + "             and substr(slipcode, 1, 1) in ('L', 'N', 'P')                                                    "
    SQL = SQL + "           union all                                                                                          "
    SQL = SQL + "          select patno,orddate,ordseqno,decode(null, 'X', '1', 'Y', '1', null) dcfg,examcode ordcd,wardno     "
    SQL = SQL + "            from oras1.sg2examt                                                                               "
    SQL = SQL + "           where ordgrp = 'C1'                                                                                "
    SQL = SQL + "             and substr(slipcode, 1, 1) in ('L', 'N', 'P')                                                    "
    SQL = SQL + "           union all                                                                                          "
    SQL = SQL + "          select patno,orddate,ordseqno,decode(discyn, 'X', '1', 'Y', '1', null) dcfg,ordcd,wardno            "
    SQL = SQL + "            from oras1.s2ord999                                                                               "
    SQL = SQL + "           where ordgrp = 'C1'                                                                                "
    SQL = SQL + "             and substr(slipcd, 1, 1) in ('L', 'N', 'P')  ) a,                                                "
    SQL = SQL + "        oraa1.appatbat b                                                                                      "
    SQL = SQL + "  where a.orddate between '" & vCase & "' and '" & vSearchKey & "'                                              "
    SQL = SQL + "   and ( a.dcfg='' or a.dcfg IS NULL )                                                                        "
    SQL = SQL + "   and a.ordcd IN ('C3710-1','C3710','C37101','C37102','C3710B','C3710C','C3710D','C3710E','C3710A')                   "
    SQL = SQL + "   and a.wardno = '" & vWardId & "'                                                                           "
    SQL = SQL + "   and a.patno = b.patno                                                                                     "

    SQL = SQL & " ORDER BY b." & mF_PTNM & ", SSN"

    GetSQLCol_Wm = SQL
End Function

Public Function GetSQLCol_Wm1(ByVal vCase As String, ByVal vSearchKey As String, _
                        Optional ByVal vWardId As String = "") As String
'병동용 채혈대상 검색
    Dim strSQL As String
    Dim SQL    As String
    
    SQL = " select DISTINCT b.patno AS ptid,                                                                             "
    SQL = SQL + "       b.patname AS ptnm,                                                                                     "
    SQL = SQL + "       crypto.dec('cpattern1',b.resno1) || crypto.dec('cpattern1',b.resno2) AS SSN,                           "
    SQL = SQL + "       to_char(b.birtdate, 'yyyymmdd') AS DOB                                                                 "
    SQL = SQL + "   from ( select patno,orddate,ordseqno,decode(discyn, 'X', '1', 'Y', '1', null) dcfg,ordcd,wardno,meddept as deptcd            "
    SQL = SQL + "            from oram1.mdexmort                                                                               "
    SQL = SQL + "           where ordgrp = 'C1'                                                                                "
    SQL = SQL + "             and substr(slipcd, 1, 1) in ('L', 'N', 'P') AND workarea IS NULL AND accdt IS NULL AND accseq IS NULL                                                      "
    SQL = SQL + "           union all                                                                                          "
    SQL = SQL + "          select patno,orddate,ordseqno,decode(discyn, 'X', '1', 'Y', '1', null) dcfg,ordcd,wardno,meddept as deptcd            "
    SQL = SQL + "            from oras1.mdbldort where workarea IS NULL AND accdt IS NULL AND accseq IS NULL                                                                              "
    SQL = SQL + "           union all                                                                                          "
    SQL = SQL + "          select patno,orddate,ordseqno,decode(null, 'X', '1', 'Y', '1', null) dcfg,examcode ordcd,wardno,meddept as deptcd     "
    SQL = SQL + "            from oras1.su2examt                                                                               "
    SQL = SQL + "           where ordgrp = 'C1'                                                                                "
    SQL = SQL + "             and substr(slipcode, 1, 1) in ('L', 'N', 'P') AND workarea IS NULL AND accdt IS NULL AND accseq IS NULL                                                   "
    SQL = SQL + "           union all                                                                                          "
    SQL = SQL + "          select patno,orddate,ordseqno,decode(null, 'X', '1', 'Y', '1', null) dcfg,examcode ordcd,wardno,meddept as deptcd     "
    SQL = SQL + "            from oras1.sg2examt                                                                               "
    SQL = SQL + "           where ordgrp = 'C1'                                                                                "
    SQL = SQL + "             and substr(slipcode, 1, 1) in ('L', 'N', 'P') AND workarea IS NULL AND accdt IS NULL AND accseq IS NULL                                                   "
    SQL = SQL + "           union all                                                                                          "
    SQL = SQL + "          select patno,orddate,ordseqno,decode(discyn, 'X', '1', 'Y', '1', null) dcfg,ordcd,wardno,meddept as deptcd            "
    SQL = SQL + "            from oras1.s2ord999                                                                               "
    SQL = SQL + "           where ordgrp = 'C1'                                                                                "
    SQL = SQL + "             and substr(slipcd, 1, 1) in ('L', 'N', 'P') AND workarea IS NULL AND accdt IS NULL AND accseq IS NULL ) a,                                                "
    SQL = SQL + "        oraa1.appatbat b                                                                                      "
    SQL = SQL + "  where a.orddate between '" & vCase & "' and '" & vSearchKey & "'                                              "
    SQL = SQL + "   and ( a.dcfg='' or a.dcfg IS NULL )                                                                        "
    SQL = SQL + "   and a.ordcd IN ('C3710-1','C3710','C37101','C37102','C3710B','C3710C','C3710D','C3710E','C3710A')                   "
    SQL = SQL + "   and deptcd = '" & vWardId & "'                                                                           "
    SQL = SQL + "   and a.patno = b.patno                                                                                     "

    SQL = SQL & " ORDER BY b." & mF_PTNM & ", SSN"

    GetSQLCol_Wm1 = SQL
End Function

Public Function GetSQLCol_Wm2(ByVal vCase As String, ByVal vSearchKey As String, _
                        Optional ByVal vWardId As String = "") As String
'병동용 채혈대상 검색
    Dim strSQL As String
    Dim SQL    As String
    
    SQL = " select DISTINCT b.patno AS ptid,                                                                             "
    SQL = SQL + "       b.patname AS ptnm,                                                                                     "
    SQL = SQL + "       crypto.dec('cpattern1',b.resno1) || crypto.dec('cpattern1',b.resno2) AS SSN,                           "
    SQL = SQL + "       to_char(b.birtdate, 'yyyymmdd') AS DOB                                                                 "
    SQL = SQL + "   from ( select patno,orddate,ordseqno,decode(discyn, 'X', '1', 'Y', '1', null) dcfg,ordcd,wardno,meddept as deptcd            "
    SQL = SQL + "            from oram1.mdexmort                                                                               "
    SQL = SQL + "           where ordgrp = 'C1'                                                                                "
    SQL = SQL + "             and substr(slipcd, 1, 1) in ('L', 'N', 'P')                                                    "
    SQL = SQL + "           union all                                                                                          "
    SQL = SQL + "          select patno,orddate,ordseqno,decode(discyn, 'X', '1', 'Y', '1', null) dcfg,ordcd,wardno,meddept as deptcd            "
    SQL = SQL + "            from oras1.mdbldort                                                                               "
    SQL = SQL + "           union all                                                                                          "
    SQL = SQL + "          select patno,orddate,ordseqno,decode(null, 'X', '1', 'Y', '1', null) dcfg,examcode ordcd,wardno,meddept as deptcd     "
    SQL = SQL + "            from oras1.su2examt                                                                               "
    SQL = SQL + "           where ordgrp = 'C1'                                                                                "
    SQL = SQL + "             and substr(slipcode, 1, 1) in ('L', 'N', 'P')                                                   "
    SQL = SQL + "           union all                                                                                          "
    SQL = SQL + "          select patno,orddate,ordseqno,decode(null, 'X', '1', 'Y', '1', null) dcfg,examcode ordcd,wardno,meddept as deptcd     "
    SQL = SQL + "            from oras1.sg2examt                                                                               "
    SQL = SQL + "           where ordgrp = 'C1'                                                                                "
    SQL = SQL + "             and substr(slipcode, 1, 1) in ('L', 'N', 'P')                                                  "
    SQL = SQL + "           union all                                                                                          "
    SQL = SQL + "          select patno,orddate,ordseqno,decode(discyn, 'X', '1', 'Y', '1', null) dcfg,ordcd,wardno,meddept as deptcd            "
    SQL = SQL + "            from oras1.s2ord999                                                                               "
    SQL = SQL + "           where ordgrp = 'C1'                                                                                "
    SQL = SQL + "             and substr(slipcd, 1, 1) in ('L', 'N', 'P') ) a,                                                "
    SQL = SQL + "        oraa1.appatbat b                                                                                      "
    SQL = SQL + "  where a.orddate between '" & vCase & "' and '" & vSearchKey & "'                                              "
    SQL = SQL + "   and ( a.dcfg='' or a.dcfg IS NULL )                                                                        "
    SQL = SQL + "   and a.ordcd IN ('C3710-1','C3710','C37101','C37102','C3710B','C3710C','C3710D','C3710E','C3710A')                   "
    SQL = SQL + "   and deptcd = '" & vWardId & "'                                                                           "
    SQL = SQL + "   and a.patno = b.patno                                                                                     "

    SQL = SQL & " ORDER BY b." & mF_PTNM & ", SSN"

    GetSQLCol_Wm2 = SQL
End Function

Public Function GetSQLCol(ByVal vCase As String, ByVal vSearchKey As String, _
                        Optional ByVal vWardId As String = "") As String
'병동용 채혈대상 검색
    Dim strSQL As String

    strSQL = " select distinct a." & mF_PTID & " as ptid, a." & mF_PTNM & " as ptnm," & _
                                  F_SSN2("a") & " as SSN, " & _
                                  F_DOB2("a") & " as DOB " & _
            " from " & mT_HIS001 & " a ," & mT_LAB101 & " b," & mT_LAB102 & " c " & _
            " where " & DBW("b.donefg =", "0") & _
            " and   " & DBW("b.bussdiv =", "2") & _
            " and    a." & mF_PTID & " = b.ptid"

'    strSQL = " select distinct a." & mF_PTID & " as ptid, a." & mF_PTNM & " as ptnm," & _
'                                  F_SSN2 & " as SSN, " & _
'                                  F_DOB2 & " as DOB " & _
'            " from " & mT_HIS001 & " a ," & mT_LAB101 & " b," & mT_LAB102 & " c " & _
'            " where " & DBW("b.donefg =", "0") & _
'            " and   " & DBW("b.bussdiv =", "2") & _
'            " and    a." & mF_PTID & " = b.ptid"
            
    If Val(vCase) Mod 2 = 1 Then
        strSQL = strSQL & " AND a." & mF_PTID & " = " & DBV(mF_PTID, vSearchKey) & " "
    Else
        strSQL = strSQL & " AND a." & mF_PTNM & " like '" & vSearchKey & "%' "
    End If

    strSQL = strSQL & " AND b.wardid = '" & vWardId & "'" & _
                " AND b.ptid=c.ptid" & _
                " AND b.orddt=c.orddt" & _
                " AND b.ordno=c.ordno" & _
                " AND (c.dcfg='' or c.dcfg is null)"


    If Val(vCase) Mod 2 = 1 Then
        strSQL = strSQL & " ORDER BY a." & mF_PTID
    Else
        strSQL = strSQL & " ORDER BY a." & mF_PTNM & ", SSN"
    End If

    GetSQLCol = strSQL
End Function

Public Function GetSQLPt(Optional ByVal vPtID As String = "") As String
    GetSQLPt = " SELECT " & mF_PTID & " as ptid, " & mF_PTNM & " as ptnm, " & _
                           F_SSN2 & " as ssn,  " & F_DOB2 & " as dob, " & _
                           F_SEX2 & " as sex,  " & mF_ZIPCODE & " as zipcd, " & _
                           mF_ADDRESS & " as addr, " & mF_TEL & " as tel, " & mF_TMPDIV & " as tmpdiv " & _
             " FROM   " & mT_HIS001 & " "
    
    GetSQLPt = GetSQLPt & " WHERE  " & DBW(mF_PTID, vPtID, 2)
End Function

Public Function GetSQLPtList() As String
    GetSQLPtList = " SELECT " & mF_PTID & " as ptid, " & mF_PTNM & " as ptnm, " & _
                           F_SSN2 & " as ssn,  " & F_DOB2 & " as dob, " & _
                           F_SEX2 & " as sex,  " & mF_ZIPCODE & " as zipcd, " & _
                           mF_ADDRESS & " as addr, " & mF_TEL & " as tel, " & mF_TMPDIV & " as tmpdiv " & _
             " FROM   " & mT_HIS001 & " "
End Function

'환자 검색 이벤트용
Public Sub LoadSearchForm()
    Set objSearchForm = frmSearchPt
    objSearchForm.Show vbModal
End Sub

Private Sub objSearchForm_SelectedID(ByVal vPtID As String)
    Unload objSearchForm
    Set objSearchForm = Nothing
    
    RaiseEvent SelectedId(vPtID)
End Sub

'Private Sub objSearchForm_Selected(ByVal vPtInfo As clsPatient)
'    RaiseEvent Selected(vPtInfo)
'
'    Set objSearchForm = Nothing
'End Sub

Private Sub Class_Initialize()
    mvarPtID = ""  '환자아뒤
    mvarPtNm = ""  '환자명
    mvarSSN = ""   '주민번호
    mvarAge = ""   '나이
    mvarSex = ""   '성별
    mvarDob = ""   '생년월일
    mvarZIPCD = "" '우편번호
    mvarADDR = ""  '주소
    mvarTEL = ""   '전화번호
    mvarPtDiv = "" '환자유형
    mvarTMPDIV = "" '검진 임시ID 구분
    
    '재원 마스터
    mvarINPTID = "" '환자아뒤
    mvarBedIndt = ""   '입원일자
    mvarBEDINTM = ""   '입원시간
    mvarBEDOUTDT = ""  '퇴원일자
    mvarBEDOUTTM = ""  '퇴원시간
    mvarDeptCd = ""  '진료과
    mvarBEDPATH = ""   '입원경로
    mvarINDISEACD = "" '입원상병
    mvarWardId = ""  '병동아뒤
    mvarROOMID = ""  '병실아뒤
    mvarBedID = ""   '베드아뒤
    mvarMajdoct = ""   '주치의아뒤
    
    '부서 마스터
'    mvarDEPTCD = "" '부서코드
    mvarDeptNm = "" '부서명
'    mvarDEPTDIV = ""   '부서구분
'
'    '병상 마스터
'    mvarWARDID = "" '병동아뒤
'    mvarWARDNM = "" '병동명
'    mvarROOMID = "" '병실아뒤
'    mvarBEDID = "" '베드아뒤
'
'    '의사 마스터 및 직원마스터
'    mvarDOCTID = "" '의사아뒤
    mvarDoctNm = "" '의사명
'    mvarEMPID = "" '직원아뒤
'    mvarEMPNM = "" '직원명
'    mvarEMPDIV = ""    '직종 구분
'    mvarEXPDT = "" '퇴사일
'
'    '상병 마스터
'    mvarICD = ""   '상병코드
'    mvarIENM = ""  '영문상병명
'    mvarIKNM = ""  '국문상병명
'
'    '수가 마스터
'    mvarAMTCD = "" '수가코드
'    mvarAMTNM = "" '코드명
'
'    '수술코드 마스터
'    mvarOCD = ""   '수술코드
'    mvarOENM = ""  '영문수술명
'    mvarOKNM = ""  '국문수술명
    
    '기타 정보
    mvarINADMISSION = False '입원 여부
    mvarEXISTADM = False
End Sub
