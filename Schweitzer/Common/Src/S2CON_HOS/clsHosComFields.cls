VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHosComFields"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

''** Fields Name **

'Public Property Get PROJECT_HOSCD() As String '적용병원코드 (가톨릭 성모자애병원)
'    PROJECT_HOSCD = mPROJECT_HOSCD
'End Property

Public Property Get F_PTID() As String       '환자ID
    F_PTID = mF_PTID
End Property
Public Property Let F_PTID(ByVal vData As String)

End Property
Public Property Get F_PTNM() As String       '환자명
    F_PTNM = mF_PTNM
End Property
Public Property Let F_PTNM(ByVal vData As String)

End Property
Public Property Get F_SSN() As String        '주민등록번호
    F_SSN = mF_SSN
End Property
Public Property Let F_SSN(ByVal vData As String)

End Property
Public Property Get F_AGE() As String        '나이
    F_AGE = mF_AGE
End Property
Public Property Let F_AGE(ByVal vData As String)

End Property
Public Property Get F_SEX() As String        '성별
    F_SEX = mF_SEX
End Property
Public Property Let F_SEX(ByVal vData As String)

End Property
Public Property Get F_DOB() As String        '생년월일
    F_DOB = mF_DOB
End Property
Public Property Let F_DOB(ByVal vData As String)

End Property

Public Property Get F_ZIPCODE() As String    '우편번호
    F_ZIPCODE = mF_ZIPCODE
End Property

Public Property Get F_ADDRESS() As String    '주소
    F_ADDRESS = mF_ADDRESS
End Property

Public Property Get F_TEL() As String        '전화번호
    F_TEL = mF_TEL
End Property

Public Property Get F_TMPDIV() As String     '검진구분 '1' 검진
    F_TMPDIV = mF_TMPDIV
End Property

Public Property Get F_INPTID() As String     '재원환자ID
    F_INPTID = mF_INPTID
End Property
Public Property Let F_INPTID(ByVal vData As String)

End Property

Public Property Get F_BEDOUTDT() As String   '퇴원일
    F_BEDOUTDT = mF_BEDOUTDT
End Property

Public Property Get F_BEDOUTTM() As String   '퇴원시
    F_BEDOUTTM = mF_BEDOUTTM
End Property

Public Property Let F_BEDOUTDT(ByVal vData As String)

End Property

Public Property Get F_BEDINDT() As String    '입원일
    F_BEDINDT = mF_BEDINDT
End Property

Public Property Get F_BEDINTM() As String    '입원시
    F_BEDINTM = mF_BEDINTM
End Property

Public Property Let F_BEDINDT(ByVal vData As String)

End Property
Public Property Get F_PTDEPTCD() As String   '재원환자진료과
    F_PTDEPTCD = mF_PTDEPTCD
End Property
Public Property Let F_PTDEPTCD(ByVal vData As String)

End Property
Public Property Get F_PTWARDID() As String   '입원병동ID
    F_PTWARDID = mF_PTWARDID
End Property
Public Property Let F_PTWARDID(ByVal vData As String)

End Property
Public Property Get F_PTROOMID() As String   '입원병실ID
    F_PTROOMID = mF_PTROOMID
End Property
Public Property Let F_PTROOMID(ByVal vData As String)

End Property
Public Property Get F_PTBEDID() As String    '입원병상ID
    F_PTBEDID = mF_PTBEDID
End Property
Public Property Let F_PTBEDID(ByVal vData As String)

End Property
Public Property Get F_PTDISEASE() As String  '입원상병코드
    F_PTDISEASE = mF_PTDISEASE
End Property
Public Property Let F_PTDISEASE(ByVal vData As String)

End Property
Public Property Get F_PTDIV() As String      '환자구분
    F_PTDIV = mF_PTDIV
End Property
Public Property Let F_PTDIV(ByVal vData As String)

End Property

Public Property Get F_MAJDOCT() As String    '주치의ID
    F_MAJDOCT = mF_MAJDOCT
End Property

Public Property Let F_MAJDOCT(ByVal vData As String)

End Property
Public Property Get F_DEPTCD() As String     '부서코드
    F_DEPTCD = mF_DEPTCD
End Property
Public Property Let F_DEPTCD(ByVal vData As String)

End Property
Public Property Get F_DEPTNM() As String     '부서명
    F_DEPTNM = mF_DEPTNM
End Property
Public Property Let F_DEPTNM(ByVal vData As String)

End Property
Public Property Get F_DEPTDIV() As String    '부서구분
    F_DEPTDIV = mF_DEPTDIV
End Property
Public Property Let F_DEPTDIV(ByVal vData As String)

End Property
Public Property Get F_BLDGB() As String      '건물구분
    F_BLDGB = mF_BLDGB
End Property
Public Property Let F_BLDGB(ByVal vData As String)

End Property
Public Property Get F_WARDID() As String     '병동ID
    F_WARDID = mF_WARDID
End Property
Public Property Let F_WARDID(ByVal vData As String)

End Property
Public Property Get F_WARDNM() As String     '병동명
    F_WARDNM = mF_WARDNM
End Property
Public Property Let F_WARDNM(ByVal vData As String)

End Property
Public Property Get F_ROOMID() As String     '병실ID
    F_ROOMID = mF_ROOMID
End Property
Public Property Let F_ROOMID(ByVal vData As String)

End Property
Public Property Get F_BEDID() As String      '병상ID
    F_BEDID = mF_BEDID
End Property
Public Property Let F_BEDID(ByVal vData As String)

End Property
Public Property Get F_DOCTID() As String     '의사ID
    F_DOCTID = mF_DOCTID
End Property
Public Property Let F_DOCTID(ByVal vData As String)

End Property
Public Property Get F_DOCTNM() As String     '의사명
    F_DOCTNM = mF_DOCTNM
End Property
Public Property Let F_DOCTNM(ByVal vData As String)

End Property
Public Property Get F_EMPID() As String      '직원ID
    F_EMPID = mF_EMPID
End Property
Public Property Let F_EMPID(ByVal vData As String)

End Property
Public Property Get F_EMPNM() As String      '직원명
    F_EMPNM = mF_EMPNM
End Property
Public Property Let F_EMPNM(ByVal vData As String)

End Property

Public Property Get F_NURSEDIV() As String     '간호사구분
    F_NURSEDIV = mF_NURSEDIV
End Property
Public Property Get F_EMPDIV2() As String     'JOB 구분2
    F_EMPDIV2 = mF_EMPDIV2
End Property

Public Property Get F_EMPDIV() As String     'JOB 구분
    F_EMPDIV = mF_EMPDIV
End Property
Public Property Let F_EMPDIV(ByVal vData As String)

End Property
Public Property Get F_EXPDT() As String      '직원명
    F_EXPDT = mF_EXPDT
End Property
Public Property Let F_EXPDT(ByVal vData As String)

End Property
Public Property Get F_ICD() As String        '상병코드
    F_ICD = mF_ICD
End Property
Public Property Let F_ICD(ByVal vData As String)

End Property
Public Property Get F_IENM() As String       '상병영문명
    F_IENM = mF_IENM
End Property
Public Property Let F_IENM(ByVal vData As String)

End Property
Public Property Get F_IKNM() As String       '상병한글명
    F_IKNM = mF_IKNM
End Property
Public Property Let F_IKNM(ByVal vData As String)

End Property
Public Property Get F_OCD() As String        '수술코드
    F_OCD = mF_OCD
End Property
Public Property Let F_OCD(ByVal vData As String)

End Property
Public Property Get F_ONM() As String        '수술명
    F_ONM = mF_ONM
End Property
Public Property Let F_ONM(ByVal vData As String)

End Property
Public Property Get F_ODIV() As String       '구분코드
    F_ODIV = mF_ODIV
End Property
Public Property Let F_ODIV(ByVal vData As String)

End Property
Public Property Get F_AMTCD() As String      '수가코드
    F_AMTCD = mF_AMTCD
End Property
Public Property Let F_AMTCD(ByVal vData As String)

End Property
Public Property Get F_AMTNM() As String      '수가명
    F_AMTNM = mF_AMTNM
End Property
Public Property Let F_AMTNM(ByVal vData As String)

End Property
Public Property Get F_MATCD() As String      'Match코드
    F_MATCD = mF_MATCD
End Property
Public Property Let F_MATCD(ByVal vData As String)

End Property
Public Property Get F_ANTNM() As String      '수가명 ----> 나중에 지울것

End Property
Public Property Let F_ANTNM(ByVal vData As String)
End Property
Public Property Get FUNC_SUBSTR() As String      'Oracle:substr, Sybase & SQL Server:substring
    FUNC_SUBSTR = mFUNC_SUBSTR
End Property
Public Property Let FUNC_SUBSTR(ByVal vData As String)

End Property
Public Property Get FUNC_CONCAT() As String      'Oracle: ||,    Sybase & SQL Server: +
    FUNC_CONCAT = mFUNC_CONCAT
End Property
Public Property Let FUNC_CONCAT(ByVal vData As String)

End Property

'###########################################################################################
'이 아래부터 병원별로 바뀌어야 될 부분
'###########################################################################################
Public Function F_SSN2(Optional ByVal tA As String) As String
    If IsMissing(tA) Then
        F_SSN2 = "sujumin1 || sujumin2"
    Else
        If tA = "" Then
            F_SSN2 = "sujumin1 || sujumin2"
        Else
            F_SSN2 = tA & ".sujumin1 || " & tA & ".sujumin2"
        End If
    End If
End Function

Public Function F_DOB2(Optional ByVal tA As String) As String
    If IsMissing(tA) Then
        F_DOB2 = "birth"
    Else
        If tA = "" Then
            F_DOB2 = "birth"
        Else
            F_DOB2 = tA & ".birth"
        End If
    End If
End Function

Public Function F_SEX2(Optional ByVal tA As String) As String
    If IsMissing(tA) Then
        F_SEX2 = "substr(sujumin2, 1, 1)"
    Else
        If tA = "" Then
            F_SEX2 = "substr(sujumin2, 1, 1)"
        Else
            F_SEX2 = "substr(" & tA & ".sujumin2, 1, 1)"
        End If
    End If
End Function

Public Function F_BEDOUTDT2(Optional ByVal tA As String) As String
    If IsMissing(tA) Then
        F_BEDOUTDT2 = "toiwon_date"
    Else
        If tA = "" Then
            F_BEDOUTDT2 = "toiwon_date"
        Else
            F_BEDOUTDT2 = tA & ".toiwon_date"
        End If
    End If
End Function

Public Function F_BEDINDT2(Optional ByVal tA As String) As String
    If IsMissing(tA) Then
        F_BEDINDT2 = "ipwon_date"
    Else
        If tA = "" Then
            F_BEDINDT2 = "ipwon_date"
        Else
            F_BEDINDT2 = tA & ".ipwon_date"
        End If
    End If
End Function

Public Function FUNC_CONVERT(ByVal strTp As String, ByVal strFld As String) As String
    Select Case UCase(strTp)
        Case "CHAR":
            FUNC_CONVERT = " to_char(" & strFld & ") "
        Case "INT":
            FUNC_CONVERT = " to_number(" & strFld & ") "
    End Select
End Function

Private Sub Class_Initialize()
    
    If IsSetFields = False Then Call SetFields
    IsSetFields = True

End Sub
