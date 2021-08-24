VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFields"
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

Public Property Get F_PTNM() As String       '환자명
    F_PTNM = mF_PTNM
End Property

Public Property Get F_SSN() As String        '주민등록번호
    F_SSN = mF_SSN
End Property

Public Property Get F_AGE() As String        '나이
    F_AGE = mF_AGE
End Property

Public Property Get F_SEX() As String        '성별
    F_SEX = mF_SEX
End Property

Public Property Get F_DOB() As String        '생년월일
    F_DOB = mF_DOB
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

Public Property Get F_BEDOUTDT() As String   '퇴원일
    F_BEDOUTDT = mF_BEDOUTDT
End Property

Public Property Get F_BEDOUTTM() As String   '퇴원시
    F_BEDOUTTM = mF_BEDOUTTM
End Property

Public Property Get F_BEDINDT() As String    '입원일
    F_BEDINDT = mF_BEDINDT
End Property

Public Property Get F_BEDINTM() As String    '입원시
    F_BEDINTM = mF_BEDINTM
End Property

Public Property Get F_PTDEPTCD() As String   '재원환자진료과
    F_PTDEPTCD = mF_PTDEPTCD
End Property

Public Property Get F_PTWARDID() As String   '입원병동ID
    F_PTWARDID = mF_PTWARDID
End Property

Public Property Get F_PTROOMID() As String   '입원병실ID
    F_PTROOMID = mF_PTROOMID
End Property

Public Property Get F_PTBEDID() As String    '입원병상ID
    F_PTBEDID = mF_PTBEDID
End Property

Public Property Get F_PTDISEASE() As String  '입원상병코드
    F_PTDISEASE = mF_PTDISEASE
End Property

Public Property Get F_PTDIV() As String      '환자구분
    F_PTDIV = mF_PTDIV
End Property

Public Property Get F_MAJDOCT() As String    '주치의ID
    F_MAJDOCT = mF_MAJDOCT
End Property

Public Property Get F_DEPTCD() As String     '부서코드
    F_DEPTCD = mF_DEPTCD
End Property

Public Property Get F_DEPTNM() As String     '부서명
    F_DEPTNM = mF_DEPTNM
End Property

Public Property Get F_DEPTDIV() As String    '부서구분
    F_DEPTDIV = mF_DEPTDIV
End Property

Public Property Get F_BLDGB() As String      '건물구분
    F_BLDGB = mF_BLDGB
End Property

Public Property Get F_WARDID() As String     '병동ID
    F_WARDID = mF_WARDID
End Property

Public Property Get F_WARDNM() As String     '병동명
    F_WARDNM = mF_WARDNM
End Property

Public Property Get F_ROOMID() As String     '병실ID
    F_ROOMID = mF_ROOMID
End Property

Public Property Get F_BEDID() As String      '병상ID
    F_BEDID = mF_BEDID
End Property

Public Property Get F_DOCTID() As String     '의사ID
    F_DOCTID = mF_DOCTID
End Property

Public Property Get F_DOCTNM() As String     '의사명
    F_DOCTNM = mF_DOCTNM
End Property

Public Property Get F_EMPID() As String      '직원ID
    F_EMPID = mF_EMPID
End Property

Public Property Get F_EMPNM() As String      '직원명
    F_EMPNM = mF_EMPNM
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

Public Property Get F_EXPDT() As String      '직원명
    F_EXPDT = mF_EXPDT
End Property

Public Property Get F_ICD() As String        '상병코드
    F_ICD = mF_ICD
End Property

Public Property Get F_IENM() As String       '상병영문명
    F_IENM = mF_IENM
End Property

Public Property Get F_IKNM() As String       '상병한글명
    F_IKNM = mF_IKNM
End Property

Public Property Get F_OCD() As String        '수술코드
    F_OCD = mF_OCD
End Property

Public Property Get F_ONM() As String        '수술명
    F_ONM = mF_ONM
End Property

Public Property Get F_ODIV() As String       '구분코드
    F_ODIV = mF_ODIV
End Property

Public Property Get F_AMTCD() As String      '수가코드
    F_AMTCD = mF_AMTCD
End Property

Public Property Get F_AMTNM() As String      '수가명
    F_AMTNM = mF_AMTNM
End Property

Public Property Get F_MATCD() As String      'Match코드
    F_MATCD = mF_MATCD
End Property

Public Property Get FUNC_SUBSTR() As String      'Oracle:substr, Sybase & SQL Server:substring
    FUNC_SUBSTR = mFUNC_SUBSTR
End Property

Public Property Get FUNC_CONCAT() As String      'Oracle: ||,    Sybase & SQL Server: +
    FUNC_CONCAT = mFUNC_CONCAT
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
Public Function F_SSN2(Optional ByVal tA As String) As String
    If IsMissing(tA) Then
        F_SSN2 = "crypto.dec('cpattern1',resno1) || crypto.dec('cpattern1',resno2)"
    Else
        If tA = "" Then
            F_SSN2 = "crypto.dec('cpattern1',resno1) || crypto.dec('cpattern1',resno2)"
        Else
            F_SSN2 = "crypto.dec('cpattern1'," & tA & ".resno1) || " & "crypto.dec('cpattern1'," & tA & ".resno2)"
        End If
    End If
End Function

' 2017-06-28 신규암호화
Public Function F_SEX2(Optional ByVal tA As String) As String
    If IsMissing(tA) Then
        F_SEX2 = "substr(crypto.dec('cpattern1',resno2), 1, 1)"
    Else
        If tA = "" Then
            F_SEX2 = "substr(crypto.dec('cpattern1',resno2), 1, 1)"
        Else
            F_SEX2 = "substr(crypto.dec('cpattern1'," & tA & ".resno2),1,1)"
        End If
    End If
End Function

' 구암호화
'Public Function F_SSN2(Optional ByVal tA As String) As String
'    If IsMissing(tA) Then
'        F_SSN2 = "CryptIT.decrypt(resno1,'pmc1898') || CryptIT.decrypt(resno2,'pmc1898')"
'    Else
'        If tA = "" Then
'            F_SSN2 = "resno1 || resno2"
'        Else
'            F_SSN2 = tA & ".resno1 || " & tA & ".resno2"
'        End If
'    End If
'End Function

' 구암호화
'Public Function F_SEX2(Optional ByVal tA As String) As String
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

'' 2017.05.23 암호화 해제로 인한 수정
'Public Function F_SSN2(Optional ByVal tA As String) As String
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
'
'' 2017.05.23  해제로 인한 수정
'Public Function F_SEX2(Optional ByVal tA As String) As String
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

Public Function F_DOB2(Optional ByVal tA As String) As String
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

Public Function F_BEDOUTDT2(Optional ByVal tA As String) As String
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

Public Function F_BEDINDT2(Optional ByVal tA As String) As String
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

Private Sub Class_Initialize()
    
    If IsSetFields = False Then Call SetFields
    IsSetFields = True

End Sub
