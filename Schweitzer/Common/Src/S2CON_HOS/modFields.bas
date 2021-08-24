Attribute VB_Name = "modFields"
Option Explicit

Public IsSetFields      As Boolean

'Public mPROJECT_HOSCD   As String   '적용병원코드 (가톨릭 성모자애병원)
Public mF_PTID          As String   '환자ID
Public mF_PTNM          As String   '환자명
Public mF_SSN           As String   '주민등록번호
Public mF_AGE           As String   '나이
Public mF_SEX           As String   '성별
Public mF_DOB           As String   '생년월일
Public mF_ZIPCODE       As String   '우편번호
Public mF_ADDRESS       As String   '주소
Public mF_TEL           As String   '전화번호
Public mF_HPTEL         As String   '휴대폰번호
Public mF_TMPDIV        As String   '검진구분 '1'검진

Public mF_INPTID        As String   '재원환자ID
Public mF_BEDOUTDT      As String   '퇴원일
Public mF_BEDOUTTM      As String   '퇴원시간
Public mF_BEDINDT       As String   '입원일
Public mF_BEDINTM       As String   '입원시간
Public mF_PTDEPTCD      As String   '재원환자진료과
Public mF_PTWARDID      As String   '입원병동ID
Public mF_PTROOMID      As String   '입원병실ID
Public mF_PTBEDID       As String   '입원병상ID
Public mF_PTDISEASE     As String   '입원상병코드
Public mF_PTDIV         As String   '환자구분
Public mF_MAJDOCT       As String   '주치의ID


Public mF_DEPTCD        As String   '부서코드
Public mF_DEPTNM        As String   '부서명
Public mF_DEPTDIV       As String   '부서구분
Public mF_BLDGB         As String   '건물구분

Public mF_WARDID        As String   '병동ID
Public mF_WARDNM        As String   '병동명
Public mF_ROOMID        As String   '병실ID
Public mF_BEDID         As String   '병상ID

Public mF_DOCTID        As String   '의사ID
Public mF_DOCTNM        As String   '의사명
Public mF_EMPID         As String   '직원ID
Public mF_EMPNM         As String   '직원명
Public mF_EMPDIV        As String   'JOB 구분
Public mF_EMPDIV2       As String   'JOB 구분2
Public mF_NURSEDIV      As String   '간호사 구분
Public mF_EXPDT         As String   '퇴사일
Public mF_ICD           As String   '상병코드
Public mF_IENM          As String   '상병영문명
Public mF_IKNM          As String   '상병한글명
Public mF_OCD           As String   '수술코드
Public mF_ONM           As String   '수술명
Public mF_ODIV          As String   '구분코드
Public mF_AMTCD         As String   '수가코드
Public mF_AMTNM         As String   '수가명
Public mF_MATCD         As String   'Match코드
Public mF_ANTNM         As String   '수가명 ----> 나중에 지울것

Public mFUNC_SUBSTR     As String   'Oracle:substr, Sybase & SQL Server:substring
Public mFUNC_CONCAT     As String   'Oracle: ||,    Sybase & SQL Server: +

Public Sub SetFields()
'
'    mPROJECT_HOSCD = ReadINI("FIELD", "PROJECT_HOSCD", "")                 '적용병원코드 ("02":가톨릭 성모자애병원)
'
'his001(h1ptntinfo) : 환자기본마스터
    mF_PTID = ReadINI("FIELD", "F_PTID", "")                     '환자ID
    mF_PTNM = ReadINI("FIELD", "F_PTNM", "")                    '환자명
    mF_SSN = ReadINI("FIELD", "F_SSN", "")                      '주민등록번호
    mF_AGE = ReadINI("FIELD", "F_AGE", "")                      '나이
    mF_SEX = ReadINI("FIELD", "F_SEX", "")                      '성별
    mF_PTDIV = ReadINI("FIELD", "F_PTDIV", "")                  '환자구분
    mF_DOB = ReadINI("FIELD", "F_DOB", "")                      '생년월일
    mF_ZIPCODE = ReadINI("FIELD", "F_ZIPCODE", "")              '우편번호
    mF_ADDRESS = ReadINI("FIELD", "F_ADDRESS", "")              '주소
    mF_TEL = ReadINI("FIELD", "F_TEL", "")                      '전화번호
    mF_HPTEL = ReadINI("FIELD", "F_HPTEL", "")                  '휴대전화번호
    mF_TMPDIV = ReadINI("FIELD", "F_TMPDIV", "")                '검진구분

'his002(h1admin) : 재원마스터 --> h7lab501 사용하기로 함. 2001.1.17 kmk
    mF_INPTID = ReadINI("FIELD", "F_INPTID", "")                '재원환자ID
    mF_BEDOUTDT = ReadINI("FIELD", "F_BEDOUTDT", "")            '"dchg_ymd"
    mF_BEDOUTTM = ReadINI("FIELD", "F_BEDOUTTM", "")
    mF_BEDINDT = ReadINI("FIELD", "F_BEDINDT", "")              '"adm_ymd"
    mF_BEDINTM = ReadINI("FIELD", "F_BEDINTM", "")
    mF_PTDEPTCD = ReadINI("FIELD", "F_PTDEPTCD", "")            '재원환자진료과
    mF_PTWARDID = ReadINI("FIELD", "F_PTWARDID", "")            '입원병동ID
    mF_PTROOMID = ReadINI("FIELD", "F_PTROOMID", "")            '입원병실ID
    mF_PTBEDID = ReadINI("FIELD", "F_PTBEDID", "")              '입원침상ID
    mF_PTDISEASE = ReadINI("FIELD", "F_PTDISEASE", "")          '입원상병코드
    mF_MAJDOCT = ReadINI("FIELD", "F_MAJDOCT", "")              '주치의ID

'his003(hzdept) : 부서마스터
    mF_DEPTCD = ReadINI("FIELD", "F_DEPTCD", "")                '부서코드
    mF_DEPTNM = ReadINI("FIELD", "F_DEPTNM", "")                '부서명
    mF_DEPTDIV = ReadINI("FIELD", "F_DEPTDIV", "")              '부서구분
    mF_BLDGB = ReadINI("FIELD", "F_BLDGB", "")                  '건물구분
    
'his004(hzdept) : 병상마스터
    mF_WARDID = ReadINI("FIELD", "F_WARDID", "")                '병동ID
    mF_WARDNM = ReadINI("FIELD", "F_WARDNM", "")                '병동명
    mF_ROOMID = ReadINI("FIELD", "F_ROOMID", "")                '병실ID
    mF_BEDID = ReadINI("FIELD", "F_BEDID", "")                  '병상ID

'his005(hzempl) : 의사마스터
    mF_DOCTID = ReadINI("FIELD", "F_DOCTID", "")                '의사ID
    mF_DOCTNM = ReadINI("FIELD", "F_DOCTNM", "")                '의사명
     
    mF_EMPID = ReadINI("FIELD", "F_EMPID", "")                  '직원ID
    mF_EMPNM = ReadINI("FIELD", "F_EMPNM", "")                  '직원명
    mF_EMPDIV = ReadINI("FIELD", "F_EMPDIV", "")                'JOB 구분
    mF_EMPDIV2 = ReadINI("FIELD", "F_EMPDIV2", "")              'JOB 구분2
    mF_EXPDT = ReadINI("FIELD", "F_EXPDT", "")                  '퇴사일
    mF_NURSEDIV = ReadINI("FIELD", "F_NURSEDIV", "")            '간호사구분
'his006(h2diag) : 상병마스터
    mF_ICD = ReadINI("FIELD", "F_ICD", "")                      '상병코드
    mF_IENM = ReadINI("FIELD", "F_IENM", "")                    '상병영문명
    mF_IKNM = ReadINI("FIELD", "F_IKNM", "")                    '상병한글명

'his007(h1actmat) : 수술마스터(medfee_class_cd = '21')
    mF_OCD = ReadINI("FIELD", "F_OCD", "")                      '수술코드
    mF_ONM = ReadINI("FIELD", "F_ONM", "")                      '수술명
    mF_ODIV = ReadINI("FIELD", "F_ODIV", "")                    '구분코드

'his008(h1actmat) : 수가마스터
    mF_AMTCD = ReadINI("FIELD", "F_AMTCD", "")                  '수가코드
    mF_AMTNM = ReadINI("FIELD", "F_AMTNM", "")                  '수가명
    mF_MATCD = ReadINI("FIELD", "F_MATCD", "")                  'Match코드
    
    mFUNC_SUBSTR = ReadINI("FIELD", "FUNC_SUBSTR", "")
    mFUNC_CONCAT = ReadINI("FIELD", "FUNC_CONCAT", "")
End Sub

