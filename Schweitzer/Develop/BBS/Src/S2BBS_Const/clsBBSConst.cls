VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBBSConst"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'처방 상태(Status)
Public Enum BBSOrderStatus
    stsORDER = 0        '처방
    stsCOLLECT          '채혈
    stsACCESS           '접수
    stsINPROCESS        '진행중(Cross-Matching결과등록시 )
    stsEnd              '완료(pheresis처방)
End Enum

'2001-11-22추가 : 호환성 유지를 위하여 하나 더 만듦.
Public Enum BBSOrdStatus
    stsORDER = 0        '처방
    stsCOLLECT          '채혈
    stsACCESS           '접수
    stsREQUEST          '수혈요청
    stsINPROCESS        '진행중(Cross-Matching결과등록시 )
    stsEnd              '완료(pheresis처방)
End Enum

'혈액 상태(Status)
Public Enum BBSBloodStatus
    stsENTER = 0        '입고
    stsRETURN           '반환
    stsASSIGN           'Assign
    stsDELIVERY         '출고
    stsEXPIRE           '폐기
    stsBAG              '회수
End Enum

Public Enum BBSOrderStatusColor
    clrOrder = 0           'black
    clrCOLLECT = 65280     'green
    clrACCESS = 16711680   'blue
    clrINPROCESS = 255     'red
    cIrEND = 12583104
End Enum

Public Enum BBSOrdStatusColor
    clrOrder = 0           'black
    clrCOLLECT = 65280     'green
    clrACCESS = 16711680   'blue
    clrREQUEST = &HFF      'red
    clrINPROCESS = &H4A4189
    cIrEND = &H8000&
End Enum

'결과타입
'검체추가요청 처리여부
Public Enum BBSSpcAddFg
    stsORDER = 0
    stsCOLLECT
End Enum
'보관장소 상태(BBS206)
Public Enum BBSSaveStatue
    stsUsed = 0         '사용가능
    stsNotUsed          '사용중
    stsNotUsedPoss      '폐기후사용
End Enum

'취소여부
Public Enum BBSCancelStatus
    stsNotCancel = 0    '취소아님
    stsCancel           '취소상태
End Enum
'폐기여부
Public Enum BBSExpFg
    stsNotExpFg = 0     '폐기 아님
    stsExpfg            '폐기
End Enum
'입원구분
Public Enum BBSBUSSDIV
    stsNotBed = 1
    stsBed
    stsEr
End Enum

Public Enum DonorStatus
    '헌혈자상태(임의,자가,지정)
    stsAccessSave = 1
    stsAccessVerify = 2
    stsAskSave = 3
    stsAskVerify = 4
    stsDonation = 5
    stsFinish = 6
    stsPrint = 7
    
    '헌혈자상태(pheresis)
End Enum

Public Enum XMethod
    xmMAJOR = 0
    xmMINOR
    xmBOTH
    xmNONE
End Enum


Public CD2_Buildings        As String

'================================
'        마스터성 기본정의
'================================
Public PRESENTDATE_FORMAT   As String               '날짜형식="YYYYMMDD"
Public PRESENTTIME_FORMAT   As String               '날짜형식="HHMMSS"

Public C_WORKAREA           As String               ' Workarea='B'
Public BC2_KEEP_HOUR        As String               ' xm검체보관일수(시간)
Public BC2_XM_STEP          As String               ' xm검사 step
Public BN_SPC_NO            As String               ' 검체번호
Public BN_ACC_NO            As String               ' 접수번호
Public BC2_COMPO            As String               ' 혈액제제
Public BC2_BRANCH           As String               ' 자병원코드
Public BC2_LOCAL            As String               ' local 병원
Public BC2_BLD_SRC          As String               ' 혈액원
Public BC2_EXP_RESON        As String               ' 폐기사유
Public BC2_ACTIVE_BLD_SRC   As String               ' 거래 혈액원
Public BC2_ACCNO_CRITERION  As String               ' 접수번호 생성기준
Public BC2_CENTER           As String               ' 센터코드
Public BC2_Trans_Reason     As String               ' 수혈사유
Public BC2_REACTION_REASON  As String               ' 수혈 부작용목록
Public BC2_REACTION_TEST    As String               ' 수혈 부작용시 검사항목
Public BC2_CANCELRSN        As String               ' 접수취소사유
Public BC2_REACTION_TEST2   As String               ' 수혈전 검사결과 일수
Public BC2_RETURN_HOUR      As String               ' 출고후 반환 가능시간
Public BC2_COMPO_GROUP      As String               ' 혈액제제 통계그룹
Public BC2_File_Server      As String           ' 파일서버위치
Public BC2_SPC_DEFAULT      As String               ' 검체번호앞에 붙는 문자
Public BC2_CompoValue       As String               ' 혈액제제별 가격설정

'================================
'        ABO관련 정의
'================================
Public BC2_ABO_RH           As String               ' ABO와 RH코드 (혈액출고시 한번 발생시키기 위함) cdval1 코드
Public BC2_RST_ABO          As String               ' ABO결과코드
Public BC2_RST_RH           As String               ' RH 결과코드
Public BC2_RST_ABOSUB       As String               ' ABO SUBTYPE  결과코드
Public BC2_RST_RHSUB        As String               ' RH  SUBGROUP 결과코드
Public BC2_RST_DU           As String               ' DU Test 결과코드
Public BC2_ABO_TEST         As String               ' ABO결과 항목
Public BC2_BARCODE          As String               ' 바코드 출력시 붙는 문자
Public BC2_ADDSPC           As String               ' 추가채혈 여부
Public BC2_ORDVAL           As String               ' 수술
Public BC2_REASONVAL        As String               ' 수혈사유
'================================
'   헌혈자 검사의뢰에서 사용
'================================
Public BN_DONOR_ID          As String               ' 헌혈자 접수 ID
Public CD1_MULTISPC         As String               ' 복수검체
Public CD2_SPECIMEN         As String               ' 검체
Public CD2_WORKAREA         As String               ' Work Area
Public CD2_SGROUP           As String               ' 검체군
Public NO_SPECIMEN          As String               ' 검체번호
Public NO_LABNO             As String               ' 접수번호
Public NO_WORKNO            As String               ' 일반Worksheet Unit
Public NO_WSUNIT            As String               ' 미생물Worksheet Unit
Public BN_TMP_ID            As String               ' 헌혈자 임시 환자 ID
Public BC2_KIT              As String               ' KIT (PHERESIS시 사용)
Public BC2_DONOR_REASON     As String               ' 헌혈자검사의뢰 판정사유
Public BC2_PHERESIS_COMPO   As String               ' 성분헌혈 혈액제제(01;혈장,02;혈소판,03;백혈구)cdval1
Public BC2_UNACCEPT_REASON  As String               ' 부적격 사유(01;ALT,02:B형간염.....등)CDVAL1코드
Public BC2_JOB              As String               ' 직업
Public BC2_TMP_ID           As String               ' 검사의뢰용 임시환자 ID
Public BC2_MATERIAL         As String               ' 추가재료내역
Public BC2_DS_TEST          As String               ' donor screening 검사항목
Public BC2_ASK              As String               ' 문진내역코드
Public BC2_RELATION         As String               ' 환자와의 관계


Public BC2_RC_COMPO As String   '혈액원에서 사용하는 제재
Public BC2_RC_ABO As String    '혈액원에서 사용하는 ABO TYPE
Public BC2_RC_VOL As String      '혈액원에서 사용하는 제재중 용량을 구한다.
Public BC2_XM_COMPO As String

'================================
'      바코드 관련정보
'================================
Public BC2_BLOOD_BAR        As String               '바코드 양식 '호환성 깨먹을때 삭제할것.... FRM301_FILE 폼 수정할 것

'상태별 표시 이름
Public STS_NM_ORDER As String '처방
Public STS_NM_COLLECT As String '채혈
Public STS_NM_ACCESS As String '접수
Public STS_NM_INPROGRESS As String '진행중
Public STS_NM_REQUEST As String '요청
Public STS_NM_DONE As String '완료
Public STS_NM_END As String '종결

'-----------------------
Private Sub Class_Initialize()

    C_WORKAREA = "B"
    
    BC2_KEEP_HOUR = "B001"              ' xm검체보관일수(시간)
    BC2_XM_STEP = "B002"                ' xm검사 step
    BC2_DS_TEST = "B003"                ' donor screening 검사항목
    BC2_COMPO = "B004"                  ' 혈액제제
    BC2_BRANCH = "B005"                 ' 자병원코드
    BC2_LOCAL = "B006"                  ' local 병원
    BC2_BLD_SRC = "B007"                ' 혈액원
    BC2_ASK = "B008"                    ' 문진내역코드
    BC2_RELATION = "B009"               ' 환자와의 관계
    BC2_CompoValue = "B032"             ' 혈액제제별 가격설정
    BC2_EXP_RESON = "B010"              ' 폐기사유
    BC2_ACTIVE_BLD_SRC = "B012"         ' 거래 혈액원
    BC2_ACCNO_CRITERION = "B013"        ' 접수번호 생성기준
    
    
    BC2_Trans_Reason = "B015"           ' 수혈사유
    BC2_JOB = "B016"                    ' 직업
    BC2_TMP_ID = "B017"                 ' 검사의뢰용 임시환자 ID
    BC2_MATERIAL = "B018"               ' 추가재료내역
    BC2_REACTION_TEST = "B019"          ' 수혈 부작용시 검사항목
    BC2_CANCELRSN = "B020"              ' 접수취소사유
    BC2_ABO_TEST = "B021"               ' ABO결과항목
    BC2_REACTION_TEST2 = "B022"         ' 수혈전 검사결과 일수
    BC2_RETURN_HOUR = "B023"            ' 반환 가능 시간(분)
    BC2_COMPO_GROUP = "B024"            ' 혈액제제 통계 그룹
    BC2_KIT = "B025"                    ' KIT (PHERESIS시 사용)
    BC2_DONOR_REASON = "B026"           ' 헌혈검사의뢰 판정사유
    BC2_PHERESIS_COMPO = "B027"         ' 성분헌혈 혈액제제(출력용)
    BC2_UNACCEPT_REASON = "B028"        ' 부적격사유(출력용)
    BC2_ABO_RH = "B029"                 ' ABO & RH 코드(혈액출고시 사용)
    BC2_BLOOD_BAR = "B030"              ' 혈액 바코드 양식
    BC2_REACTION_REASON = "B031"        ' 수혈 부작용 리스트
    
'검사결과코드
    BC2_RST_ABO = "B111"
    BC2_RST_RH = "B112"
    BC2_RST_ABOSUB = "B113"
    BC2_RST_RHSUB = "B114"
    BC2_RST_DU = "B115"

'마감일자관리
    BC2_File_Server = "B202"            ' 파일서버위치
    BC2_SPC_DEFAULT = "3"               ' 검체번호 앞에 붙는 문자
    BC2_BARCODE = "000000000"           ' 바코드출력시 붙는 문자
    BC2_ADDSPC = "추가채혈"             ' 추가채혈
    BC2_ORDVAL = "A5"                   ' 수술수유
    BC2_REASONVAL = "B1"                ' 수혈
'헌혈자 검사의뢰에서 사용되는 Cdindex
    CD1_MULTISPC = "C102"               ' 복수검체
    CD2_SPECIMEN = "C215"               ' 검체
    CD2_WORKAREA = "C213"               ' Work Area
    CD2_SGROUP = "C217"                 ' 검체군

    BN_TMP_ID = "C001"                  ' 검사의뢰용 임시 환자 ID
    BN_SPC_NO = "B002"                  ' 검체 번호
    BN_ACC_NO = "B004"                  ' 접수번호
    BN_DONOR_ID = "B006"                ' 헌혈자 ID

'헌혈자 검사의뢰에서 사용되는 Noindex
    NO_SPECIMEN = "L001"                ' 검체번호
    NO_LABNO = "L002"                   ' 접수번호
    NO_WORKNO = "L003"                  ' 일반Worksheet Unit
    NO_WSUNIT = "L009"                  ' 미생물Worksheet Unit

    BC2_CENTER = "C230"
    
    BC2_RC_COMPO = "B301"    '혈액원에서 사용하는 제재
    BC2_RC_ABO = "B302"  '혈액원에서 사용하는 ABO TYPE
    BC2_RC_VOL = "B303"  '혈액원에서 사용하는 제재 중 용량을 구함
    BC2_XM_COMPO = "B032" 'XM필수 입력 제제
    
'FORMAT 함수 변수
    PRESENTDATE_FORMAT = "YYYYMMDD"
    PRESENTTIME_FORMAT = "HHMMSS"
    
    STS_NM_ORDER = "처방" '처방
    STS_NM_COLLECT = "채혈" '채혈
    STS_NM_ACCESS = "접수" '접수
    STS_NM_INPROGRESS = "진행중" '진행중
    STS_NM_REQUEST = "요청" '요청
    STS_NM_DONE = "완료" '완결
    STS_NM_END = "종결" '종결
End Sub

