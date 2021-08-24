Attribute VB_Name = "modCmtConstants"
Option Explicit

Global Const HospitalNm = "가천의과대학 부속 길병원"

'Global declare the data class
Global Const DatabaseName$ = "Lab"
Global Const Connect$ = "Lab/Lab"
Global Const ConnectString = "dsn=sybaseODBC;uid=hisbase;pwd=hispass;"
'
Global SB_ServerNm As String    '-- 서버명
Global SB_DatabaseNm As String  '-- 데이타베이스명
Global SB_LoginId As String     '-- 로긴아뒤
Global SB_Password As String    '-- 패스워드

Global SB_ConnStatus As Integer

'병원구분===============>>>> 추후검토 요.
Global Const HosptGb = "10"

Global Const CS_DateMask = "0###-##-##"
Global Const CS_DateLMask = "0###/##/##"
Global Const CS_DateSMask = "0#-##-##"
Global Const CS_TimeSMask = "0#:##"
Global Const CS_TimeLMask = "0#:##:##"
Global Const CS_BlankMask = "____/__/__"
Global Const CS_DateSFormat = "YY-MM-DD"
Global Const CS_DateFormat = "YYYY-MM-DD"
Global Const CS_TimeSFormat = "HH:MM"
Global Const CS_TimeLFormat = "HH:MM:SS"
Global Const CS_DateDbFormat = "YYYYMMDD"
Global Const CS_TimeDbFormat = "HHMMSS"

'****************  Sybase DB 에서 System Data & Time 구하는 함수  ****************'
Global Const CS_SybaseDate = "convert(char(8),getdate(),112)"
Global Const CS_SybaseTime = "substring(convert(char(8),getdate(),108),1,2)+" & _
                             "substring(convert(char(8),getdate(),108),4,2)+" & _
                             "substring(convert(char(8),getdate(),108),7,2)"
'***********************************************************************************'

'Color
Global Const CR_LIGHT_BLUE = &HB9602F
Global Const CR_LIGHT_BLUE1 = &HDF6A3E
Global Const CR_LIGHT_RED = &H7477EF
Global Const CR_LIGHT_YELLOW = &HC0FFFF
Global Const CR_GREY = &HC0C0C0
Global Const CR_BROWN = &H404080


' 테이블명 상수 ( LIS Tables )
Global Const T_HIS001 = "h1ptntinfo"   '환자기본마스터
Global Const T_HIS002 = "h1admin"      '환자기본마스터
Global Const T_HIS003 = "hzdept"       '부서마스터
Global Const T_HIS004 = "hzdept"       '병동마스터
Global Const T_HIS005 = "HIS005"       '병상마스터
Global Const T_HIS007 = "hzempl"       '의사마스터
Global Const T_HIS008 = "h2diag"       '상병마스터
Global Const T_HIS009 = "h1actmat"     '수가마스터

Global Const T_LAB001 = "h7lab001"     '검사항목마스터
Global Const T_LAB002 = "h7lab002"     '검사동의어마스터
Global Const T_LAB003 = "h7lab003"     '검사별장비마스터
Global Const T_LAB004 = "h7lab004"     '지정검체마스터
Global Const T_LAB005 = "h7lab005"     '기준치마스터
Global Const T_LAB006 = "h7lab006"     '장비마스터
Global Const T_LAB007 = "h7lab007"     '외부검사마스터
Global Const T_LAB008 = "h7lab008"     'Worksheet마스터
Global Const T_LAB009 = "h7lab009"     '공지사항내역
Global Const T_LAB011 = "h7lab011"     'QC마스터
Global Const T_LAB012 = "h7lab012"     'QC검사정보마스터
Global Const T_LAB013 = "h7lab013"     '미생물QC마스터
Global Const T_LAB014 = "h7lab014"     'QC컨트롤마스터
Global Const T_LAB015 = "h7lab015"     '직원마스터

Global Const T_LAB021 = "h7lab021"     'QC Control Master
Global Const T_LAB022 = "h7lab022"     'QC Item Master
Global Const T_LAB023 = "h7lab023"     'QC Master
Global Const T_LAB024 = "h7lab024"     'QC Item Information
Global Const T_LAB025 = "h7lab025"     'QC Schedule
Global Const T_LAB026 = "h7lab026"     'QC 결과내역
Global Const T_LAB027 = "h7lab027"     'QC 수정내역
Global Const T_LAB028 = "h7lab028"     'QC Text 내역

Global Const T_LAB031 = "h7lab031"     '공통코드마스터1
Global Const T_LAB032 = "h7lab032"     '공통코드마스터2
Global Const T_LAB033 = "h7lab033"     '공통코드마스터3
Global Const T_LAB034 = "h7lab034"     '공통코드마스터3
Global Const T_LAB035 = "h7lab035"     '템플릿마스터
Global Const T_LAB036 = "h7lab036"     '기타검사템플릿마스터
Global Const T_LAB099 = "h7lab099"     '번호부여마스터

Global Const T_LAB101 = "h7lab101"     '처방Header
Global Const T_LAB102 = "h7lab102"     '처방Body
Global Const T_LAB103 = "h7lab103"     'QC처방Body

Global Const T_LAB201 = "h7lab201"     '채혈접수내역
Global Const T_LAB202 = "h7lab202"
Global Const T_LAB203 = "h7lab203"     '연속검사내역
Global Const T_LAB204 = "h7lab204"     '일괄채혈내역
Global Const T_LAB205 = "h7lab205"     '외부의뢰내역

Global Const T_LAB301 = "h7lab301"     'Worksheet내역
Global Const T_LAB302 = "h7lab302"     '일반결과내역
Global Const T_LAB303 = "h7lab303"     '일반텍스트결과내역
Global Const T_LAB304 = "h7lab304"     'FootNote내역
Global Const T_LAB305 = "h7lab305"     'Supplemental내역
Global Const T_LAB306 = "h7lab306"     '자동화장비 전송내역
Global Const T_LAB307 = "h7lab307"     'QC결과내역
Global Const T_LAB308 = "h7lab308"     '일반결과수정내역

Global Const T_LAB350 = "h7lab350"     '기타검사설정내역
Global Const T_LAB351 = "h7lab351"     '기타검사결과내역
Global Const T_LAB352 = "h7lab352"     '기타검사Numeric결과
Global Const T_LAB353 = "h7lab353"     '기타검사Text결과
Global Const T_LAB354 = "h7lab354"     '기타검사수정내역

Global Const T_LAB401 = "h7lab401"     '미생물Worksheet내역
Global Const T_LAB402 = "h7lab402"     '미생물Worksheet내역
Global Const T_LAB403 = "h7lab403"     '미생물Growth Reading내역
Global Const T_LAB404 = "h7lab404"     '미생물결과내역
Global Const T_LAB405 = "h7lab405"     '미생물감수성결과내역
Global Const T_LAB406 = "h7lab406"     '미생물QC결과내역
Global Const T_LAB407 = "h7lab407"     '미생물수정내역

'** 종합검증/판독 보고 관련 테이블 **'
Global Const T_LAB501 = "h7lab501"     '입원환자내역
Global Const T_LAB502 = "h7lab502"     '보고서내역
Global Const T_LAB503 = "h7lab503"     '결과내역
Global Const T_LAB504 = "h7lab504"     '수정내역
Global Const T_LAB505 = "h7lab505"     '설정내역
Global Const T_LAB506 = "h7lab506"     'Template

Global Const T_LAB999 = "h7lab999"     '구 system 결과내역

' 공통코드1 (LAB031) Index 상수
Global Const CD1_Index = "C100"
Global Const CD1_Panel = "C101"         ' Panel처방 Item
Global Const CD1_MultiSpc = "C102"      ' 복수검체
Global Const CD1_Detail = "C103"        ' Detail Items
Global Const CD1_KeyMap = "C104"        ' Keyboard mapping
Global Const CD1_AttrItem = "C105"      ' 속성 세부 Item
Global Const CD1_SpcMedia = "C106"      ' 검체군 - 배지
Global Const CD1_MediaBio = "C107"      ' 배지 - Bio Chemical Item
Global Const CD1_MicroAnti = "C108"     ' 균종 - 항생제
Global Const CD1_Machine = "C109"       ' 장비 - Item
Global Const CD1_ItemResult = "C110"    ' Item - 결과코드
Global Const CD1_WAResult = "C111"      ' WorkArea - 결과코드
Global Const CD1_QcControl = "C112"     ' QC Control
Global Const CD1_MBatchRst = "C113"     ' 미생물 웤싵별 - 배치 결과 코드
Global Const CD1_RelTest = "C114"       ' 관련검사코드
Global Const CD1_ColListTm = "C115"     ' 건물별 채혈리스트 출력시간
Global Const CD1_CumItem = "C116"       ' 누적결과조회 Item

' 공통코드2 (LAB032) Index 상수
Global Const CD2_Index = "C200"
Global Const CD2_DrGrade = "C201"       ' 의사Grade
Global Const CD2_BedGrade = "C202"      ' 병상등급
Global Const CD2_BedStatus = "C203"     ' 병상상태
Global Const CD2_DeptDiv = "C204"       ' 과분류
Global Const CD2_HighItem = "C205"      ' 다빈도처방
Global Const CD2_PocItem = "C206"       ' Point of Care
Global Const CD2_Bypass = "C207"        ' Bypass
Global Const CD2_RoundTime = "C208"     ' Round채혈 시간대
Global Const CD2_ColTeam = "C209"       ' 채혈팀
Global Const CD2_OutLab = "C210"        ' 외부의뢰처
Global Const CD2_RefLab = "C211"        ' Referral Lab
Global Const CD2_Vander = "C212"        ' Vander 코드
Global Const CD2_WorkArea = "C213"      ' Work Area
Global Const CD2_Section = "C214"       ' Section
Global Const CD2_Specimen = "C215"      ' 검체
Global Const CD2_VerifyFg = "C216"      ' Auto Verify On/Off
Global Const CD2_SGroup = "C217"        ' 검체군
Global Const CD2_Media = "C218"         ' 배지
Global Const CD2_Microbe = "C219"       ' 균
Global Const CD2_Species = "C220"       ' 균종
Global Const CD2_AntiBiotic = "C221"    ' 항생제
Global Const CD2_BioChemical = "C222"   ' 생화학적 동정검사
Global Const CD2_Volume = "C223"        ' 정도코드
Global Const CD2_Infect = "C224"        ' 법정전염병
Global Const CD2_QCOrderTime = "C225"   ' QC자동처방 시간대
Global Const CD2_BedDiv = "C226"        ' 병동분류
Global Const CD2_NoGrowth = "C227"      ' 미생물 Nogrowth Code
Global Const CD2_WorkSheetName = "C228" ' 워크쉬트 이름
Global Const CD2_StoreCd = "C229"       ' 보관구분
Global Const CD2_Buildings = "C230"     ' 건물코드
Global Const CD2_MWSKinds = "C231"      ' 미생물 웤싵 종류
Global Const CD2_FileServer = "C232"    ' File Server Location
Global Const CD2_StaticItem = "C233"    ' 월간 통계 항목
Global Const CD2_StaticGroup = "C234"   ' 월간 통계 Workarea
Global Const CD2_PrinterId = "C235"     ' Printer ID
Global Const CD2_StartDate = "C236"     ' 과별 검색기간 설정
Global Const CD2_PtDiv = "C237"         ' 환자구분


' 공통코드3 (LAB033) Index 상수
Global Const CD3_Index = "C300"
Global Const CD3_ScrLock = "C301"       ' Screen Lock Interval
Global Const CD3_PrgOnOff = "C302"      ' Program On/Off
Global Const CD3_FnctOnOff = "C303"     ' Fuction On/Off
Global Const CD3_InfectCond = "C304"    ' 원내감염 기준
Global Const CD3_BarFormat = "C305"     ' Barcode Label Format
Global Const CD3_BarTime = "C306"       ' 연속검사 Barcode Label 출력시점
Global Const CD3_WSPrtTime = "C307"     ' 기타검사 Worksheet 출력시점
Global Const CD3_Hospital = "C308"      ' 병원이름, 주소, 검사실이름
Global Const CD3_CumulTime = "C309"     ' 누적결과 출력시점
Global Const CD3_LabelTime = "C310"     ' 배지 Label 출력시점
Global Const CD3_TempUnit = "C311"      ' 냉장고 온도 단위
Global Const CD3_DateFormat = "C312"    ' 날짜 Format
Global Const CD3_TimeFormat = "C313"    ' 시간 Format

' Template (LAB034) Index 상수
Global Const CD4_Index = "C400"
Global Const CD4_Morphology = "C401"    ' 균 성상
Global Const CD4_UncolReason = "C402"   ' 미채혈 사유
Global Const CD4_Remark = "C403"        ' Remark
Global Const CD4_FootNote = "C404"      ' Foot Note
Global Const CD4_WarnInfect = "C405"    ' Warning/Infection
Global Const CD4_TextResult = "C406"    ' Text 결과
Global Const CD4_SPTextResult = "C407"  ' 기타검사 Text 결과
Global Const CD4_DCReason = "C408"      ' 처방취소 사유
Global Const CD4_CancelReason = "C409"  ' 접수취소 사유
Global Const CD4_ModifyReason = "C410"  ' 결과수정 사유
Global Const CD4_QCRejReason = "C411"   ' QC Reject 사유
Global Const CD4_TempReason = "C412"    ' 온도계 Reject 사유
Global Const CD4_ClinicalNotice = "C413" ' Clinical Notice

'Help Context ID
Global Const HLP_LogOn = 1003
Global Const HLP_EmpMaster = 1004
Global Const HLP_Order = 1011
Global Const HLP_Round = 1012
Global Const HLP_NulCol = 1013
Global Const HLP_Access = 1014
Global Const HLP_SendPt = 1015
Global Const HLP_Referral = 1016
Global Const HLP_BarReprint = 1017
Global Const HLP_WSBuild = 1021
Global Const HLP_AccEntry = 1022
Global Const HLP_InstEntry = 1023
Global Const HLP_WSEntry = 1024
Global Const HLP_ItemEntry = 1025
Global Const HLP_Modify = 1026
Global Const HLP_RstView = 1041
Global Const HLP_AllRstView = 1042
Global Const HLP_ItemMaster = 1061
Global Const HLP_SpcMaster = 1062
Global Const HLP_RefMaster = 1063
Global Const HLP_WSMaster = 1064

'검사종류
Global Const TST_RouTest = "0"      ' 대부분 검사
Global Const TST_SpeTest = "1"      ' 특수 검사
Global Const TST_MicTest = "2"      ' 미생물 검사

' 번호부여정보
Global Const NO_Specimen = "01"     '검체번호
Global Const NO_LabNo = "02"        '접수번호
Global Const NO_WorkNo = "03"       '일반Worksheet Unit
Global Const NO_WSUnit = "09"       '미생물Worksheet Unit

' BussDiv
Global Const CS_BussOut = "1"       '외래환자
Global Const CS_BussIn = "2"        '입원환자
Global Const CS_BussEr = "3"        '응급환자

' Panel Flag
Global Const PN_Group = "G"         'Group Item
Global Const PN_Detail = "D"        'Detail Item
Global Const PN_Normal = ""         '일반 Item

' FootNote 유무 (FootNoteFg in lab201)
Global Const RST_FootNote = "Y"

' Status Code
Global Const STS_Order = "0"        '처방
Global Const STS_HaveSpc = "1"      '채혈
Global Const STS_Access = "2"       '접수
Global Const STS_Worksheet = "3"    'In-Process
Global Const STS_MidRst = "4"       'Partial Verify / 중간결과
Global Const STS_FinRst = "5"       '확인 / 최종확인
Global Const STS_Modify = "6"       '수정

' 미생물 Worksheet 작성 대상 Flag
Global Const MWS_Ready = "1"        'Worksheet 작성
Global Const MWS_Holding = "2"      'Worksheet build 제외
Global Const MWS_Growth = "3"       'Growth 판정 - 비대상
Global Const MWS_Final = "4"        '최종결과 입력 완료 - 비대상

' 미생물 Worksheet 결과 구분 Flag
Global Const MRT_GSen = "S"
Global Const MRT_MSen = "C"
Global Const MRT_Stain = "G"
Global Const MRT_AFC = "M"
Global Const MRT_AFS = "N"
Global Const MRT_Both = "B"

Global Const MNM_GSen = "일반 감수성"        ' 감수성 입력화면용 표현
Global Const MNM_MSen = "MIC 감수성"
Global Const MNM_AFC = "AFB/Fungus Culture"
Global Const MNM_AFS = "AFB/Fungus Stain"

Global Const MCD_GSen = "GS"                 ' 검사에 따른 균종별 항생제 분류 구분자
Global Const MCD_MSen = "MS"

' 미생물 감수성 결과 존재 유무 (SenFg in lab404)
Global Const MRT_SenRst = "Y"
Global Const MRT_SenRstCd = "RISPN-"

' 기타검사 세부결과 존재 유무
Global Const ERT_ValRst = "Y"  '(ValFg in lab351)
Global Const ERT_TxtRst = "Y"  '(TxtFg in lab351)

' 기타검사 Worksheet Flag
Global Const EWS_OK = "1"
Global Const EWS_NO = "0"

'Other Error Codes
Global Const CONNECT_SUCCESS = 0
Global Const CONNECT_ERROR = 1

'Open Recordset 의 parameter
Global Const adOpenForwardOnly = 0
Global Const adOpenKeyset = 1
Global Const adOpenDynamic = 2
Global Const adOpenStatic = 3

