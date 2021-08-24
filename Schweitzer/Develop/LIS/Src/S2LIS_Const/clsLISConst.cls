VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISConst"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-----------------------------------------------------------------------------'
'   파일명 : clsLISConst
'   작성자 :
'   내  용 : LIS 공통변수
'   작성일 :
'   버  전 :
'       1. 5.0.5: 이상대(2005-01-03)
'          - 혈액학부 Workarea(CBC_WorkArea) 변수추가
'          - WBC 검사코드설정(LC3_WBCCode) 변수추가
'          - NRBC 검사코드설정(LC3_NRBCCode) 변수추가
'-----------------------------------------------------------------------------'

Option Explicit

'공통상수
'--------
Public LIS_ORDDIV           As String                         '임상병리 처방구분
Public APS_ORDDIV           As String                         '진단병리 처방구분
Public BBS_ORDDIV           As String                         '혈액은행 처방구분

Public POC_ORDDIV           As String                         'POC 구분
Public CMT_ORDDIV           As String                         'CMT 구분

'Public HospitalNm           As String
Public CentralLab           As String
Public CentralLabNm         As String
Public EmergencyNo          As Long
Public EmergencyLab         As String
Public EmergencyLabNm       As String
Public AneLab               As String
Public AneLabNm             As String
Public WomLab               As String
Public WomLabNm             As String
Public HrtLab               As String
Public HrtLabNm             As String
Public CS_BarFormat         As String
Public CS_AllCaption        As String
Public CS_DefaultBuilding   As String
Public APS_BUILDCD          As String
Public APS_BUILDNM          As String

Public LABName              As String
Public APSName              As String
Public BBSName              As String

Public APSBarNm             As String
Public BBSBarNm             As String

'공통코드 마스터1(T_LAB031)
'---------------------------
Public LC2_Index        As String
Public LC2_Panel        As String           ' Panel처방 Item
Public LC2_MultiSpc     As String           ' 복수검체
Public LC2_Detail       As String           ' Detail Items
Public LC2_AttrItem     As String           ' 속성 세부 Item
Public LC2_SpcMedia     As String           ' 검체군 - 배지
Public LC2_MediaBio     As String           ' 배지 - Bio Chemical Item
Public LC2_MicroAnti    As String           ' 균종 - 항생제
Public LC2_Machine      As String           ' 장비 - Item
Public LC2_ItemResult   As String           ' Item - 결과코드
Public LC2_WAResult     As String           ' WorkArea - 결과코드
Public LC2_QcControl    As String           ' QC Control
Public LC2_MBatchRst    As String           ' 미생물 웤싵별 - 배치 결과 코드
Public LC2_RelTest      As String           ' 관련검사코드
Public LC2_ColListTm    As String           ' 건물별 채혈리스트 출력시간
Public LC2_CumItem      As String           ' 누적결과조회 Item
Public LC2_MicQcRst     As String           ' 미생물QC 결과코드
Public LC2_MicTemp      As String           ' 미생물QC Template
Public LC2_SpeTemp      As String           ' 기타검사 Template
Public LC2_SpeAddTemp   As String           ' 기타검사 세부 Template
Public LC2_Calculation  As String           ' 기타검사 세부 Template
Public LC2_KeyMap       As String

Public LC2_OGYITEM      As String           ' 산부인과 검사항목코드
Public LC2_TempletTest  As String           ' 특수검사 템플릿 검사항목 설정.
Public LC2_TempletText1 As String           ' 특수검사 템플릿 text1
Public LC2_TempletText2 As String           ' 특수검사 템플릿 text1
Public LC2_TempletText3 As String           ' 특수검사 템플릿 text1
Public LC2_SelReview    As String           ' 항목별 Review
Public LC2_TelePhone    As String           ' TelePhone정보

'** 추가 전문의 확인 검사항목 등록을 위해... By M.G.Choi 2006.09.01
Public LC2_DoctTest     As String           ' 전문의 확인 검사항목 등록
' 예수병원 추가 By M.G.Choi 2007.07.02 =================================
Public LC2_Warning        As String      'Warning Message
'=======================================================================



'공통코드 마스터2(T_LAB032)
'---------------------------
Public LC3_Index            As String
Public LC3_DrGrade          As String           ' 의사Grade
Public LC3_BedGrade         As String           ' 병상등급
Public LC3_BedStatus        As String           ' 병상상태
Public LC3_DeptDiv          As String           ' 과분류
Public LC3_HighItem         As String           ' 다빈도처방
Public LC3_ByPass           As String           ' ByPass항목코드
Public LC3_PocItem          As String           ' Point of Care

Public LC3_RoundTime        As String           ' Round채혈 시간대
Public LC3_ColTeam          As String           ' 채혈팀
Public LC3_OutLab           As String           ' 외부의뢰처
Public LC3_RefLab           As String           ' Referral Lab
Public LC3_Vander           As String           ' Vander 코드
Public LC3_WorkArea         As String           ' Work Area
Public LC3_Section          As String           ' Section
Public LC3_Specimen         As String           ' 검체
Public LC3_VerifyFg         As String           ' Auto Verify On/Off
Public LC3_SGroup           As String           ' 검체군
Public LC3_Media            As String           ' 배지
Public LC3_Microbe          As String           ' 균
Public LC3_Species          As String           ' 균종
Public LC3_AntiBiotic       As String           ' 항생제
Public LC3_BioChemical      As String           ' 생화학적 동정검사
Public LC3_Volume           As String           ' 정도코드
Public LC3_Infect           As String           ' 법정전염병
Public LC3_QCOrderTime      As String           ' QC자동처방 시간대
Public LC3_BedDiv           As String           ' 병동분류
Public LC3_NoGrowth         As String           ' 미생물 Nogrowth Code
Public LC3_WorkSheetName    As String           ' 워크쉬트 이름
Public LC3_StoreCd          As String           ' 보관구분
Public LC3_Buildings        As String           ' 건물코드
Public LC3_MWSKinds         As String           ' 미생물 웤싵 종류
Public LC3_FileServer       As String           ' File Server Location
Public LC3_StaticItem       As String           ' 월간 통계 항목
Public LC3_StaticGroup      As String           ' 월간 통계 Workarea
Public LC3_PrinterId        As String           ' Printer ID
Public LC3_StartDate        As String           ' 과별 검색기간 설정
Public LC3_PtDiv            As String           ' 환자구분
Public LC3_PartCd           As String           ' 분과코드
Public LC3_HospCd           As String           ' 의뢰지병원
Public LC3_Method           As String           ' 외부정도관리 방법코드
Public LC3_BldQcRst         As String           ' 혈액은행QC 결과코드
Public LC3_ICUTestCd        As String           ' ICU 검사항목
Public LC3_POCTestCd        As String           ' POC 검사항목
Public LC3_POCTimes         As String           ' POC 입력시간대
Public LC3_ElectronicSign   As String           ' 임상병리 전자서명관리
Public LC3_ReportTesctCd    As String           ' 회진용레포트 검사코드
Public LC3_BatchColDept     As String           ' 일괄채혈 가능 부서코드
Public LC3_MachHistory      As String           ' 장비상태관리내역
Public LC3_DiffCount        As String           ' bone marrow diff count
Public LC3_WBCDiffCode      As String           ' WBC Diff 검사코드
Public LC3_DiffKeyMap       As String           ' Diff Count Key Mapping
Public LC3_TUBERCLE         As String           ' 결핵균검사 약제코드
Public LC3_ItemWorkLoad     As String           ' Item별 WorkLoad
Public LC3_EtcWorkLoad      As String           ' Etc Workload
Public LC3_HosFrmUsing      As String           ' 병원별 화면 로딩
Public LC3_WorkTime         As String
Public LC3_WBCCode          As String           ' WBC 검사코드 설정
Public LC3_NRBCCode         As String           ' NRBC 검사코드 설정

' 예수병원 추가 By M.G.Choi 2004.11.01 =================================
Public LC3_INFECTION        As String           ' 감염관리 의뢰검체
Public LC3_INFECTIONTEST    As String           ' 감염관리 균종관리

Public LC3_RESULTREADTEST   As String           ' 감염관리 균종관리
'=======================================================================

'템플릿 마스터(T_COM004)
'---------------------------
Public LC4_Index            As String
Public LC4_Morphology       As String           ' 균 성상
Public LC4_UncolReason      As String           ' 미채혈 사유
Public LC4_Remark           As String           ' Remark
Public LC4_FootNote         As String           ' Foot Note
Public LC4_WarnInfect       As String           ' Warning/Infection
Public LC4_TextResult       As String           ' Text 결과
Public LC4_SPTextResult     As String           ' 기타검사 Text 결과
Public LC4_DCReason         As String           ' 처방취소 사유
Public LC4_CancelReason     As String           ' 접수취소 사유
Public LC4_ModifyReason     As String           ' 결과수정 사유
Public LC4_QCRejReason      As String           ' QC Reject 사유
Public LC4_TempReason       As String           ' 온도계 Reject 사유
Public LC4_ClinicalNotice   As String           ' Clinical Notice
Public LC4_Peripheral       As String           ' Peripheral
Public LC4_BMComment        As String           ' bone marrow comment
Public LC4_Hematologic      As String           ' hematologic

Public LC4_TempletCommet    As String           ' 템플릿 commet
Public LC4_TempletDignosis  As String           ' 템플릿 진단
Public LC4_Calibration      As String           ' QC Calibration 소견
Public LC4_TestItemComment  As String           ' 검사항목별 Comment등록
Public LC4_AccessComment    As String           ' 외래접수시 Comment 템플릿

' 예수병원 추가 By M.G.Choi 2004.11.01 =================================
Public LC4_Infection        As String    '감염관리 검사방법
'=======================================================================

'프로젝트 파라미터 마스터(T_COM005)
'---------------------------
Public LC5_Index As String
Public LC5_ScrLock As String        ' Screen Lock Interval
Public LC5_PrgOnOff As String       ' Program On/Off
Public LC5_FnctOnOff As String      ' Fuction On/Off
Public LC5_InfectCond As String     ' 원내감염 기준
Public LC5_BarFormat As String      ' Barcode Label Format
Public LC5_BarTime As String        ' 연속검사 Barcode Label 출력시점
Public LC5_WSPrtTime As String      ' 기타검사 Worksheet 출력시점
Public LC5_Hospital As String       ' 병원이름, 주소, 검사실이름
Public LC5_CumulTime As String      ' 누적결과 출력시점
Public LC5_LabelTime As String      ' 배지 Label 출력시점
Public LC5_TempUnit As String       ' 냉장고 온도 단위
Public LC5_DateFormat As String     ' 날짜 Format
Public LC5_TimeFormat As String     ' 시간 Format

'번호부여 마스터(T_COM099)
'-------------------------
Public COM99_LIS_Specimen As String     '검체번호
Public COM99_LIS_LabNo As String        '접수번호
Public COM99_LIS_SeqNo As String       '접수순번(Workarea 별)
Public COM99_LIS_WorkNo As String       '일반Worksheet Unit
Public COM99_LIS_WSUnit As String       '미생물Worksheet Unit
Public COM99_LIS_Report As String       '임상병리 결과지 출력일자 관리

Public LIS_BarDiv  As String            '바코드구분(1:임상,2:해부,3:혈액)
Public LIS_BarFormat As String          '바코드포맷

Public LIS_LabDayFormat As String             '접수번호부여기준 - 일단위
Public LIS_LabMonthFormat As String           '접수번호부여기준 - 월단위
Public LIS_LabYearFormat As String            '접수번호부여기준 - 년단위


'결과등록 클래스의 상수
'-------------------------
Public NoIndex_WSBuild  As Long          'WorkSheet Sequence
'
Public Splt_Delimeter   As String
'
'Constants for Result- Data Entry Type
Public RESULT_BY_DEFAULT   As Long
Public RESULT_BY_ACCESSION As Long
Public RESULT_BY_WORKSHEET As Long
Public RESULT_BY_EQUIPMENT As Long
Public RESULT_BY_ITEM      As Long
Public RESULT_BY_MODIFY    As Long
Public RESULT_BY_DIFFCOUNT As Long

'처방속성코드
'---------------------
Public LIS_ATTR_SPECIAL_INTEREST    As String           'Special Interest for Ex.
Public LIS_ATTR_CLINICAL_SUMMARY    As String           '환자병력
Public LIS_ATTR_OPERATION_FINDINGS  As String           '수술소견
Public LIS_ATTR_CLINICAL_DX_CD      As String           '임상진단코드(임상상병)
Public LIS_ATTR_OPERATION_CD        As String           '수술코드
Public LIS_ATTR_OPERATION_DX_CD     As String           '수술상병코드
Public LIS_ATTR_LMP                 As String           '최종월경일
Public LIS_ATTR_MESTRUATION_CYCLE   As String           '월경주기
Public LIS_ATTR_MENAPAUSE           As String           '폐경기나이

'기타 공용선언
'-------------
Public HLDIV_HIGH_CD        As String
Public HLDIV_HIGH_FG        As String
Public HLDIV_LOW_CD         As String
Public HLDIV_LOW_FG         As String

Public LIS_RST_REQUIRED     As String
Public LIS_RST_ALTERNATIVE  As String

Public LIS_Nothing          As String

'처방 STATUS
'-----------
Public STS_LIS_Order        As String   '처방
Public STS_LIS_HaveSpc      As String   '채혈
Public STS_LIS_Access       As String   '접수
Public STS_LIS_Worksheet    As String   'In-Process
Public STS_LIS_Reading      As String   '판독
Public STS_LIS_MidRst       As String   'Partial Verify / 중간결과
Public STS_LIS_FinRst       As String   '최종확인
Public STS_LIS_Partial      As String   '부분
Public STS_LIS_Verify       As String   '결과
Public STS_LIS_Modify       As String   '수정
Public STS_LIS_Cancel       As String   '채혈/취소

'Panel Flag
'-------------------------
Public PN_Group     As String           'Group Item
Public PN_Detail    As String           'Detail Item
Public PN_Normal    As String           '일반 Item

'외부의뢰(SEND OUT) STATUS
'-------------------------
Public STS_OUTLAB_ACCESSION As String   '접수
Public STS_OUTLAB_TRANSFER  As String   '전송
Public STS_OUTLAB_WITHDRAW  As String   '회송
Public STS_OUTLAB_RESULT    As String   '결과


'수탁검사(REFERRAL) STATUS
'-------------------------
Public STS_REFERRAL_ACCESSION               '접수
Public STS_REFERRAL_INPROCESS               '검사
Public STS_REFERRAL_RESULT                  '회송
Public STS_REFERRAL_RESULTTRANS             '결과


Public MIC_WorkArea As String               '미생물 Work Area
Public CBC_WorkArea As String               '혈액학부 Work Area

' 미생물 Worksheet 작성 대상 Flag
Public MWS_Ready        As String           'Worksheet 작성
Public MWS_Holding      As String           'Worksheet build 제외
Public MWS_Growth       As String           'Growth 판정 - 비대상
Public MWS_Final        As String           '최종결과 입력 완료 - 비대상
Public MWS_Selected     As String
Public MWS_Deselected   As String
Public MWS_ForCulture   As Long             '미생물 배양검사에 해당하는 Worksheet코드
Public MWS_ForSensi     As Long             '미생물 배양검사에 해당하는 Worksheet코드
Public MWS_ForStain     As Long             '미생물 Stain검사에 해당하는 Worksheet코드
Public MWS_ForAll       As Long             '미생물 모든검사에 해당하는 Worksheet코드

' 미생물 Worksheet 결과 구분 Flag
Public MRT_GenSen   As String               'General Sensitivity
Public MRT_MicSen   As String               'MIC Sensitivity
Public MRT_Stain    As String               'Gram Stain
Public MRT_AFC      As String               'AFB Culture
Public MRT_AFS      As String               'AFB Stain
Public MRT_AFN      As String               'AFB Sensitivity
Public MRT_FUC      As String               'Fungus Culture
Public MRT_FUS      As String               'Fungus Stain
Public MRT_Both     As String               '?
Public MRT_OTHER    As String               '?

Public MNM_GSen As String                   ' 감수성 입력화면용 표현
Public MNM_MSen As String
Public MNM_AFC  As String
Public MNM_AFS  As String

Public MCD_GSen As String                   ' 검사에 따른 균종별 항생제 분류 구분자
Public MCD_MSen As String

' 미생물 감수성 결과 존재 유무 (SenFg in lab404)
Public MRT_NoSenRst As String
Public MRT_SenRst   As String
Public MRT_SenRstCd As String

' 기타검사 세부결과 존재 유무
Public ERT_ValRst   As String               '(ValFg in lab351)
Public ERT_TxtRst   As String               '(TxtFg in lab351)

' 기타검사 Worksheet Flag
Public EWS_OK       As String
Public EWS_NO       As String

' 기타검사 결과지 출력 Work Area
Public ERT_REPORT   As String

' QC Level 코드
Public QC_LEVEL_HIGH    As String
Public QC_LEVEL_NORMAL  As String
Public QC_LEVEL_LOW     As String

'핵의학 WorkArea
Public RI_WORKAREA      As String

' Interface 오류 데이타 상수
Public CS_EqpError      As String

'DECLARE ENUM
'-------------
Public Enum enOrdFG                           '처방유형
    Order_OCS = 0
    Order_BackUp = 1
    Order_ItemAdd = 2
    Order_Reverse = 3
    Order_Change = 4
    Order_Repeat = 5
    Order_DC = 6
    Order_QC = 7
    Order_SpecialStudy = 8
End Enum

Public Enum enBussDiv                       '업무구분
    BussDiv_OutPatient = "1"
    BussDiv_InPatient = "2"
    BussDiv_Emergency = "3"
    BussDiv_HealthCaer = "4"
    BussDiv_Referral = "5"
    BussDiv_ICU = "9"
End Enum

Public Enum enSpOrdDiv                      '특수처방구분
    SpOrdDiv_None = 0
    SpOrdDiv_VIP = 1
    SpOrdDiv_Research = 2
    SpOrdDiv_Other = 3
End Enum

Public Enum enTrustFg                       '수탁구분
    TrustFg_None = 0
    TrustFg_Referral = 1
End Enum

Public Enum enStsCd                       '처방 STATUS
    StsCd_LIS_Order = 0                   '처방
    StsCd_LIS_Collection = 1              '채취
    StsCd_LIS_Accession = 2               '접수
    StsCd_LIS_InProcess = 3               '검사중
    StsCd_LIS_MidRst = 4                  '중간보고
    StsCd_LIS_FinRst = 5                  '결과확인(최종결과)
    StsCd_LIS_Modify = 6                  '결과수정
    StsCd_LIS_Cancel = 7                  '결과취소
End Enum

'*
Public Enum enTestDiv                       '검사구분
    TST_RouTest = "0"      ' 일반 검사
    TST_SpeTest = "1"      ' 특수 검사
    TST_MicTest = "2"      ' 미생물 검사
    TST_AboTest = "3"      ' 혈액형 검사
End Enum

Public Enum enLabDiv                       '접수번호부여기준
    LabDiv_ByDay = 1        ' 일단위
    LabDiv_ByMonth = 2      ' 월단위
    LabDiv_ByYear = 3       ' 년단위
    LabDiv_BySpc = 4        ' 검체군단위
End Enum

Public Enum enFlag
    Flag_Off = "0"                        '부
    Flag_On = "1"                         '여
End Enum


Public Enum enDoneFg                         '접수여부
    DoneFg_Order = 0
    DoneFg_Collection = 1
    DoneFg_Accession = 2
End Enum

'
'Public Enum FootnoteFg
'    FootNote_Exist = "Y"
'    FootNote_None = ""
'End Enum

Public Enum enLabNoY2kFlag
    LabNoY2kFlag_20C = "19"
    LabNoY2kFlag_21C = "20"
End Enum

Public Type Field
    Field As String
End Type

Public Type ORDB
    PtId As String
    OrdDt As String
    OrdNo As Long
    OrdSeq As Long
    OrdCd As String
    SpcCd As String
    RcvDt As String
    RCVTM As String
    StsCd As enStsCd
    StatFg As String
    DoneFg As enDoneFg
    WorkArea As String
    AccDt As String
    AccSeq As Long
    ReceptNo As Long
    Mesg As String
End Type

Public Type ACCESSION
    WorkArea As String
    AccDt As String
    AccSeq As Integer
    PtId As String
    OrdDt As String
    Sex As String
    Age As Integer
    BedInDt As String
    ReqDt As String
    ReqTm As String
    DeptCd As String
    OrdDoct As String
    MajDoct As String
    EntId As String
    OrdDiv As String
    OrdFG As enOrdFG
    SpOrdDiv As enSpOrdDiv
    DoneFg As enDoneFg
    WardID As String
    HosilId As String
    BedID As String
    RefHosCd As String
    TrustFg As enTrustFg
    TrustDt As String
    OrdNo As Long
    OrdTm As String
    EntDt As String
    EntTm As String
    OrgAccNo As String
    HosCd As String
    Seq As Integer
    RcvDt As String
    RCVTM As String
    SPCYY As String
    SPCNO As String
    StsCd As enStsCd
    StatFg As String
    FZFG As String
End Type

Public Enum enSaveCollection
    Save_Ward_Collection = 0                        '병동채혈
    Save_Nurse_Collection = 1                       '병동간호사채혈
    save_OutPatient_Collection = 2                  '외래간호사채혈
End Enum

'= 처방 Body ='
Public Type OrderBody
   OrdSeq As Integer    ' 처방 Seq
   OrdCd As String      ' 처방코드
   SpcCd As String      ' 검체코드
   StoreCd As String    ' 보관구분
   DcFg As String       ' DC 여부
   DcDt As String       ' DC 일자
   DcNo As Integer      ' DC 처방번호
   AttrCd As String     ' 속성여부
   ExamDt As String     ' 실제 실시일자
   ExamTm As String     ' 실제 실시시간
   ExamDoct As String   ' 실제 실시의사
   StsCd As String      ' Status
   StatFg As String     ' 응급여부
   InsDiv As String     ' 급여구분
   DoneFg As String
End Type

'= 처방 Header ='
Public Type OrderHeader
   OrdNo As Integer     ' 처방번호
   ReqDt As String      ' 실시요청일자
   ReqTm  As String     ' 실시요청시간
   OrdBody() As OrderBody
End Type


Public Type ResultTable
    TestCd As String     '/* 검사항목 코드   */
    RstVal As String     '/* 결과(Numeric)   */
    RstCd As String      '/* 결과코드(Alpha) */
    RstUnit As String    '/* Unit            */
    HLDiv As String      '/* High/Low(H:Hig,L:Low) */
    DPDiv As String      '/* Delta/Panic(D:Delta,P:Panic) */
    SpcCd As String      '/* 검체코드        */
    StatFg As String     '/* 응급여부('0':부,'1':여) */
    LastRst As String    '/* 최근결과        */
    LastVfyDt As String  '/* 최근결과확인일  */
    LastVfyTm As String  '/* 최근결과확인시간 */
    LastVfyId As Integer '/* 최근결과확인자  */
    VfyDt As String      '/* 결과확인일자    */
    VfyTm As String      '/* 결과확인시간    */
    VfyId As Integer     '/* 결과확인자      */
    AttrCd As String     '/* 속성코드   */
    MfyFg As String      '/* 수정여부('0':부,'1':여) */
    GrpFg As String      '/* 그래프 결과여부('0':부,'1':여) */
    ValFg As String      '/* 수치 결과여부('0':부,'1':여)  - 기타검사에서만 사용*/
    TxtFg As String      '/* TEXT 결과여부('0':부,'1':여) */
    RstType As String    '/* 결과유형(N,Alpha,비율,Free) */
    RstDiv As String     '/Alternative,Required 여부 */
    DetailFg As String   '/* 상세항목Group Seq */
    PtId As String       '/* 환자 ID */
    OrdDt As String      '/* 처방일 */
    OrdNo As Integer     '/* 처방번호 */
    OrdSeq As Integer    '/* 처방Seq */
    SenFg As String      '/* 감수성결과 여부 */
    AutoFg As String
    EqpCd As String
    
    TestDiv As String
End Type

Public Type OutLabTable
    WorkArea As String  '/* Work Area(접수번호용) */
    AccDt As String     '/* 접수일(접수번호용)    */
    AccSeq As Integer   '/* 접수순번(접수번호용)  */
    TestCd As String    '/* 검사항목        */
    PtId As String      '/* 환자ID  */
    SpcCd As String     '/* 검체코드  */
    OutLabCd As String  '/* 기관코드        */
    StsCd As String     '/* STATUS          */
    RcvDt As String     '/* 접수일자        */
    SendDt As String    '/* SEND Date       */
    ChargeDt As String  '/* 청구일자        */
    SendId As String    '/* 전송자 ID   */
End Type

Public Type tpMicWorkSheet
    SpcGrp As String
    SpcGrpNm As String
    WsCode As String
    WsName As String
    WsType As String
    WsRstType As String
End Type

Public Type tTmpValue
    sTpcd As String
    sTpnm As String
    swsfg As String           ' 1 : check,   0: uncheck
    sTpdata As String
    sValueCnt As Long
    sRstFields As String
    sSaveDecision As Integer  ' 1 : save ,   0 : Unsave
End Type

Private Sub Class_Initialize()

    '공통코드 마스터1(T_COM002)
    LC2_Index = "C100"
    LC2_Panel = "C101"         ' Panel처방 Item
    LC2_MultiSpc = "C102"      ' 복수검체
    LC2_Detail = "C103"        ' Detail Items
    LC2_AttrItem = "C105"      ' 속성 세부 Item
    LC2_SpcMedia = "C106"      ' 검체군 - 배지
    LC2_MediaBio = "C107"      ' 배지 - Bio Chemical Item
    LC2_MicroAnti = "C108"     ' 균종 - 항생제
    LC2_Machine = "C109"       ' 장비 - Item
    LC2_ItemResult = "C110"    ' Item - 결과코드
    LC2_WAResult = "C111"      ' WorkArea - 결과코드
    LC2_QcControl = "C112"     ' QC Control
    LC2_MBatchRst = "C113"     ' 미생물 웤싵별 - 배치 결과 코드
    LC2_RelTest = "C114"       ' 관련검사코드
    LC2_ColListTm = "C115"     ' 건물별 채혈리스트 출력시간
    LC2_CumItem = "C116"       ' 누적결과조회 Item
    LC2_MicQcRst = "C117"      ' 미생물QC 결과코드
    LC2_MicTemp = "C118"       ' 미생물QC Template
    LC2_SpeTemp = "C119"       ' 기타검사 결과 Template
    LC2_SpeAddTemp = "C120"    ' 기타검사 결과 세부 Template
    LC2_Calculation = "C121"   ' 계산공식 적용
    LC2_OGYITEM = "C122"       ' 산부인과 검사항목 코드
    LC2_TempletTest = "C124"   ' 템플릿 검사항목 설정.
    LC2_TempletText1 = "C125"
    LC2_TempletText2 = "C126"
    LC2_TempletText3 = "C127"
    LC2_SelReview = "C128"      ' SelReview 검사코드/검체명
    LC2_TelePhone = "C129"      ' 전화번호정보
    
    '** 추가 전문의 확인 검사항목 By M.G.Choi 2006.09.01
    LC2_DoctTest = "C150"
    ' 예수병원 추가 By M.G.Choi 2007.07.02 =================================
    LC2_Warning = "C151"        ' Warning Message
    '=======================================================================
    '---------------------------------------------------
    
    '공통코드 마스터2(T_COM003)
    LC3_Index = "C200"
    LC3_DrGrade = "C201"       ' 의사Grade
    LC3_BedGrade = "C202"      ' 병상등급
    LC3_BedStatus = "C203"     ' 병상상태
    LC3_DeptDiv = "C204"       ' 과분류
    LC3_HighItem = "C205"      ' 다빈도처방
    LC3_PocItem = "C206"       ' Point of Care
    LC3_ByPass = "C207"        ' Bypass
    LC3_RoundTime = "C208"     ' Round채혈 시간대
    LC3_ColTeam = "C209"       ' 채혈팀
    LC3_OutLab = "C210"        ' 외부의뢰처
    LC3_RefLab = "C211"        ' Referral Lab
    LC3_Vander = "C212"        ' Vander 코드
    LC3_WorkArea = "C213"      ' Work Area
    LC3_Section = "C214"       ' Section
    LC3_Specimen = "C215"      ' 검체
    LC3_VerifyFg = "C216"      ' Auto Verify On/Off
    LC3_SGroup = "C217"        ' 검체군
    LC3_Media = "C218"         ' 배지
    LC3_Microbe = "C219"       ' 균
    LC3_Species = "C220"       ' 균종
    LC3_AntiBiotic = "C221"    ' 항생제
    LC3_BioChemical = "C222"   ' 생화학적 동정검사
    LC3_Volume = "C223"        ' 정도코드
    LC3_Infect = "C224"        ' 법정전염병
    LC3_QCOrderTime = "C225"   ' QC자동처방 시간대
    LC3_BedDiv = "C226"        ' 병동분류
    LC3_NoGrowth = "C227"      ' 미생물 Nogrowth Code
    LC3_WorkSheetName = "C228" ' 워크쉬트 이름
    LC3_StoreCd = "C229"       ' 보관구분
    LC3_Buildings = "C230"     ' 건물코드
    LC3_MWSKinds = "C231"      ' 미생물 웤싵 종류
    LC3_FileServer = "C232"    ' File Server Location
    LC3_StaticItem = "C233"    ' 월간 통계 항목
    LC3_StaticGroup = "C234"   ' 월간 통계 Workarea
    LC3_PrinterId = "C235"     ' Printer ID
    LC3_StartDate = "C236"     ' 과별 검색기간 설정
    LC3_PtDiv = "C237"         ' 환자구분
    LC3_PartCd = "C238"        ' 분과코드
    LC3_HospCd = "C239"        ' 의뢰지병원
    LC3_Method = "C240"        ' 외부정도관리 방법코드
    LC3_BldQcRst = "C241"      ' 혈액은행QC 결과코드
    LC3_ICUTestCd = "C242"     ' ICU 항목코드
    LC3_POCTestCd = "C243"     ' POC 항목코드
    LC3_POCTimes = "C244"      ' POC 입력시간대
    LC3_ElectronicSign = "C245" '전자서명
    LC3_ReportTesctCd = "C246" ' 회진용레포트 검사코드
    LC3_WBCDiffCode = "C247"    'WBC Diff 검사코드
    LC3_BatchColDept = "C248"
    LC3_DiffCount = "C251"      ' bone marrow diff count
    LC3_MachHistory = "C252"    ' 장비상태관리내역
    LC3_TUBERCLE = "C253"       ' 결핵균검사 약제코드
    LC3_ItemWorkLoad = "C254"   ' Item WorkLoad
    LC3_EtcWorkLoad = "C255"    ' Etc WorkLoad
    LC3_DiffKeyMap = "C256"     ' Diff Key Mapping
    LC3_HosFrmUsing = "C261"    ' 임상병리 병원별 로딩화면 설정 혈액은행은 "C262" 'C257이 원래 값
    LC3_WorkTime = "C258"       ' 업무시간대 작성
    LC3_WBCCode = "C270"        ' WBC 검사코드 설정
    LC3_NRBCCode = "C271"       ' NRBC 검사코드 설정
    
    ' 예수병원 추가 By M.G.Choi 2004.11.01 =================================
    LC3_INFECTION = "C259"      ' 감염관리 의뢰검체
    LC3_INFECTIONTEST = "C260"  ' 감염관리 균종관리
    
    LC3_RESULTREADTEST = "C265" ' 종합판독보고 검사항목
    '=======================================================================
    
    

    '템플릿 마스터(T_COM004)
    LC4_Index = "C400"
    LC4_Morphology = "C401"         ' 균 성상
    LC4_UncolReason = "C402"        ' 미채혈 사유
    LC4_Remark = "C403"             ' Remark
    LC4_FootNote = "C404"           ' Foot Note
    LC4_WarnInfect = "C405"         ' Warning/Infection
    LC4_TextResult = "C406"         ' Text 결과
    LC4_SPTextResult = "C407"       ' 기타검사 Text 결과
    LC4_DCReason = "C408"           ' 처방취소 사유
    LC4_CancelReason = "C409"       ' 접수취소 사유
    LC4_ModifyReason = "C410"       ' 결과수정 사유
    LC4_QCRejReason = "C411"        ' QC Reject 사유
    LC4_TempReason = "C412"         ' 온도계 Reject 사유
    LC4_ClinicalNotice = "C413"     ' Clinical Notice
    LC4_Peripheral = "C420"         ' Peripheral
    LC4_BMComment = "C421"          ' bone marrow comment
    LC4_Hematologic = "C422"        ' hematologic
    LC4_TempletCommet = "C423"      ' 템플릿 commet
    LC4_TempletDignosis = "C424"    ' 템플릿 진단
    LC4_Calibration = "C425"        ' 장비 Calibration
    LC4_TestItemComment = "C426"    ' 검사항목별 Comment
    LC4_AccessComment = "C427"      ' 외래접수시 Comment 템플릿
    
    ' 예수병원 추가 By M.G.Choi 2004.11.01 =================================
    LC4_Infection = "C428"          '감염관리 검사방법
    '=======================================================================
    
    '프로젝트 파라미터 마스터(T_COM005)
    LC5_Index = "C500"
    LC5_ScrLock = "C501"       ' Screen Lock Interval
    LC5_PrgOnOff = "C502"      ' Program On/Off
    LC5_FnctOnOff = "C503"     ' Fuction On/Off
    LC5_InfectCond = "C504"    ' 원내감염 기준
    LC5_BarFormat = "C505"     ' Barcode Label Format
    LC5_BarTime = "C506"       ' 연속검사 Barcode Label 출력시점
    LC5_WSPrtTime = "C507"     ' 기타검사 Worksheet 출력시점
    LC5_Hospital = "C508"      ' 병원이름, 주소, 검사실이름
    LC5_CumulTime = "C509"     ' 누적결과 출력시점
    LC5_LabelTime = "C510"     ' 배지 Label 출력시점
    LC5_TempUnit = "C511"      ' 냉장고 온도 단위
    LC5_DateFormat = "C512"    ' 날짜 Format
    LC5_TimeFormat = "C513"    ' 시간 Format

    '공통상수
    LIS_ORDDIV = "L"
    APS_ORDDIV = "A"
    BBS_ORDDIV = "B"
    POC_ORDDIV = "P"
    CMT_ORDDIV = "C"
        'Constants for Result- Data Entry Type
    RESULT_BY_DEFAULT = 1
    RESULT_BY_ACCESSION = 2
    RESULT_BY_WORKSHEET = 3
    RESULT_BY_EQUIPMENT = 4
    RESULT_BY_ITEM = 5
    RESULT_BY_MODIFY = 6
    RESULT_BY_DIFFCOUNT = 7

    '진단병리 처방속성코드
    LIS_ATTR_SPECIAL_INTEREST = "A1"        'Special Interest for Ex.
    LIS_ATTR_CLINICAL_SUMMARY = "A2"        '환자병력
    LIS_ATTR_OPERATION_FINDINGS = "A4"      '수술소견
    LIS_ATTR_CLINICAL_DX_CD = "A3"          '임상진단코드(임상상병)
    LIS_ATTR_OPERATION_CD = "A5"            '수술코드
    LIS_ATTR_OPERATION_DX_CD = "A6"         '수술상병코드
    LIS_ATTR_LMP = "A7"                     '최종월경일
    LIS_ATTR_MESTRUATION_CYCLE = "A8"       '월경주기
    LIS_ATTR_MENAPAUSE = "A9"               '폐경기나이


    '기타공용선언
    STS_LIS_Order = "처방"
    STS_LIS_HaveSpc = "채취"
    STS_LIS_Access = "접수"
    STS_LIS_Worksheet = "검사중"
    STS_LIS_Reading = "판독"
    STS_LIS_MidRst = "중간"
    STS_LIS_FinRst = "최종"
    STS_LIS_Verify = "결과"
    STS_LIS_Partial = "부분"
    STS_LIS_Modify = "수정"
    STS_LIS_Cancel = "취소"
    '
    PN_Group = "G"         'Group Item
    PN_Detail = "D"        'Detail Item
    PN_Normal = ""         '일반 Item

    '외부의뢰(SEND OUT) STATUS
    STS_OUTLAB_ACCESSION = "1"              '접수
    STS_OUTLAB_TRANSFER = "2"               '전송
    STS_OUTLAB_WITHDRAW = "3"               '회송
    STS_OUTLAB_RESULT = "4"                 '결과

    '수탁검사(REFERRAL) STATUS
    STS_REFERRAL_ACCESSION = "1"            '접수
    STS_REFERRAL_INPROCESS = "2"            '검사
    STS_REFERRAL_RESULT = "3"               '회송
    STS_REFERRAL_RESULTTRANS = "4"          '결과

'    HospitalNm = "가천의과대학 부속 길병원"
    CentralLab = "10"           '-- 중앙검사실
    CentralLabNm = "중앙"       '-- 중앙검사실
    WomLab = "20"               '-- 여성클리닉
    WomLabNm = "여성클리닉"     '-- 여성클리닉
    HrtLab = "30"               '-- 심장센터
    HrtLabNm = "심장센터"       '-- 심장센터
    AneLab = "40"               '-- 안이센터
    AneLabNm = "안이센터"       '-- 안이센터
    EmergencyNo = 5             '-- 응급센터
    EmergencyLab = "50"         '-- 응급센터
    EmergencyLabNm = "응급센터"   '-- 응급센타

    CS_AllCaption = "(전체)"
    CS_DefaultBuilding = "10"
    APS_BUILDCD = "10"
    APS_BUILDNM = "본원"
    LABName = "임상"
    APSName = "해부"
    APSBarNm = "AP"
    BBSName = "혈액"
    BBSBarNm = "XM"
    LIS_BarDiv = "1"

    LIS_LabDayFormat = "YYYYMMDD"     '일단위
    LIS_LabMonthFormat = "YYYYMM"     '월단위
    LIS_LabYearFormat = "YYYY"        '년단위

    NoIndex_WSBuild = 10            'WorkSheet Sequence
    Splt_Delimeter = "$"

    HLDIV_HIGH_CD = "H"
    HLDIV_HIGH_FG = "▲"
    HLDIV_LOW_CD = "L"
    HLDIV_LOW_FG = "▼"

    LIS_RST_REQUIRED = "R"
    LIS_RST_ALTERNATIVE = "A"

    LIS_Nothing = "(없음)"

' 미생물 Worksheet 작성 대상 Flag
    MWS_Ready = "1"                     'Worksheet 작성
    MWS_Holding = "2"                   'Worksheet build 제외
    MWS_Growth = "3"                    'Growth 판정 - 비대상
    MWS_Final = "4"                     '최종결과 입력 완료 - 비대상

    MWS_Selected = "1"
    MWS_Deselected = "0"

    MWS_ForAll = 0
    MWS_ForCulture = 1
    MWS_ForStain = 2
    MWS_ForSensi = 3

    MRT_GenSen = "S"        'General Sensi
    MRT_MicSen = "C"        'MIC Sensi
    MRT_Stain = "G"         'Gram Stain
    MRT_AFC = "M"           'AFB Culture
    MRT_AFS = "B"           'AFB Stain
    MRT_AFN = "A"           'AFB Sensi
    MRT_FUC = "F"           'Fungus Culture
    MRT_FUS = "N"           'Fungus Stain
    MRT_Both = "B"
    MRT_OTHER = "O"

    MNM_GSen = "일반 감수성"        ' 감수성 입력화면용 표현
    MNM_MSen = "MIC 감수성"
    MNM_AFC = "AFB/Fungus Culture"
    MNM_AFS = "AFB/Fungus Stain"

    MCD_GSen = "GS"                 ' 검사에 따른 균종별 항생제 분류 구분자
    MCD_MSen = "MS"

' 미생물 감수성 결과 존재 유무 (SenFg in lab404)
    MRT_NoSenRst = ""
    MRT_SenRst = "Y"
    MRT_SenRstCd = "RISPN-"

' 기타검사 세부결과 존재 유무
    ERT_ValRst = "Y"  '(ValFg in lab351)
    ERT_TxtRst = "Y"  '(TxtFg in lab351)

' 기타검사 Worksheet Flag
    EWS_OK = "1"
    EWS_NO = "0"

    ERT_REPORT = "BM"

' QC Level 코드
    QC_LEVEL_HIGH = "H"
    QC_LEVEL_NORMAL = "N"
    QC_LEVEL_LOW = "L"

' Interface 오류 데이타 상수
    CS_EqpError = "Error"
    RI_WORKAREA = "RI"
    
' 번호부여정보
    COM99_LIS_Specimen = "L001"           '검체번호
    COM99_LIS_LabNo = "L002"              '접수번호
    COM99_LIS_WorkNo = "L003"             '일반Worksheet Unit
    COM99_LIS_SeqNo = "L004"              '접수순번(WorkArea별)
    COM99_LIS_WSUnit = "L009"             '미생물Worksheet Unit
    COM99_LIS_Report = "L010"             '결과지 출력일자 관리
    LIS_BarFormat = "0########"           '9자리(성모자애병원)
    CS_BarFormat = "000000000"
    MIC_WorkArea = "04"
    CBC_WorkArea = "01"
End Sub
