Attribute VB_Name = "modIISCommon"
'-----------------------------------------------------------------------------'
'   파일명  : modIISCommon.bas
'   내  용  : 공통코드 클래스
'   버  전  :
'          - 미생물 일반감수성(IISGENSENSI), MIC(IISMIC) 결과유형 변수추가
'          - PT(%)값의 결과코드(IISPT) 변수추가
'          - CONCAT(FCONCAT) 변수추가
'          - 검사항목별 결과코드(mCRESULTCD) 변수추가
'          - CBC Workarea(mIISCBCWA) 변수추가
'          - 하한값(mIISLMTLOW), 상한값(mIISLMTHIGH) 결과코드 변수추가
'-----------------------------------------------------------------------------'

Option Explicit

'## Database 연결정보
Public mDbCon           As ADODB.Connection     '메인Db Connection
Public mCliCon          As ADODB.Connection     'ClientDb Connection
Public mDbType          As String               'DB Type (0:Oracle, 1:Sybase, 2:MS-SQL, 3:ACCESS)
Public mSource          As String               'Data Source
Public mCatalog         As String               'Initial Catalog
Public mUid             As String               'User ID
Public mPwd             As String               'Password
Public mUserCancel      As Boolean              '사용자가 DB설정을 종료하면 True

'## Error 컬렉션
Public mError           As clsIISError          '에러 클래스

'## 레지스트리의 DB설정 정보
Public Const cDBSERVER  As String = "DbServer"
Public Const cDBTYPE    As String = "DbType"
Public Const cSOURCE    As String = "Source"
Public Const cCATALOG   As String = "Catalog"
Public Const cUID       As String = "Uid"
Public Const cPWD       As String = "Pwd"

'## 레지스트리 정보
Public mAppName         As String       'App Name
Public mExePath         As String       'EXE 파일경로
Public mLogPath         As String       'Log 파일경로
Public mClientDbPath    As String       'ClientDb 경로+파일명
Public mIniPath         As String       'INI 파일경로

'## 사용자 정보
Public mEmpId           As String       '사용자 아이디
Public mEmpNm           As String       '사용자 이름

'## mdiIISMain Form
Public mMainFrm         As Object       'mdiIISMain 폼

'## 통신기호
Public Const mENQ As Long = &H5         'Chr(5),  ""
Public Const mACK As Long = &H6         'Chr(6),  ""
Public Const mSTX As Long = &H2         'Chr(2),  ""
Public Const mETB As Long = &H17        'Chr(23), ""
Public Const mETX As Long = &H3         'Chr(3),  ""
Public Const mEOT As Long = &H4         'Chr(4),  ""
Public Const mNAK As Long = &H15        'Chr(21), ""
Public Const mSOH As Long = &H1         'Chr(1),  ""
Public Const mDLE As Long = &H10        'Chr(16), ""
Public Const mSYN As Long = &H16        'Chr(22), ""

'## 병원별 정보
Public mPROJECTCODE     As String       'Project Code
Public mPROJECTTYPE     As String       'Project Type(A:자사, B:타사, C:독립)
Public mHOSPITALNM      As String       '병원이름
Public mSPCLEN          As Long         'SPC Length
Public mSPCYYLEN        As Long         'SPCYY Length
Public mSPCNOLEN        As Long         'SPCNO Length
Public mIISNEGATIVE     As String       'Negative
Public mIISPOSITIVE     As String       'Positive
Public mIISGRAYZONE     As String       'Grayzone
Public mIISERROR        As String       'Error
Public mIISSPCLEN       As Long         '검체구분길이
Public mIISSPCSERUM     As String       'Sereum
Public mIISSPCURINE     As String       'Urine
Public mIISSPCPLASMA    As String       'Plasma
Public mIISSPCCSF       As String       'CSF
Public mIISSPCBLOOD     As String       'Blood
Public mIISSPCFLUID     As String       'Boyd Fluid
Public mIISSPCCAPD      As String       'CAPD
Public mIISQCLOW        As String       'QC Low Level
Public mIISQCNORMAL     As String       'QC Normal Level
Public mIISQCHIGH       As String       'QC High Level
Public mIISPANICCHECK   As String       'Panic 체크 테이블
Public mIISMICWA        As String       '미생물 Workarea
Public mIISMQTCD        As String       '미생물 감수성결과 정도코드
Public mIISGENSENSI     As String       '미생물 일반감수성 결과코드
Public mIISMIC          As String       '미생물 MIC 결과코드
Public mIISSERUMINDEX   As String       'Hitachi7600장비의 Serum Index 사용유무(0:무,1:유)
Public mIISPT           As String       'PT(%)값이 100이상일때 결과코드
Public mIISCBCWA        As String       'CBC Workarea

Public mIISREACTIVE     As String       'Reactive
Public mIISNREACTIVE    As String       'NonReactive
Public mIISWPOSITIVE    As String       'WaekPositive

'   - 상한값, 하한값 변수
Public mIISLMTLOW       As String       '하한값
Public mIISLMTHIGH      As String       '상한값

'## 공통코드 인덱스
Public mCODE            As String       'CDINDEX
Public mCSPCCD          As String       '검체
Public mCWACD           As String       'WorkArea
Public mCDETAILCD       As String       '상세항목
Public mCPANELCD        As String       '그룹항목
Public mCREPEATCD       As String       '다빈도처방
Public mCLOCATIONCD     As String       '건물코드
Public mCVANDCD         As String       '업체코드
Public mCFOOTNOTECD     As String       'FootNote
Public mCSPCRMKCD       As String       '검체 Remark
Public mCACCRSNCD       As String       '접수취소 사유
Public mCMDYRSNCD       As String       '결과수정 사유
Public mCQCREJECTCD     As String       'QC Reject 사유
Public mCMnmCd          As String       '균코드
Public mCRESULTCD       As String       '검사항목별 결과코드

'## 테이블
Public mTHIS001         As String       '환자 마스터
Public mTHIS002         As String       '병동 마스터
Public mTHIS003         As String       '진료과 마스터
Public mTHIS004         As String       '처방의 마스터
Public mTHIS005         As String       '직원 마스터1 (로그인정보)
Public mTHIS006         As String       '직원 마스터2 (직원정보)

Public mTCOM001         As String       '공통코드 마스터
Public mTIIS001         As String       '공통코드 마스터1
Public mTIIS002         As String       '공통코드 마스터2
Public mTIIS003         As String       '템플릿 마스터

Public mTIIS101         As String       '검체 마스터
Public mTIIS102         As String       '검사항목 마스터
Public mTIIS103         As String       '지정검체 마스터
Public mTIIS104         As String       '참고치 마스터
Public mTIIS105         As String       '템플릿 마스터
Public mTIIS106         As String       '풋노트 마스터
Public mTIIS107         As String       'QC 소견마스터

Public mTIIS201         As String       '처방내역(H)
Public mTIIS202         As String       '처방내역(B)
Public mTIIS203         As String       '접수내역
Public mTIIS204         As String       '결과내역
Public mTIIS205         As String       '수정전 결과
Public mTIIS206         As String       '재검전 결과

Public mTIIS301         As String       'QC 컨트롤(H)
Public mTIIS302         As String       'QC 컨트롤(B)
Public mTIIS303         As String       'QC 마스터(H)
Public mTIIS304         As String       'QC 마스터(B)
Public mTIIS305         As String       'QC 스케줄
Public mTIIS306         As String       'QC 접수
Public mTIIS307         As String       'QC 결과
Public mTIIS308         As String       'QC 소견내역

Public mTIIS401         As String       '장비 마스터
Public mTIIS402         As String       '장비통신 마스터
Public mTIIS403         As String       '장비 검사항목 마스터(H)
Public mTIIS404         As String       '장비 검사항목 마스터(B)
Public mTIIS405         As String       '장비전송내역
Public mTIIS406         As String       '장비오류내역

Public mTIIS501         As String       '미생물 WorkSheet(H)
Public mTIIS502         As String       '미생물 WorkSheet(B)
Public mTIIS503         As String       '미생물 WorkSheet 추가내역
Public mTIIS504         As String       '미생물 결과내역
Public mTIIS505         As String       '미생물 감수성 결과내역

'   - 특수검사 결과내역 테이블변수 추가
Public mTIIS601         As String       '특수검사 결과내역

Public mTACC203         As String       'ClientDb 접수내역
Public mTACC204         As String       'ClientDb 결과내역

'## 필드
'HIS001 (환자 마스터)
Public mFPTID           As String       '환자ID
Public mFPTNM           As String       '이름
Public mFJUMIN          As String       '주민번호
Public mFSEX            As String       '성별
Public mFAGE            As String       '나이

'HIS002 (부서 마스터)
Public mFDEPTCD         As String       '부서코드
Public mFDEPTNM         As String       '부서명

'HIS003 (병동 마스터)
Public mFWARDCD         As String       '병동코드
Public mFWARDNM         As String       '병동명

'HIS004 (처방의 마스터)
Public mFDOCTCD         As String       '처방의코드
Public mFDOCTNM         As String       '처방의명

'HIS006 (직원 마스터, 직원정보)
Public mFEMPID          As String       '직원ID
Public mFEMPNM          As String       '직원이름

'공통코드 마스터2
Public mFSPCCD          As String       '검체코드
Public mFSPCNM          As String       '검체명

'## Database별 구분자
Public mFCONCAT         As String       'Concatenate 연산자(Oracle:||, MS-SQL:+)
