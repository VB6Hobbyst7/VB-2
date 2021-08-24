Attribute VB_Name = "BasDbRef"
Option Explicit
        
    '********************************************************************
    ' 환자 기본정보 Data Base Reference Field
    '********************************************************************
    '--------------------------------------------------------------------
    '1) 환자 기본 인적 Data Base (PbsInf)
    '--------------------------------------------------------------------
Type PbsInfRec
    PbsChtNum  As String * 8   'PbsInfKey 1  챠트번호
    PbsPatNam  As String       '          2  수진자성명
    PbsResNum  As String       '          3  주민등록번호
    PbsZipCod  As String       '          4  우편번호
    PbsDtlAdr  As String       '          5  상세주소
    PbsPhnNum  As String       '          6  전화번호
    PbsNewDte  As String       '          7  신환일자
    PbsMdcTyp  As String       '          8  진료형태
    PbsSexTyp  As String       '          9  성별
    PbsArtYon  As String       '          10 인공신장
    PbsSpcFlg  As String       '          11 특기사항
    PbsRefCmd  As String       '          12 참조사항
    PbsUpdTim  As String       '          13 수정일자
    PbsUidCod  As String       '          14 담당자
    PbsOsuYon  As String       '          15 진찰권 발행여부
    PbsOldNum  As String       '          16 이전 챠트번호
    PbsIcmYon  As String       '          17 입원여부
    PbsCruYon  As String       '          18 병원직원여부 Y/N
    PbsHndPhn  As String       '          19 핸드폰 번호
    PbsE_Mail  As String       '          20 E-Mail
    PbsMomCht  As String       '          21 신생아 구분을 위한 어머니차트
    PbsRecUid  As String       '          22 추천인 아이디
    PbsRecNam  As String       '          23 추천인 성명
    PbsPatDte  As String       '          24 생리시작일(LMP)
    
    
End Type
    '------------------------------------
    ' 환자 기타 인적 Data Base (PbsInf) =정신병원
    '------------------------------------
Type PspInfRec
    PspChtNum  As String * 8   'PspInfKey 1  챠트번호
    PspParNam  As String       '          2  보호자성명
    PspRelTyp  As String       '          3  보호자관계
    PspPhnNum  As String       '          4  집전화번호
    PspPatEdu  As String       '          5  학력
    PspMryYon  As String       '          6  결혼여부
    PspParYon  As String       '          7  부모유무
    PspPatJob  As String       '          8  환자직업
    PspPatRlg  As String       '          9  종교
    PspUpdTim  As String       '          10 수정일자
    PspUidCod  As String       '          11 담당자

    '-------추가
    PspZipCod  As String       '          12 우편번호
    PspDtlAdr  As String       '          13 상세주소
    PspSchCod  As String       '          14 학력2
    PspDtlJob  As String       '          15 상세직업
    PspMrgCod  As String       '          16 결혼구분

    PspComPhn  As String       '          17 회사전화번호
    PspPcsPhn  As String       '          18 핸드폰
    PspCslUid  As String       '          19 면접자(Counselling)
    PspCstUid  As String       '          20 의뢰자
    '-------추가
End Type
    
    
    '--------------------------------------------------------------------
    '2) 보험인적 사항 Data Base (PmdInf)
    '--------------------------------------------------------------------
Type PmdInfRec
    PmdChtNum  As String * 8   'PmdInfKey 1  챠트번호
    PmdInsCod  As String       'PmdInfKey 2  보험유형
    PmdInsSeq  As String * 2   'PmdInfKey 3  유형순서
    PmdAssCod  As String       '          4  조합기호
    PmdInsNum  As String       '          5  증번호
    PmdPasNam  As String       '          6  피보험자성명
    PmdResNum  As String       '          7  피보험자 주민번호
    PmdRelTyp  As String       '          8  피보험자와관계
    PmdDsoNum  As String       '          9  장애인 수첩번호
    PmdXplNum  As String       '          10 국가유공자 번호
    PmdRcuNum  As String       '          11 보호정신과 요양승인번호
    PmdRgnYon  As String       '          12 타진료권 승인 여부
    PmdAdpDte  As String       '          13 자격 취득일자
    PmdExpDte  As String       '          14 자격 종료일자
    PmdEntNam  As String       '          15 사업체명
    PmdUpdDtm  As String       '          15 수정일자
    PmdUidCod  As String       '          16 담당자
    PmdXplAss  As String       '          17 장애인조합기호
End Type
    
    '--------------------------------------------------------------------
    '3) 산재인적 사항 PwkInf
    '--------------------------------------------------------------------
Type PwkInfRec
    PwkChtNum  As String * 8   'PwkInfKey 1  챠트번호
    PwkInsCod  As String       'PwkInfKey 2  보험유형
    PwkInsSeq  As String * 2   'PwkInfKey 3  유형순서
    PwkAssCod  As String       '          4  복지공단기호
    PwkEntNam  As String       '          5  사업체명
    PwkSetCod  As String       '          6  결정코드
    PwkRcuNum  As String       '          7  요양승인번호
    PwkRcuDte  As String       '          8  요양승인일자
    PwkDsaDte  As String       '          9  재해발생일자
    PwkReqRcu  As String       '          10 승인연장신청일
    PwkMcrDte  As String       '          11 진료개시일
    PwkInjRgn  As String       '          12 상병부위
    PwkCurRst  As String       '          13 치료결과
    PwkAdpDte  As String       '          14 자격 취득일자
    PwkExpDte  As String       '          15 자격 종료일자
    PwkUpdDtm  As String       '          16 수정일자
    PwkUidCod  As String       '          17 담당자
End Type
    
    
    '--------------------------------------------------------------------
    '4) 자보인적 사항 PcrInf
    '--------------------------------------------------------------------
Type PcrInfRec
    PcrChtNum  As String * 8   'PcrInfKey 1  챠트번호
    PcrInsCod  As String       'PcrInfKey 2  보험유형
    PcrInsSeq  As String * 2   'PcrInfKey 3  유형순서
    PcrInjDte  As String       '          4  상해일시
    PcrFstDte  As String       '          5  초진일시
    PcrAssCod  As String       '          6  보험회사코드
    PcrVeiNum  As String       '          7  차량번호
    PcrVeiOwn  As String       '          8  차량소유자
    PcrAcpNum  As String       '          9  접보번호
    PcrAcpDte  As String       '          11 접보처리일자  '96/05/22
    PcrCarUid  As String       '          10 자보담당자    '
    PcrAdpDte  As String       '          12 자격 취득일자
    PcrExpDte  As String       '          13 자격 종료일자
    PcrUpdDtm  As String       '          14 수정일자
    PcrUidCod  As String       '          15 담당자
    PcrCarRem  As String       '          16 자보 특기사항
    PcrLmtAmt  As String       '          17 자보 한도액
End Type
    
    
    '--------------------------------------------------------------------
    '5) 타지역승인정보 OrgInf
    '--------------------------------------------------------------------
Type OrgInfRec
    OrgChtNum  As String * 8   'OrgInfkey 챠트번호
    OrgInsCod  As String       '          보험유형
    OrgRcuNum  As String       '          타지역승인번호
    OrgRcuTyp  As String       '          승인구분
    OrgAdpDte  As String       '          적용일자
    OrgExpDte  As String       '          종료일자
    OrgDepCod  As String       '          진료과
    OrgIcdCod  As String       '          상병명
End Type
    
    '--------------------------------------------------------------------
    '6) 연대보증인 Data Base (GrnInf)
    '--------------------------------------------------------------------
Type GrnInfRec
    GrnOcmNum  As String * 10   'GrnInfKey 1  내원번호
    GrnChtNum  As String * 8   '           2  환자차트번호
    GrnPatNam  As String       '           3  성명
    GrnResNum  As String       '           4  주민등록번호
    GrnDtlAdr  As String       '           5  상세주소
    GrnPhnNum  As String       '           6  전화번호
    GrnSexTyp  As String       '           7  성별
    GrnRelTyp  As String       '           8  환자와의관계
    GrnComNam  As String       '           9  직장명
    GrnPrtNam  As String       '           10 부서및직위
    GrnComTel  As String       '           11 회사전화번호
    GrnEtc     As String       '           12 기타 참조 사항
End Type
    
    
    '********************************************************************
    ' 외래 내원 기본정보 Data Base Referance Field
    '********************************************************************
    '--------------------------------------------------------------------
    '1) 외래 내원 환자정보 OcmInf
    '--------------------------------------------------------------------
Type OcmInfRec
    OcmNum     As String * 10  'OcmInfKey 1  내원번호
    OcmChtNum  As String * 8   '        1 2  챠트번호
    OcmComStt  As String       '        2 3  접수상태(Add,Cancel)
    OcmDepCod  As String       '        3 4  진료과
    OcmDtrCod  As String       '        4 5  주치의
    OcmComRut  As String       '        5 6  내원경로
    OcmAcpDtm  As String       '        6 7  접수일시
    OcmInsCod  As String       '        7 8  보험유형
    OcmInsSeq  As String * 2   '        8 9  보험유형순서
    OcmDgsNfs  As String       '        9 10 초재진구분
    OcmDgsDnh  As String       '       10 11 주,야,공휴구분
    OcmFreRsn  As String       '       11 12 진찰료미발생사유
    OcmSpcYon  As String       '       12 13 특진여부
    OcmArtYon  As String       '       13 14 인공신장여부
    OcmRsuYon  As String       '       14 15 진찰권재발행여부
    OcmDgsCht  As String       '       15 16 의무 기록지구분
    OcmMdcDay  As String       '       16 17 투약일수
    OcmNul     As String       '       17 18 Null  '응급실 처리사항(97/12/11)
    OcmArrStt  As String       '       18 19 도착상태
    OcmLevTim  As String       '       19 20 응급실 퇴실시간
    OcmLevRst  As String       '       20 21 응급실 퇴실결과
    OcmEmgCod  As String       '       21 22 응급등급( "Y" or "")
    OcmUpdTim  As String       '       22 23 수정일시
    OcmUidCod  As String       '       23 24 담당자
    OcmNonIns  As String       '       24 25 보험100%    -->97.4.9
    OcmIcmNum  As String       '       25 26 입원전환내원번호
    OcmMdcTyp  As String       '       26 27 의료부대(8), 기타수입(7)
    OcmTrmDtr  As String       '       27 28 처치의사
    OcmArrDtm  As String       '       28 29 응급실 도착일시
    OcmImgYon  As String       '       29 30 수입조정여부
    OcmEmgKnd  As String       '       30 31 사고종류
    OcmActFlg  As String       '       31 32 진찰료금액분리발생사유(물리치료:PHA)
    OcmEndStt  As String       '       32 33 마감당시 내원 상태
    OcmHanAmt  As String       '       33 34 한방비급여금액
    OcmHanCmt  As String       '       34 35 한방특기사항(탕약명칭) or 추가진단
    OcmPhyRev  As String       '       35 36 미래처방 진찰료 선택
    OcmRomYon  As String       '       36 37 낯병동 사용여부
    OcmFutDay  As String       '       37 38 미래처방 일자지정
    OcmOutCod  As String       '       38 39 원외예외구분
    OcmOutNum  As String * 5   '       39 40 교부번호
    OcmRcpCmt  As String       '       40 41 수납전달사항(Bestian 치과 때문에 생성함) '2002/05/06
    OcmCvtYon  As String       '       41 42 입원예정여부
    OcmCvtDtm  As String       '       42 43 입원예정일시
    'OcmCvtStt  As String       '       43 44 입원Ststus
    OcmCasStb  As String        '수납 대기 상태
End Type
    
    '--------------------------------------------------------------------
    '2) 외래 처방 정보 OspInf
    '--------------------------------------------------------------------
Type OspInfRec

    OspOcmNum  As String * 10  'OspInfKey 내원번호
    OspOdrNum  As String * 4   'OspInfKey 처방번호
    OspOdrSeq  As String * 5   'OspInfKey 처방순서
    OspOdrCod  As String       '1         처방코드
    OspOdrTyp  As String       '2         처방유형
    OspOdrStt  As String       '3         처방진행상태
    OspStkStt  As String       '4         재고진행상태
    OspFeeCod  As String       '5         수가코드
    OspAddCod  As String       '6         가산코드
    OspDepCod  As String       '7         진료과목
    
    OspSlpDep  As String       '8         처방전달부서
    
    OspSlpCod  As String       '9         처방전달범주
    OspItmCod  As String       '10        처방항목코드
    OspOdrDtm  As String       '11        처방일시
    OspOdrPrc  As String       '12        단가1
    OspOdrSib  As String       '13        단가2
    OspOdrQty  As String       '14        투여량
    OspOdrTms  As String       '15        횟수
    OspOdrDay  As String       '16        일수
    OspUsgCod  As String       '17        용법
    OspMthCod  As String       '18        행위코드
    OspSpmcod  As String       '19        검체코드
    OspInsYon  As String       '20        급여/비급여구분
    OspInsCod  As String       '21        보험유형
    OspInsSeq  As String       '22        보험유형순서
    OspDgsEtc  As String       '23        기타가산
    OspDgsRol  As String       '24        Right/Left
    OspOprDnh  As String       '25        주,야,공,휴
    OspOprDtm  As String       '26        수술시각
    OspPrePay  As String       '27        선처방일수
    OspEmgYon  As String       '28        응급여부
    OspSpcYon  As String       '29        특진구분
    OspSlpAmt  As String       '30        계산금액
    OspIncCod  As String       '31        수입원
    OspSotCod  As String       '32        처방항목
    
    OspEntDtm  As String       '33        입력일시
    
    OspDtrCod  As String       '34        처방담당
    OspCasYon  As String       '35        수납여부
    OspCasDtm  As String       '36        수납일시
    OspUidCod  As String       '37        수정담당
    OspUpdDte  As String       '38        수정일시
    OspPreDtm  As String       '39        예약일시
    OspSplYon  As String       '40        특기사항여부
    OspSplCmt  As String       '41        특기사항
    OspChkStt  As String       '42        전달여부  "-1" : 취소처방을 지원부서에서 확인,"0" : 처방을 지원부서에서 확인, "1"이상 : 지원부서에서 시행할 일수
    OspMdcNum  As String       '43        투약번호
    OspCanMdc  As String       '44        취소시의 투약번호
    OspStgCod  As String       '45        수탁검사시 검사의뢰기관 코드
    OspBasUnt  As String       '46        의사사용량
    OspMntUsg  As String       '47        정신과사용용법
    OspXryPtb  As String       '48        Xray Portable
    OspDtrPrt  As String       '49        입력부서
    OspUpdPrt  As String       '50        수정부서
    OspImgYon  As String       '51        수입조정여부
    OspCanNum  As String       '52        취소Order Number
    OspCanSeq  As String       '52        취소Order Seq
    OspDenRgn  As String       '53        코드별 치식
    OspOdrNam  As String       '54        코드이름
    OspOdrNo   As String       '55        오더번호(EMR에서 화면에 표시하기 위한 순서 및 Group)
    OspQtyGbn  As String       '56        *  or  #     투여량 과 횟수를 곱해야 할지 나누어야 할지... 총투여량인지 일회투여량인지...
End Type
    
    '--------------------------------------------------------------------
    '3) 외래 영수증 OrpInf
    'Primary Key    OrpInf          (K-1,K-2)       OcmNum/RvnTyp
    'Index-1        OrpInfCht       (D-1)           ChtNum
    'Index-2        OrpInfDtmRvn    (D-27,K-2)      UpdDtm/RvnTyp
    'Index-3        OrpInfMan       (D-24)          ManNum
    'Index-4        OrpInfRcpMan    (D-22,D-24)     RcpNum/ManNum
    '--------------------------------------------------------------------
Type OrpInfRec
    OrpOcmNum  As String * 10  'OrpInfKey 1  내원번호
    OrpRvnTyp  As String       'OrpInfKey 2  접수,수납구분(M:접수,O:수납,E:기타수입,T:?)
    OrpChtNum  As String * 8   '          1  챠트번호
    OrpDepCod  As String       '          2  진료과
    OrpDtrCod  As String       '          3  주치의
    OrpInsCod  As String       '          4  보험유형
    OrpInsSeq  As String       '          5  보험유형순서
    OrpRcpStt  As String       '          6  상태구분(1:표준)
    OrpTotAmt  As String       '          7  진료비총액
    OrpInsAmt  As String       '          8  급여총액           InsTot
    OrpNonAmt  As String       '          9  비급여총액         NonOwn
    OrpCorAmt  As String       '          10 조합청구액
    OrpOwnAmt  As String       '          11 급여 일부부담액    InsOwn
    OrpTotOwn  As String       '          12 본인부담총액
    OrpSpcAmt  As String       '          13 특진료
    OrpDisAmt  As String       '          14 할인금액
    OrpFutAmt  As String       '          15 후불금액
    OrpAskAmt  As String       '          16 환자청구액
    OrpOldAmt  As String       '          17 기수납액
    OrpNewAmt  As String       '          18 수납액
    OrpRcpYon  As String       '          19 수납여부
    OrpRetRsn  As String       '          20 환불사유
    OrpPubYon  As String       '          21 영수증발행여부
    OrpRcpNum  As String * 10  '          22 영수증번호
    OrpOldNum  As String * 10  '          23 이전영수증번호
    OrpManNum  As String * 10  '          24 주 영수증번호
    OrpMdcNum  As String       '          25 약번호
    OrpBknDtm  As String       '          26 예약일시
    OrpUpdDtm  As String       '          27 계산일자           CalDtm
    OrpUidCod  As String       '          28 담당자코드
    OrpPrcFun  As String       '          29 처리구분
    OrpMdcTyp  As String       '          30 기타수입원(종진,기타수입..일 경우 마감을 위해 수입원을 써준다.)
    OrpDimAmt  As String       '          31 감면액
    OrpNonIns  As String       '          32 본인100여부
    OrpEtcDtl  As String       '          33 기타수입원타입
    OrpCarFut  As String       '          34 카드입금액
    OrpOutNum  As String       '          35 교부번호
    OrpAccDte  As String       '          36 회계일자
    OrpFodAmt  As String       '          37 기카드입금액
    
    '20040101..HTS..add
    OrpNinAmt  As String       '          40 전액본인부담
    
End Type
    
    
    '--------------------------------------------------------------------
    '4) 외래 영수증 변경 OhtInf
    '--------------------------------------------------------------------
Type OhtInfRec
    OhtRcpNum  As String * 10  'OhtInfKey 1  영수증번호
    OhtOcmNum  As String * 10  '          1  내원번호
    OhtRvnTyp  As String       '          2  접수,수납구분
    OhtChtNum  As String * 8   '          3  챠트번호
    OhtDepCod  As String       '          4  진료과
    OhtDtrCod  As String       '          5  주치의
    OhtInsCod  As String       '          6  보험유형
    OhtInsSeq  As String * 2   '          7  보험유형순서
    OhtRcpStt  As String       '          8  상태구분 (2:환불) - 주석추가
    OhtTotAmt  As String       '          9  진료비총액
    OhtInsAmt  As String       '          10 급여총액           InsTot
    OhtNonAmt  As String       '          11 비급여총액         NonOwn
    OhtCorAmt  As String       '          12 조합청구액
    OhtOwnAmt  As String       '          13 급여 일부부담액    InsOwn
    OhtTotOwn  As String       '          14 본인부담총액
    OhtSpcAmt  As String       '          15 특진료
    OhtDisAmt  As String       '          16 할인금액
    OhtFutAmt  As String       '          17 후불금액
    OhtAskAmt  As String       '          18 환자청구액
    OhtOldAmt  As String       '          19 기수납액
    OhtNewAmt  As String       '          20 수납액
    OhtRcpYon  As String       '          21 수납여부
    OhtRetRsn  As String       '          22 환불사유
    OhtPubYon  As String       '          23 영수증발행여부
    OhtOldRcp  As String * 10  '          24 영수증번호
    OhtOldNum  As String * 10  '          25 이전영수증번호
    OhtManNum  As String * 10  '          26 주 영수증번호
    OhtMdcNum  As String       '          27 약번호
    OhtBknDtm  As String       '          28 예약일시
    OhtUpdDtm  As String       '          29 계산일시           CalDtm
    OhtUidCod  As String       '          30 담당자코드
    OhtPrcFun  As String       '          31 처리구분
    OhtMdcTyp  As String       '          32 기타수입원(종진,기타수입..일 경우 마감을 위해 수입원을 써준다.)
    OhtDimAmt  As String       '          33 감면액
    OhtNonIns  As String       '          34 본인100여부
    OhtEtcDtl  As String       '          35 기타수입원상세
    OhtCarFut  As String       '          36 카드미수금액
    OhtOutNum  As String       '          37 교부번호
    OhtAccDte  As String       '          38 회계일자
    OhtFodAmt  As String       '          39 기수납금액
    
    '----------------------------
    '20040101..HTS..add
    OhtNinAmt  As String       '          42 전액본인부담
    '----------------------------
End Type
    
    
    '--------------------------------------------------------------------
    '5) 외래 영수증 상세내역 정보 OdlInf
    '--------------------------------------------------------------------
Type OdlInfRec
    OdlRcpNum  As String * 10  'OdlInfKey 영수증번호
    OdlIncCod  As String * 2   'OdlInfKey 수입원
    OdlChtNum  As String * 8   ' 3        챠트번호
    OdlInsCod  As String       ' 4        보험유형
    OdlInsSeq  As String * 2   ' 5        유형순서
    OdlDepCod  As String       ' 6        진료과목
    OdlInsAct  As String       ' 7        급여행위
    OdlInsStf  As String       ' 8        급여재료              InsMat
    OdlNonAct  As String       ' 9        비급여행위
    OdlNonStf  As String       ' 10       비급여재료            NonMat
    OdlInsAmt  As String       ' 11       급여총액
    OdlNonAmt  As String       ' 12       비급여액              InsOwn
    OdlOwnAmt  As String       ' 13       급여일부(본인)부담액  NonAmt
    OdlTotOwn  As String       ' 14       총본인부담액
    OdlSpcAmt  As String       ' 15       특진료
    
    '=========================================
    '20040101..HTS..add
    OdlNinAct  As String       ' 16       전액본인부담행위
    OdlNinStf  As String       ' 17       전액본인부담재료
    OdlNinAmt  As String       ' 18       전액본인부담
    '=========================================
End Type
    
    '--------------------------------------------------------------------
    '6) 내원환자 상병명정보 OicInf (입원,외래 공통)
    '--------------------------------------------------------------------
Type OicInfRec
    OicOcmNum  As String * 10  'OicInfkey 내원번호
    OicSeq     As String * 2   'OicInfKey 순서
    OicChtNum  As String * 8   '          챠트번호
    
    OicIcdCod  As String       '          상병기호
    OicIcdPri  As String       '          우선순위
    OicEeeCod  As String       '          E-Code
    OicVeeCod  As String       '          V-Code
    OicOprYon  As String       '          수술여부
    OicDenRgn  As String       '          치과상부위
    OicDgnDte  As String       '          진단일자
    OicDepCod  As String       '          진단과목
    OicCurRst  As String       '          치료결과
    OicCurGrd  As String       '          진단등급
    OicAddIcd  As String       '
    OicFinIcd  As String        '확진병명 여부
    OicSpcCmt  As String        '특기사항
    
    '20030228 lek add for 상병명 추가
    OicIdcNam As String         '상병명
    
End Type
    
    '--------------------------------------------------------------------
    '7) 외래, 입원(?) 원무과 참조사항 - 첨단병원요청사항
    '--------------------------------------------------------------------
Type OscInfRec
    OscChtNum  As String * 8    'OscInfkey 챠트번호
    OscSplCmt  As String        '원무과 특기사항
End Type
    
    '////////////////////////////////////////////////////////
    '//카드미수내역
    '////////////////////////////////////////////////////////
'Type CrdInfRec
'    CrdCrdNum  As String       '           카드번호
'    CrdCrdApp  As String       '           승인번호
'    CrdAdpAmt  As String       '           계정금액
'    CrdUidCod  As String       '           입력자
'End Type
    
    '--------------------------------------------------------------------
    '7) 내원 계정정보 OacInf
    '--------------------------------------------------------------------
Type OacInfRec
    OacRcpNum  As String * 10  'OacInfkey  영수증번호
    OacAccCod  As String       'OacInfKey  계정코드
    OacChtNum  As String * 8   '           진찰권번호
    OacOcmNum  As String * 10  '           내원번호
    OacAccAmt  As String       '           계정금액
    OacAccDsc  As String       '           대표계정내역
    OacAccRat  As String       '           대표계정률
    OacAccDgs  As String       '           계정구분
    OacFdgRat  As String       '           접수비 초진계정률
    OacFdgAmt  As String       '           접수비 초진계정금액
    OacSdgRat  As String       '           접수비 재진계정률
    OacSdgAmt  As String       '           접수비 재진계정금액
    OacCalRat  As String       '           수납   계정률
    OacCalAmt  As String       '           수납   계정금액
    OacCalSeq  As String       '           수납계정순서
    OacEmpCod  As String       '           사원번호
    OacRelCod  As String       '           관계코드
    '20010816 카드수납때문에 없앤다...yk
    'OacCrdMax  As String       '           카드남긴 수
    'OacCrdDat(1 To 10) As CrdInfRec     '  카드미수
End Type
    
    '********************************************************************
    ' 입원 내원 기본정보 Data Base Reference
    '********************************************************************
    '--------------------------------------------------------------------
    '1) 입원 내원 환자정보
    'Primary Key(IcmInf)            (K-1)               IcmOcmNum
    'Index-1    (IcmInfAcpInsLev)   (D-2, D-8, D-16)    AcpStt/InsCod/LevDtm
    'Index-2    (IcmInfAcpLevIns)   (D-2, D-16,D-8)     AcpStt/LevDtm/InsCod
    'Index-3    (IcmInfAcpStt)      (D-3, D-2)          AcpDtm/AcpStt
    'Index-4    (IcmInfChtDtm)      (D-1, D-3)          ChtNum/AcpDtm
    'Index-5    (IcmInfInsCht)      (D-8, D-1)          InsCod/ChtNum
    'Index-6    (IcmInfLevStt)      (D-16,D-2)          LevDtm, AcpStt
    'Index-7    (IcmInfNss)         (D-12)              NssCod
    '--------------------------------------------------------------------
Type IcmInfRec
    IcmOcmNum  As String * 10  'IcmInfKey 입원번호
    IcmChtNum  As String * 8   '1         챠트번호
    IcmAcpStt  As String       '2         입원상태(A 재원,D 퇴원,D 퇴원,F 퇴원예정, R 입원예정,S 가퇴원 )
    IcmAcpDtm  As String       '3         입원일시
    IcmArrPat  As String       '4         도착경로(1 타요양기관, 2 응급구조대, 3 기타)
    IcmAcpRut  As String       '5         입원경로(1 응급실, 2 외래)
    IcmDepCod  As String       '6         진료과목
    IcmDtrCod  As String       '7         담당의사
    IcmInsCod  As String       '8         보험유형
    IcmInsSeq  As String * 2   '9         보험유형순서
    IcmDupYon  As String       '10        이중보험여부("N","Y")
    IcmConYon  As String       '11        협의진단여부
    IcmNssCod  As String       '12        병동코드
    IcmRomCod  As String       '13        병실코드
    IcmBedCod  As String       '14        병상코드
    IcmLevCnt  As String       '15        퇴원번호
    IcmLevDtm  As String       '16        퇴원일시
    IcmNtcDtm  As String       '17        퇴원예정일시
    IcmOutDtm  As String       '18        외출일시
    IcmRtnDtm  As String       '19        귀원일시
    IcmDgsNfs  As String       '20        초재진구분 96/03/22 추가
    IcmDgsDnh  As String       '21        주,야,휴일구분 96/03/22 추가
    IcmSpcYon  As String       '22        특진여부 96/03/22 추가
    IcmUpdDtm  As String       '23        수정일시
    icmUidCod  As String       '24        담당자
    IcmNonIns  As String       '25        본인100%여부
    IcmImgYon  As String       '26        수납조정여부
    IcmRemark  As String       '27        내원당 특이사항
    IcmRcpCmt  As String       '28        수납전달사항(Bestian 치과 때문에 생성함) '2002/05/06
    IcmMomCht  As String       '29        신생아를 구분하기위한 어머니 차트
    IcmRptDte  As String       '30        자동생성 order생성 일자 (자동생성된 order data 가 저장된다.)
    IcmPreSts  As String       '31        퇴원Status(PRETBL 심사대기중-1,심사중-2,심사완료-3,수납대기-4,수납완료-5)
    IcmPreDtm  As String       '32        퇴원예정시간
    IcmOdrDtm  As String       '33        퇴원오더시간
    IcmCfmYon  As String       '34        퇴원완료,예정구분(완료=OT,예정=OR)
    IcmIspBak  As String       '35        IspInf를 IspLev로 BackUp 받았는지 여부...
    IcmSimDtm  As String       '36        사전심사한 최종일자
    '------------------------------------------------------------------------대구성삼추가
    IcmPedOcm  As String * 10  '37        산모차트와 연결된 아가의 내원번호
    IcmPedDtm  As String       '38        분만일자..
    '------------------------------------------------------------------------대구성삼추가
End Type
                        
    '--------------------------------------------------------------------
    '2) 병실이동 정보
    '--------------------------------------------------------------------
    ' ItrInfDtmCht : K-2,D-2
Type ItrInfRec
    ItrOcmNum  As String * 10  'ItrInfKey 입원번호
    ItrStrDtm  As String       'ItrInfKey 시작일시
    ItrEndDtm  As String       ' 3        종료일시
    ItrChtNum  As String * 8   ' 4        챠트번호
    ItrDepCod  As String       ' 5        진료과
    ItrDtrCod  As String       ' 6        주치의
    ItrNssCod  As String       ' 7        병동
    ItrRomCod  As String       ' 8        병실
    ItrBedCod  As String       ' 9        병상
    ItrBedGrd  As String       ' 10       병실등급
    ItrSpcYon  As String       ' 11       특진구분
    ItrWhyCod  As String       ' 12       이동사유
    ItrUpdDtm  As String       ' 13       수정일시
    ItrUidCod  As String       ' 14       수정담당자
End Type
    
    '--------------------------------------------------------------------
    '2-1) 병실이동 정보 History
    '--------------------------------------------------------------------
Type ItrHstRec
    ItrOcmNum  As String * 10  'ItrHstKey 입원번호
    ItrStrDtm  As String       'ItrHstKey 시작일시
    ItrDelDtm  As String       'ItrHstKey 삭제일시
    ItrSttFlg  As String       'ItrHstKey 삭제및 수정여부
    ItrEndDtm  As String       ' 4        종료일시
    ItrChtNum  As String * 8   ' 5        챠트번호
    ItrDepCod  As String       ' 6        진료과
    ItrDtrCod  As String       ' 7        주치의
    ItrNssCod  As String       ' 8        병동
    ItrRomCod  As String       ' 9        병실
    ItrBedCod  As String       ' 10       병상
    ItrBedGrd  As String       ' 11       병실등급
    ItrSpcYon  As String       ' 12       특진구분
    ItrWhyCod  As String       ' 13       이동사유
    ItrUpdDtm  As String       ' 14       수정일시
    ItrUidCod  As String       ' 15       수정담당자
End Type
    
    '--------------------------------------------------------------------
    '3) 식이 변경 정보
    '--------------------------------------------------------------------
Type IdiInfRec
    IdiOcmNum  As String * 10  'IdiInfKey 입원번호
    IdiFrmDte  As String       'IdiInfKey 시작일
    IdiFrmTyp  As String       'IdiInfKey 구분(끼니)
    IdiFeeCod  As String       '          급식코드
    IdiCalTyp  As String       '          급식계산구분
    IdiEndDte  As String       '          급식종료일자
    IdiEndTyp  As String       '          급식종료유형
    IdiNssCod  As String       '          병동
    IdiRomCod  As String       '          병실
    IdiBedCod  As String       '          병상
    IdiAddCod  As String       '          추가코드
    IdiDepCod  As String       '          진료과
    IdiDtrCod  As String       '          주치의
    IdiWhyCod  As String       '          이동사유
    IdiUpdDte  As String       '          수정일시
    IdiUidCod  As String       '          담당자
End Type
    
    '--------------------------------------------------------------------
    '4) 입원 처방 정보 IspInf
    '--------------------------------------------------------------------
Type IspInfRec
    
    IspOcmNum  As String * 10  'IspInfKey 입원번호
    IspOdrNum  As String * 4   'IspInfKey 처방번호
    IspOdrSeq  As String * 5   'IspInfKey 처방순서
    IspOdrCod  As String       ' 1        처방코드
    IspOdrTyp  As String       ' 2        처방유형
    IspSlpDep  As String       ' 3        처방전달부서
    IspSlpCod  As String       ' 4        처방전달범주
    IspItmCod  As String       ' 5        처방항목코드
    IspFeeCod  As String       ' 6        수가코드
    IspOdrPrc  As String       ' 7        단가1
    
    IspOdrSib  As String       ' 8        'MIX.....투약용으로 사용하기로 함....2001/02/01
    IspOdrQty  As String       ' 9        투여량
    IspOdrDay  As String       '10        일수
    IspOdrTms  As String       '11        횟수
    IspInsYon  As String       '12        급여/비급여구분
    IspInsCod  As String       '13        보험유형
    IspInsSeq  As String * 2   '14        보험유형순서
    IspDgsEtc  As String       '15        기타 (주종, 수술,처치...)
    IspOdrDnh  As String       '16        주야공휴
    IspOprDtm  As String       '17        처치,수술,마취.. 실시시간
    
    IspDepCod  As String       '18        진료과
    IspOdrDtm  As String       '19        처방일시
    IspOdrStt  As String       '20        처방진행상태
    IspStkStt  As String       '21        재고진행상태
    IspEmgYon  As String       '22        응급여부
    IspSpcYon  As String       '23        특진구분
    IspCmpSym  As String       '24        약품분류기호
    IspUsgCod  As String       '25        용법
    IspMthCod  As String       '26        행위코드
    IspSpmCod  As String       '27        검체코드
    
    IspIncCod  As String * 2   '28        수입원
    IspSotCod  As String       '29        처방항목
    IspSlpAmt  As String       '30        계산금액
    IspAddCod  As String       '31        가산코드
    IspDupYon  As String       '32        이중보험여부(1,2,3,4...)
    IspDscMed  As String       '33        퇴원약구분
    IspAftYon  As String       '34        퇴원후처방여부
    IspAddAmt  As String       '35        가산금액
    IspPreDtm  As String       '36        예약일시
    
    IspEntDtm  As String       '37        입력일시
    
    IspUidCod  As String       '38        입력담당자
    IspCncDtm  As String       '39        취소일시
    IspCncUid  As String       '40        취소담당자
    IspSplYon  As String       '41        특기사항
    IspSplCmt  As String       '42        특기사항
    IspChkStt  As String       '43        전달여부  "-1" : 취소처방을 지원부서에서 확인,"0" : 처방을 지원부서에서 확인, "1"이상 : 지원부서에서 시행할 일수
    IspDgsRol  As String       '44        Right/Left
    IspCvtYon  As String       '45        입원전환여부
    IspMdcNum  As String       '46        투약번호
    IspCanMdc  As String       '47        최소시의 투약번호
    
    IspStgCod  As String       '48        수탁검사시 의뢰기관 코드
    IspBasUnt  As String       '49        의사사용량
    IspMntUsg  As String       '50        정신과용법
    IspXryPtb  As String       '51        Xray Portable
    IspPreStt  As String       '52        선수납Sts ("", 선수납전표수납전= P, 선수납전표수납후=Irc영수증번호)
    IspConYon  As String       '53        Consult여부
    
    IspUidPrt  As String       '54        입력담당부서
    
    IspCncPrt  As String       '55        취소담당부서
    IspImgYon  As String       '56        수입조정여부
    IspCanNum  As String       '57        취소Order Number
    
    IspCanSeq  As String       '58        취소Order Seq
    IspAstCod  As String       '59        보조분류코드
    IspMixNum  As String       '60        MixNum
    IspDenRgn  As String       '61        코드별 치식 입력 02.03.21 sebal.
    
    IspEodNum  As String        '62       EodInf order Number
    IspEodSeq  As String        '63       EodInf Order Sequence
    
    IspIctNum  As String        '64       IctInf order Number
    IspIctSeq  As String        '65       IctInf Order Sequence

End Type
    '--------------------------------------------------------------------
    ' 입원 계산 정보 IctInf(원무과 계산용)
    '--------------------------------------------------------------------
Type IctInfRec
    IspOcmNum  As String * 10  'IspInfKey 입원번호
    IspOdrNum  As String * 4   'IspInfKey 처방번호
    IspOdrSeq  As String * 5   'IspInfKey 처방순서
    IspOdrCod  As String       ' 1        처방코드
    IspOdrTyp  As String       ' 2        처방유형
    IspSlpDep  As String       ' 3        처방전달부서
    IspSlpCod  As String       ' 4        처방전달범주
    IspItmCod  As String       ' 5        처방항목코드
    IspFeeCod  As String       ' 6        수가코드
    IspOdrPrc  As String       ' 7        단가1
    
    IspOdrSib  As String       ' 8        'MIX.....투약용으로 사용하기로 함....2001/02/01
    IspOdrQty  As String       ' 9        투여량
    IspOdrDay  As String       '10        일수
    IspOdrTms  As String       '11        횟수
    IspInsYon  As String       '12        급여/비급여구분
    IspInsCod  As String       '13        보험유형
    IspInsSeq  As String * 2   '14        보험유형순서
    IspDgsEtc  As String       '15        기타 (주종, 수술,처치...)
    IspOdrDnh  As String       '16        주야공휴
    IspOprDtm  As String       '17        처치,수술,마취.. 실시시간
    
    IspDepCod  As String       '18        진료과
    IspOdrDtm  As String       '19        처방일시
    IspOdrStt  As String       '20        처방진행상태
    IspStkStt  As String       '21        재고진행상태
    IspEmgYon  As String       '22        응급여부
    IspSpcYon  As String       '23        특진구분
    IspCmpSym  As String       '24        약품분류기호
    IspUsgCod  As String       '25        용법
    IspMthCod  As String       '26        행위코드
    IspSpmCod  As String       '27        검체코드
    
    IspIncCod  As String * 2   '28        수입원
    IspSotCod  As String       '29        처방항목
    IspSlpAmt  As String       '30        계산금액
    IspAddCod  As String       '31        가산코드
    IspDupYon  As String       '32        이중보험여부(1,2,3,4...)
    IspDscMed  As String       '33        퇴원약구분
    IspAftYon  As String       '34        퇴원후처방여부
    IspAddAmt  As String       '35        가산금액
    IspPreDtm  As String       '36        예약일시
    
    IspEntDtm  As String       '37        입력일시
    
    IspUidCod  As String       '38        입력담당자
    IspCncDtm  As String       '39        취소일시
    IspCncUid  As String       '40        취소담당자
    IspSplYon  As String       '41        특기사항
    IspSplCmt  As String       '42        특기사항
    IspChkStt  As String       '43        전달여부  "-1" : 취소처방을 지원부서에서 확인,"0" : 처방을 지원부서에서 확인, "1"이상 : 지원부서에서 시행할 일수
    IspDgsRol  As String       '44        Right/Left
    IspCvtYon  As String       '45        입원전환여부
    IspMdcNum  As String       '46        투약번호
    IspCanMdc  As String       '47        최소시의 투약번호
    
    IspStgCod  As String       '48        수탁검사시 의뢰기관 코드
    IspBasUnt  As String       '49        의사사용량
    IspMntUsg  As String       '50        정신과용법
    IspXryPtb  As String       '51        Xray Portable
    IspPreStt  As String       '52        선수납Sts ("", 선수납전표수납전= P, 선수납전표수납후=Irc영수증번호)
    IspConYon  As String       '53        Consult여부
    
    IspUidPrt  As String       '54        입력담당부서
    
    IspCncPrt  As String       '55        취소담당부서
    IspImgYon  As String       '56        수입조정여부
    IspCanNum  As String       '57        취소Order Number
    
    IspCanSeq  As String       '58        취소Order Seq
    IspAstCod  As String       '59        보조분류코드
    IspMixNum  As String       '60        MixNum
    IspDenRgn  As String       '61        코드별 치식 입력 02.03.21 sebal.
    
    IspEodNum  As String        '62       IspInf order Number
    IspEodSeq  As String        '63       IspInf Order Sequence

End Type

    
    '--------------------------------------------------------------------
    '5) 입원 계산정보 IrpInf
    '--------------------------------------------------------------------
Type IrpInfRec
    IrpOcmNum  As String * 10  'IrpInfKey 입원번호
    IrpOcmSeq  As String * 2   'IrpInfKey 내원변경순서
    IrpDupSeq  As String * 2   'IrpInfKey 이중보험순서
    IrpDepCod  As String       'IrpInfKey 진료과
    IrpChtNum  As String * 8   '1         챠트번호
    IrpDtrCod  As String       '2         주치의
    IrpInsCod  As String       '3         보험유형
    IrpInsSeq  As String * 2   '4         보험유형순서
    IrpTotAmt  As String       '5         진료비총액
    IrpCorAmt  As String       '6         조합청구액
    IrpNonAmt  As String       '7         비급여총액          NonOwn
    IrpOwnAmt  As String       '8         급여 본인부담액     InsOwn
    IrpTotOwn  As String       '9         본인부담총액
    IrpInsAmt  As String       '10        급여총액            InsTot
    IrpSpcAmt  As String       '11        특진료
    IrpAskAmt  As String       '12        환자청구액
    IrpDisAmt  As String       '13        할인액
    IrpFutAmt  As String       '14        후불액
    IrpOldAmt  As String       '15        기수납액
    IrpNewAmt  As String       '16        수납액
    IrpGrnAmt  As String       '17        보증금중 계산에 적용된 금액
    IrpRcpNum  As String * 10  '18        영수증번호
    IrpOldNum  As String * 10  '19        이전영수증번호
    IrpCalDte  As String       '20        중간계산일
    IrpUpdDtm  As String       '21        계산일시(system time)
    IrpUidCod  As String       '22        담당자코드
    IrpOvrAmt  As String       '23        과체납액
    IrpRcpFlg  As String       '24        계산Flag 중간계산:MIDCAL,중간수납:MIDRCP,퇴원계산:DISCAL,퇴원수납:DISRCP,보증금:GRNRCP,선수납:PRERCP,미수입금:FUTRCP
    IrpDimAmt  As String       '25        감면액
    IrpNonIns  As String       '26        보험100% 여부    '^_^ 980427
    IrpCarFut  As String       '27        카드미수
    IrpFodAmt  As String       '28        기카드미수
'--------------------------------------> 추가
    IrpNinAmt  As String       '30        전액본인부담액 20040102..HTS..add
'--------------------------------------> 추가
End Type
    
    '--------------------------------------------------------------------
    '6) 입원 계산 변경정보 IhpInf
    '--------------------------------------------------------------------
Type IhtInfRec
    IhtRcpNum  As String * 10  'IhtInfKey 영수증번호
    IhtOcmNum  As String * 10  '1         입원번호
    IhtOcmSeq  As String * 2   '2         내원변경순서
    IhtDupSeq  As String * 2   '3         이중보험순서
    IhtDepcod  As String       '4         진료과
    IhtChtNum  As String * 8   '5         챠트번호
    IhtDtrCod  As String       '6         주치의
    IhtInsCod  As String       '7         보험유형
    IhtInsSeq  As String * 2   '8         보험유형순서
    IhtTotAmt  As String       '9         진료비총액
    IhtCorAmt  As String       '10        조합청구액
    IhtNonAmt  As String       '11        비급여총액          NonOwn
    IhtOwnAmt  As String       '12        급여 본인부담액     InsOwn
    IhtTotOwn  As String       '13        본인부담총액
    IhtInsAmt  As String       '14        급여총액            InsTot
    IhtSpcAmt  As String       '15        특진료
    IhtAskAmt  As String       '16        환자청구액
    IhtDisAmt  As String       '17        할인액
    IhtFutAmt  As String       '18        후불액
    IhtOldAmt  As String       '19        기수납액
    IhtNewAmt  As String       '20        수납액
    IhtGrnAmt  As String       '21        보증금중 계산에 적용된 금액
    IhtRcpCur  As String * 10  '22        영수증번호
    IhtOldNum  As String * 10  '23        이전영수증번호
    IhtCalDte  As String       '24        중간계산
    IhtUpdDtm  As String       '25        계산일시(system time)
    IhtUidCod  As String       '26        담당자코드
    IhtOvrAmt  As String       '27        과체납액
    IhtRcpFlg  As String       '28        계산Flag 중간계산:MIDCAL,중간수납:MIDRCP,퇴원계산:DISCAL,퇴원수납:DISRCP,보증금:GRNRCP,선수납:PRERCP,미수입금:FUTRCP
    IhtDimAmt  As String       '29        감면액
    IhtNonIns  As String       '30        보험 100%여부
    IhtCarFut  As String       '31        카드미수
    IhtFodAmt  As String       '32        기카드미수
'--------------------------------------> 추가
    IhtNinAmt  As String       '34        전액본인부담액 20040102..HTS.. add
'--------------------------------------> 추가
End Type
    
    '**************************************
    '   입원 Daily Summary
    '**************************************
Type IdaInfRec
    IdaOcmNum As String * 10    'IdaInfKey 내원번호
    IdaDupSeq As String * 2     'IdaInfKey 이중보험순서
    IdaOdrDte As String         'IdaInfKey 처방일자
    IdaDepCod As String         'IdaInfKey 진료과목
    IdaItmCod As String         'IdaInfKey 항목코드
    IdaAstCod As String         'IdaInfKey 보조분류
    IdaChtNum As String * 8     '          챠트번호
    IdaInsMat As String         '          급여재료
    IdaInsAct As String         '          급여행위
    IdaNonMat As String         '          비급여재료
    IdaNonAct As String         '          비급여행위
    IdaSpcAmt As String         '          특진료
    IdaUpdDtm As String         '          수정일시
    IdaUidCod As String         '          User ID
End Type
    
    '--------------------------------------------------------------------
    '7) 입원 계산 상세정보 IdlInf
    '--------------------------------------------------------------------
Type IdlInfRec
    IdlRcpNum  As String * 10  'IdlInfKey 영수증번호
    IdlIncCod  As String * 2   'IdlInfKey 수입원
    IdlChtNum  As String * 8   '          챠트번호
    IdlInsCod  As String       '          보험유형
    IdlInsSeq  As String * 2   '          유형순서
    IdlDepCod  As String       '          진료과목
    IdlDtrCod  As String       '          주치의
    IdlInsAct  As String       '          급여행위
    IdlInsMat  As String       '          급여재료
    IdlNonAct  As String       '          비급여행위
    IdlNonMat  As String       '          비급여재료
    IdlInsAmt  As String       '          급여총액
    IdlNonAmt  As String       '          비급여액
    IdlInsOwn  As String       '          급여 본인부담액
    IdlTotOwn  As String       '          본인부담액
    IdlSpcAmt  As String       '          특진료
'--------------------------------------> 추가
    '20040102..HTS..add
    IdlNinAct  As String       '          전액본인부담행위
    IdlNinMat  As String       '          전액본인부담재료
    IdlNinAmt  As String       '          전액본인부담
'--------------------------------------> 추가
End Type
    
    '--------------------------------------------------------------------
    '8) 입원 환자 Logging 정보 IloInf
    '--------------------------------------------------------------------
Type IloInfRec
    IloUpdDtm  As String       'IloInfKey 처리일시
    IloSavSeq  As String * 2   'IloInfKey 저장순서
    IloOcmNum  As String * 10  '          내원번호
    IloChtNum  As String * 8   '          챠트번호
    IloFrmIns  As String       '          보험유형(From)
    IloFrmDep  As String       '          진료과목(From)
    IloFrmDtr  As String       '          주치의(From)
    IloFrmNss  As String       '          병동(From)
    IloFrmRom  As String       '          병실(From)
    IloFrmBed  As String       '          병상(From)
    IloToIns   As String       '          보험유형(to)
    IloToDep   As String       '          진료과목(to)
    IloToDtr   As String       '          주치의(to)
    IloToNss   As String       '          병동(to)
    IloToRom   As String       '          병실(to)
    IloToBed   As String       '          병상(to)
    IloFunCod  As String       '          업무기능
End Type
    
    '--------------------------------------------------------------------
    '9) 입원 수납 정보 (보증금,중간입금,퇴원입금) IrcInf
    '--------------------------------------------------------------------
Type IrcInfRec
    IrcOcmNum  As String * 10  'IrcInfKey 내원번호
    IrcOcmSeq  As String * 2   'IrcInfKey 내원변경순서
    IrcDupSeq  As String * 2   'IrcInfKey 이중보험순서
    IrcRcpNum  As String * 10  'IrcInfKey 영수증번호
    IrcChtNum  As String * 8   '  1       챠트번호
    IrcIrpNum  As String * 10  '  2       연결 계산서 번호
    IrcRcpTyp  As String       '  3       수납구분(보증금,중간입금,퇴원입금,보증금대체,중간금대체)
    IrcDepCod  As String       '  4       진료과
    IrcRetYon  As String       '  5       반환구분
    IrcRcpAmt  As String       '  6       수납액
    IrcRcpDtm  As String       '  7       수납일시
    IrcRetAmt  As String       '  8       반환금
    IrcRetDtm  As String       '  9       반환일시
    IrcUidCod  As String       ' 10       담당자
    IrcRetUid  As String       ' 11       반환담당자
    IrcRelCod  As String       ' 12       환자와의 관계
    IrcManNam  As String       ' 13       입금및 반환자성명
    IrcPreCas  As String       ' 14       외래이체,선수납(Not = "", 외래이체는 O, 선수납은 P)
    IrcInsCod  As String       ' 15       보험유형
    IrcDtlNum  As String       ' 16       연결계산서의 과별 영수증 번호
    IrcCarFut  As String       ' 17       카드미수
    IrcCarRet  As String       ' 17       카드미수반환금
End Type
    
    '--------------------------------------------------------------------
    '10) 입원 보험유형 정보 IisInf
    '--------------------------------------------------------------------
Type IisInfRec
    IisOcmNum  As String * 10  'IisInfKey 내원번호
    IisOcmSeq  As String * 2   'IisInfKey 내원변경순서
    IisDupSeq  As String * 2   'IisInfKey 이중보험순서
    IisChtNum  As String * 8   '          챠트번호
    IisInsCod  As String       '          보험유형
    IisInsSeq  As String * 2   '          보험유형순서
    IisSpcYon  As String       '          특진구분
    IisArtYon  As String       '          인공신장구분
    IisAdpDte  As String       '          적용개시일
    IisExpDte  As String       '          적용종료일
    IisAcpDay  As String       '          입원일수
    IisIcuDay  As String       '          입원일수(ICU)
    IisRcpYon  As String       '          계산수납유무 (0.없음, 1.계산 2.수납)
    IisRcpDtm  As String       '          수납일시
    IisDepCod  As String       '          과목코드
    IisDtrCod  As String       '          의사코드
    IisNonIns  As String       '          보험100% 여부
    IisBilYon  As String       '          청구여부(청구한 년월로 저장한다, ex) "199705")
    IisDrgYon  As String       '          DRG 청구여부
    IisUidCod  As String       '          User ID
    IisDrgCod  As String       '          DRG 코드
    IisDrgDay  As String       '          DRG 적용일수
    IisUpdDtm  As String
    IisCstYon  As String       '          한방/양방 협진 여부
    IisLmtAmt  As String       '          자보책임/자손(한도금액)
End Type
    
    '----------------------------------------------------------------------
    '10) 입원 보험유형 History IisHst
    '--------------------------------------------------------------------
Type IisHstRec
    IisOcmNum  As String * 10  'IisHstKey 내원번호
    IisOcmSeq  As String * 2   'IisHstKey 내원변경순서
    IisDupSeq  As String * 2   'IisHstKey 이중보험순서
    IisDelDtm  As String       'IisHstKey 삭제일시
    IisChtNum  As String * 8   '          챠트번호
    IisInsCod  As String       '          보험유형
    IisInsSeq  As String * 2   '          보험유형순서
    IisSpcYon  As String       '          특진구분
    IisArtYon  As String       '          인공신장구분
    IisAdpDte  As String       '          적용개시일
    IisExpDte  As String       '          적용종료일
    IisAcpDay  As String       '          입원일수
    IisIcuDay  As String       '          입원일수(ICU)
    IisRcpYon  As String       '          계산수납유무 (0.없음, 1.계산 2.수납)
    IisRcpDtm  As String       '          수납일시
    IisDepCod  As String       '          과목코드
    IisDtrCod  As String       '          의사코드
    IisNonIns  As String       '          보험100% 여부
    IisBilYon  As String       '          청구여부(청구한 년월로 저장한다, ex) "199705")
    IisDrgYon  As String       '          Drg여부
    IisUidCod  As String       '          User ID
    IisDrgCod  As String       '          DRG 코드
    IisDrgDay  As String       '          DRG 적용일수
    IisUpdDtm  As String
    IisActTyp  As String
    IisCstYon  As String       '          한방/양방 협진 여부
    IisLmtAmt  As String       '          자보책임/자손(한도금액)
    
End Type
    
    '**************************************
    '   입원 대기자 명부
    '**************************************
Type IwlInfRec
    IwlTypCod   As String     'IwlInfKey  1 구분 Key : GM,GF,SM,SF,OM,OF,DM,DF,MM,MF,PM,PF,EM,EF,AL,AM,CM,CF
    IwlAcpNum   As String * 4 'IwlInfKey  2 접수번호
    IwlChtNum   As String * 8 '           3 챠트번호
    IwlAcpDte   As String     '           4 접수일자
    IwlReqNam   As String     '           5 신청인명
    IwlActDte   As String     '           6 통보시행일
    IwlEntDte   As String     '           7 입원일자
    IwlNssCod   As String     '           8 병동
    IwlFstPhn   As String     '           9 전화번호1
    IwlSndPhn   As String     '           10 전화번호2
    IwlSplCmt   As String     '           11 비고
    IwlStrUid   As String     '           12 입력담당자
    IwlStrDtm   As String     '           13 입력일시
    IwlUpdUid   As String     '           14 수정담당자
    IwlUpdDtm   As String     '           15 수정일시
End Type
    
    
    '**************************************
    '   퇴원 기록지
    '**************************************
Type LhrInfRec
    LhrOcmNum   As String * 10  'LhrInfKey   입원번호
    LhrLevDte   As String       '           1 퇴원일자
    LhrEduYer   As String       '           2 교육년수
    LhrJobCod   As String       '           3 입원전직업
    LhrJobCmt   As String       '           4 입원전직업: 기타
    LhrEcnStt   As String       '           5 경제상태
    LhrMrgStt   As String       '           6 결혼상태
    LhrRlgCod   As String       '           7 종교
    LhrFstAge   As String       '           8 초발연령
    LhrWrsDte   As String       '           9 최근악화시기
    LhrInhTyp   As String       '           10 입원형태
    LhrOthCnt   As String       '           11 타병원 입원회수
    LhrManPbm   As String       '           12 입원시 주문제
    LhrFmlHst   As String       '           13 가족의 정신병력
    LhrFstFml   As String       '           14 Y=>환자와의 관계
    LhrFstDgn   As String       '           15 Y=>진단
    LhrFstCmt   As String       '           16 Y=>진단: 기타
    LhrSndFml   As String       '           17 Y=>환자와의 관계
    LhrSndDgn   As String       '           18 Y=>진단
    LhrSndCmt   As String       '           19 Y=>진단: 기타
    LhrTrdFml   As String       '           20 Y=>환자와의 관계
    LhrTrdDgn   As String       '           21 Y=>진단
    LhrTrdCmt   As String       '           22 Y=>진단: 기타
    LhrFthFml   As String       '           23 Y=>환자와의 관계
    LhrFthDgn   As String       '           24 Y=>진단
    LhrFthCmt   As String       '           25 Y=>진단: 기타
    LhrClnDgn   As String       '           26 진단기록(임상진단)
    LhrTrbDgn   As String       '           27 진단기록(발달장애,성격장애)
    LhrPhyDgn   As String       '           28 진단기록(신체질환)
    LhrLmtCls   As String       '           29 진단기록(상태정도 분리)
    LhrFunInh   As String       '           30 진단기록(전반적인 기능 척도-입원)
    LhrFunLeh   As String       '           31 진단기록(전반적인 기능 척도-퇴원)
    LhrFunBin   As String       '           32 진단기록(전반적인 기능 척도-입원전 1년)
    LhrEstAct   As String       '           33 입원동안 치료및 검사(EST실시)
    LhrEegLab   As String       '           34 입원동안 치료및 검사(EEG검사)
    LhrPsyLab   As String       '           35 입원동안 치료및 검사(심리검사)
    LhrLevJud   As String       '           36 퇴원경위
    LhrTrtRst   As String       '           37 치료결과 평가
    LhrLevTrt   As String       '           38 퇴원후 치료권유
    LhrSpcDtr   As String       '           39 지정의
    LhrDtrCod   As String       '           40 주치의
End Type
    
    '-----------
    ' 삭제 화일  -> 정신병원
    '-----------
Type DctInfRec
    DctChtNum  As String * 8   'PspInfKey 1  챠트번호
    DctUpdTim  As String       '          2  수정일자
    DctUidCod  As String       '          3  담당자
End Type
    
    '********************************************************************
    ' Mail Box 기본정보
    '********************************************************************
    '--------------------------------------------------------------------
    ' Mail Editor 정보
    '--------------------------------------------------------------------
Type MalInfRec
    MalRcvUid As String     ' MalInfKey Mail받는이 Code
    MalSndDtm As String     ' MalInfKey Mail보내는 시각
    MalSndUid As String     ' MalInfKey Mail보내는이 Code
    MalCfmYon As String     '           Mail의 확인 여부
    MalSndSts As String     '           Mail보내는 형태( "B": 전체, "D":부서 "I":개인)
    MalMsgDtl As String     '           Mail의 내용
    MalMsgSbj As String     '           Mail의 제목
    MalApdFle As String     '첨부화일
End Type

Type MalGrpRec
    MalGrpUid As String     ' MalGrpKey Mail그룹user
    MalGrpCod As String     ' MalGrpKey Mail그룹코드
    MalGrpNam As String     ' 그룹코드명칭
    
End Type

Type MalDtlRec
    MalGrpUid As String     ' MalDtlKey Mail그룹user
    MalGrpCod As String     ' MalDtlKey Mail그룹코드
    MalDtlUid As String     ' MalDtlKey Mail받는 user
    
End Type
    
    '--------------------------------------------------------------------
    ' 개인미수 기본 정보    (FutInf)
    
    '--------------------------------------------------------------------
Type FutInfRec
    FutChtNum  As String * 8   ' K1     1 챠트번호
    FutCurSts  As String       '        2 미수상태      'P:미납, O:완납
    FutFutAmt  As String       '        3 총발생액
    FutPayAmt  As String       '        4 총납입액
    FutDisAmt  As String       '        5 총상각액
    FutRemAmt  As String       '        6 총미수액
    FutEmpNum  As String       '        7 직원번호
    FutStrDte  As String       '        8 미수발생시작일
    FutEndDte  As String       '        9 최종미수발생일
    FutFutRsn  As String       '       10 미수사유
    FutExpDte  As String       '       11 미수입금예정일
End Type
    
    '--------------------------------------------------------------------
    ' 개인미수 발생 정보    (ForInf)
    'Primary Key  ForInf        (K-1,K-2,K-3,K-4)
    'Index Key    ForInfPrcRcp  (D-14, D-11)
    '--------------------------------------------------------------------
Type ForInfRec
    ForChtNum  As String * 8   ' K1     1 챠트번호
    ForFutTyp  As String       ' K2     2 미수구분          'O:발생, S:소송 B:청구미수
    ForOcmNum  As String * 10  ' K3     3 내원번호
    ForRvnTyp  As String       ' K4     4 M'접수, O'수납, I'입원
    ForOcmSeq  As String * 2   ' K5     5 내원변경순서"  "
    ForDupSeq  As String * 2   ' K6     6 이중보험순서"  "
    ForFutSts  As String       ' D1     7 상태              'O:완납, P:미납
    ForPatTyp  As String       ' D2     8 환자구분          'I:입원, O:외래
    ForInsCod  As String       ' D3     9 보험유형
    ForDepCod  As String       ' D4    10 진료과
    ForSerNum  As String       ' D5    11 일련번호
    ForOcrAmt  As String       ' D6    12 발생액
    ForDisAmt  As String       ' D7    13 상각액
    ForPayAmt  As String       ' D8    14 수납액
    ForRemAmt  As String       ' D9    15 미수액
    ForCorAmt  As String       ' D10   16 청구액
    ForRcpNum  As String * 10  ' D11   17 영수증번호
    ForOldNum  As String * 10  ' D12   18 이전영수증번호
    ForOcrRsn  As String       ' D13   19 발생사유
    ForPrcDtm  As String       ' D14   20 처리일시
    ForUidCod  As String       ' D15   21 담당자
    ForDisYon  As String       ' D16   22 상각사유
End Type
    
    '--------------------------------------------------------------------
    ' 개인미수 발생 변경 정보    (FhtInf)
    'Primary Key   FhtInf(K-1)
    'IndexKey-1    FhtInfChtFutOcmRvn(D-1, D-2, D-3, D-4)
    'IndexKey-2    FhtInfPrcRcp  (D-18, D-15)
    '--------------------------------------------------------------------
Type FhtInfRec
    FhtRcpNum  As String * 10  ' K1     1 영수증번호
    FhtFutTyp  As String       ' K2     2 미수구분      'O:발생, S:소송
    FhtChtNum  As String * 8   '  1     3 챠트번호
    FhtFutOld  As String       '  2     4 미수구분      'O:발생, S:소송
    FhtOcmNum  As String * 10  '  3     5 내원번호
    FhtRvnTyp  As String       '  4     6 접수,수납
    FhtOcmSeq  As String * 2   '  5     7 내원변경순서
    FhtDupSeq  As String * 2   '  6     8 이중보험순서
    FhtFutSts  As String       '  7     9 상태          'O:완납, P:미납
    FhtPatTyp  As String       '  8    10 환자구분      'I:입원, O:외래
    FhtInsCod  As String       '  9    11 보험유형
    FhtDepCod  As String       ' 10    12 진료과
    FhtSerNum  As String       ' 11    13 일련번호
    FhtOcrAmt  As String       ' 12    14 발생액
    FhtDisAmt  As String       ' 13    15 상각액
    FhtPayAmt  As String       ' 14    16 수납액
    FhtRemAmt  As String       ' 15    17 미수액
    FhtCorAmt  As String       ' 16    18 청구액
    FhtRcpOld  As String * 10  ' 17    19 영수증번호
    FhtOldNum  As String * 10  ' 18    20 이전영수증번호
    FhtOcrRsn  As String       ' 19    21 발생사유
    FhtPrcDtm  As String       ' 20    22 처리일시
    FhtUidCod  As String       ' 21    23 담당자
    FhtDisYon  As String       ' 22    24 상각사유
End Type
    
    '--------------------------------------------------------------------
    ' 개인미수 입금 정보    (FpaInf)
    '--------------------------------------------------------------------
Type FpaInfRec
    FpaChtNum  As String       ' K1     1 챠트번호
    FpaPaySeq  As String       ' K2     2 수납순서
    FpaPayAmt  As String       '        3 입금액
    FpaRemDsc  As String       '        4 비고
    FpaPrcDtm  As String       '        5 처리일시
    FpaUidCod  As String       '        6 사용자코드
    FpaRcpNum  As String       '        7 입금된 영수증번호
End Type
    
    '*****************
    '   의무 기록지
    '*****************
Type ChtInfRec
    ChtNum         As String * 8   'ChtInfKey 챠트번호
    ChtDscCc       As String       '          챠트내용(C.C)
    ChtDscPhx      As String       '          챠트내용(PHX)
    ChtDscSpc      As String       '          챠트내용(특이사항)
    ChtOldDscCc    As String       '          이전챠트내용(C.C)
    ChtOldDscPhx   As String       '          이전챠트내용(PHX)
    ChtOldDscSpc   As String       '          이전챠트내용(특이사항)
    ChtUidCod      As String       '          입력담당자
    ChtEntDtm      As String       '          입력일시
End Type
    
    '**************************************
    '   예약 화일 (OCS write 원무 read)
    ' RsvInfDtrDtmStt D-7, D-1, D-2
    '**************************************
Type RsvInfRec
    RsvOcmNum  As String * 10  'RsvInfKey 내원번호
    RsvDtm     As String       '          예약일시
    RsvSts     As String       '          예약상태    ("OS" : OCS Write, "OR":수납 Write, "OC": 취소)
    RsvChtNum  As String * 8   '          챠트번호
    RsvDepCod  As String       '          진료과
    RsvUidCod  As String       '          입력담당자
    RsvChkYon  As String       '          예약처리Check 여부
    RsvDtrCod  As String       '          의사코드 Added by JES at 97/02/01 St. John
End Type
    
    
    '**************************************
    '   검사 예약
    '**************************************
Type RctInfRec
    RctCod      As String   'RctInfKey 검사코드
    RctDte      As String   'RctInfKey 검사일시
    RctTotCnt   As String   '          검사예약총수
    RctCurCnt   As String   '          현재 검사예약수
End Type
    
    '**************************************
    '   검사 예약2
    '**************************************
Type RcsInfRec
    RcsDte      As String       'RcsInfKey 검사일시
    RcsOcmNum   As String * 10  'RcsInfKey 내원번호
    RcsCod      As String       'RcsInfKey 검사코드
    RcsStt      As String       '          검사상태
    RcsSlpDep   As String       '          검사과목
End Type
    
    
    '**************************************
    '   전실 예정 IcrInfRec
    '**************************************
Type IcrInfRec
    IcrHopDtm   As String       ' IcrInfKey 전실희망일자
    IcrChtNum   As String * 8   ' IcrInfKey 챠트번호
    IcrOcmNum   As String * 10  '           내원번호
    IcrCurDep   As String       '           진료과
    IcrCurDtr   As String       '           담당의사
    IcrCurNss   As String       '           병동
    IcrCurRom   As String       '           병실
    IcrCurBed   As String       '           병상
    IcrCurGrd   As String       '           병실등급
    IcrHopDep   As String       '           진료과
    IcrHopDtr   As String       '           담당의사
    IcrHopNss   As String       '           병동
    IcrHopRom   As String       '           병실
    IcrHopBed   As String       '           병상
    IcrHopGrd   As String       '           병실등급
    IcrTrsDtm   As String       '           전실일시
    IcrNssYon   As String       '           병동확인
    IcrNssUid   As String       '           병동입력자
    IcrWonYon   As String       '           원무확인
    IcrWonUid   As String       '           원무입력자
End Type
    
    '**************************************
    '   내원 Reference 등록
    '**************************************
Type RefInfRec
    RefOcmNum   As String * 10  'CltInfKey  내원번호
    RefSplCmt   As String       '           특기사항
End Type
    
    '------------------------------------------------------
    ' 정액,정률 계산을 위한 개인별 영수액 저장   ZfmInf
    '       일별로 개인당 누적 화일
    '       급여에 관련된 금액만을 사용하여 저장
    '       접수비 후불에서는 ZfmAskNew = 0
    '       이외의 경우는 ZfmOwnAmt = ZfmAskNew
    '------------------------------------------------------
Type ZfmInfRec
    ZfmChtNum  As String * 8    'ZfmInfKey 챠트번호
    ZfmAdpDte  As String        'ZfmInfKey 적용 일자
    ZfmInsAmt  As String        '          급여 총액
    ZfmCorAmt  As String        '          조합 부담액
    ZfmOwnAmt  As String        '          급여 본인 부담액
    ZfmAskNew  As String        '          급여 본인중 환자 수납액
End Type
    
    '--------------------------------------------------------------------
    '   직원가족 등록 정보
    '--------------------------------------------------------------------
Type OffInfRec
    OffChtNum As String * 8     'OffInfKey 챠트번호
    OffRelTyp As String         '관계
    OffRelEmp As String         '연결직원코드
    OffEmpNam As String         '연결직원이름
    OffUidCod As String         '입력담당자
End Type
    
    '--------------------------------------------------------------------
    '   기타 수입
    '--------------------------------------------------------------------
Type EtcInfRec
    EtcOcmNum As String * 10        '1Key    내원번호
    EtcChtNum As String * 8         '2       Chart Number
    EtcPatNam As String             '3       환자성명
    EtcResNum As String             '4       주민번호
    EtcInsCod As String             '5       유형코드
    EtcInsSeq As String             '6       유형순서
    EtcDepCod As String             '6       과목코드
    EtcAssCod As String
    EtcOdrDtm As String             '7       처방일시
    EtcEntNam As String             '8       사업체명
    EtcSplCmt As String             '9       특기사항
    EtcUidCod As String             '10      담당자
    EtcGbnCod As String             '11      구분
    EtcSpcYon As String             '12      (국립정신용) 특진여부
    EtcTotAmt As String             '13      (국립정신용) 금액
    EtcMmsEtc As String             '14      M'의료수익 E'기타수익
    EtcGbnDtl As String
    EtcTelNum As String             '전화번호     2002.02.27 sebal
    EtcZipCod As String             '우편번호     2002.02.27 sebal
    EtcAddRes As String             '주소         2002.02.27 sebal
    EtcCalTel As String             '연락전화번호 2002.02.27 sebal
    EtcCalZip As String             '연락우편번호 2002.02.27 sebal
    EtcCalAdd As String             '연락주소     2002.02.27 sebal
    EtcJobNam As String             '직장명       2002.02.27 sebal
    EtcHndPhn As String             '핸드폰       2002.02.27 sebal
    EtcE_Mail As String             '이메일       2002.02.27 sebal
End Type
    
    '---------------------
    ' 의사 스케줄 (DUTY)
    '---------------------
Type DutInfRec
    DutDtrCod As String         ' DutInfKey 의사 User ID
    DutSttDtm As String         ' DutInfKey 의사 Off Duty 시작일시
    DutEndDtm As String         '           의사 Off Duty 끝일시
    DutOdrCmt As String         '           의사 comment
End Type
    
    '***************
    '가수금 예치대장
    '***************
Type GasInfRec
    GasAcpDtm As String             'Key    가수금발생일시
    GasChtNum As String * 8         '       Chart Number
    GasPatNam As String             '       환자성명
    GasResNum As String             '       주민번호
    GasInsCod As String             '       유형코드
    GasDepCod As String             '       과목코드
    GasInOut  As String             '       외래/입원구분
    GasSavNam As String             '       예치자성명
    GasSavTel As String             '       예치자연락처(전화,주소)
    GasGwnCod As String             '       환자와의 관계
    GasSavAmt As String             '       예치금액
    GasRcpNum As String             '       예치금영수증번호
    GasOldNum As String             '       이전영수증번호
    GasSplCmt As String             '       특기사항
    GasUidCod As String             '       담당자
    
End Type
    
    '--------------------------------------------------------------------
    ' 접수 대기 정보
    '--------------------------------------------------------------------
Type StbInfRec
    StbDepCod    As String      ' 접수과 코드  Key
    StbAcpDtm    As String      ' 접수 일시    Key
    StbOcmNum    As String * 10 ' 내원 번호    Key
    StbChtNum    As String * 8  ' 1     차트 번호
    StbPatNam    As String      ' 2     환자 명
    StbAcpStt    As String      ' 3     접수 상태  - 접수 (OA) - 예약 (OR) - 접수 차트 도착 (OI) - 접수 취소 (OC) - 접수 보류 (ON) - 재접수 (OH) - Transfer(OX)-Consult(OY)
    StbFlgStt    As String      ' 4     차트 불출 체크 Y - 불출, C-취소, D-오류챠트
    StbEmgYon    As String      ' 5     응급여부(Y/N)
    StbSplCmt    As String      ' 6     특기사항
    StbCstDep    As String      ' 7     Consult
    StbDtrCod    As String      ' 8     의사코드
    StbXryFlg    As String
    StbOrgDtm    As String      '10     '02.03.25 sebal 컨설트 보낸과의 최초 접수시간.
    StbCfmStt    As String      '11     수납대기상태 (OS : 대기, OT : 수납완료) - 환자의 지원부서 진행 현황등..
    StbAtoBar    As String      '12     BarCode자동출력 옵션(Y:출력, N:미출력)
    
    StbLabStt    As String      '13
    StbXryStt    As String      '14
    
End Type
    
    '--------------------------------------------------------------------
    ' 검사 접수 대기 정보
    '--------------------------------------------------------------------
Type LbqInfRec
    LbqSotCod    As String      '   항목 분류    Key
    LbqAcpDte    As String      '   접수 일자    Key
    LbqOcmNum    As String * 10 '   내원 번호    Key
    LbqChtNum    As String * 8  '1  차트 번호
    LbqPatNam    As String      '2  환자 명
    LbqDepCod    As String      '3  접수과 코드
    LbqWrdCod    As String      '4  병동 코드
    LbqRomCod    As String      '5  병실 코드
    LbqAcpStt    As String      '6  접수 상태  - 외래 (OA) - 입원 (IA) - 응급실 (EA) - 외래보류 (OH) - 입원보류 (IH) - 응급실보류 (IH)
    LbqCodCnt    As String      '7  발생 코드 수
    LbqEmgCnt    As String      '8  응급 코드 수
    LbqCasCnt    As String      '9  수납 코드 수
    LbqCanCnt    As String      '10 취소 코드 수
    LbqPreDay    As String      '11 선처방 일수 (일수가 7day면 6)
    LbqRsvYon    As String      '12 예약여부
    LbqAcpTms    As String      '13 접수시각
    LbqCfmDte    As String      '14 확인일자
    LbqOdrDte    As String      '15 처방일자    97/05/23 XRAY
    LbqCasDtm    As String      '16 수납일시
    LbqRstMak    As String      '17 대기자 진행상태
    LbqStdPat    As String      '18 대기자 Display상태
End Type
    
    '--------------------------------------------------------------------
    ' 검사 접수 정보    (LAbAdm)
    '--------------------------------------------------------------------
Type LbaInfRec
    LbaSotCod  As String       'LbaInfKey 항목분류
    LbaOcmNum  As String * 10  'LbaInfKey 내원번호
    LbaSeq     As String       'LbaInfKey 처방일자로 대치 5/28
    LbaSlpCod  As String       'LbaInfKey 슬립순서
    LbaChtNum  As String * 8   '        1 챠트번호
    LbaOdrDte  As String       '        2 처방일자
    LbaComStt  As String       '        3 외래(OA) /입원(IA) /응급(EA)
    LbaDepCod  As String       '        4 진료과
    LbaDtrCod  As String       '        5 진료의사
    LbaRomCod  As String       '        6 병동/병실
    LbaEmgYon  As String       '        7 응급여부
    LbaOdrStt  As String       '        8 처방 진행 상태 /OT(완료) /OC(취소) /OH(보류) /OP(진행중)
    LbaSpmNum  As String * 10  '        9 검체번호
    LbaAcpDtm  As String       '       10 접수일시
    LbaAcpUid  As String       '       11 접수담당
    LbaRstDte  As String       '       12 결과예정일
    LbaRptDte  As String       '       13 보고일자
    LbaRptUid  As String       '       14 보고담당
    LbaSlpDep  As String       '       15 전달부서
    LbaSplCmt  As String       '       16 특기사항
    LbaUidCod  As String       '       17 입력담당자
    LbaFstRed  As String       '       18 판독지소견입력(1) --> 97/01/10 추가
    LbaSndRed  As String       '       19 판독지소견입력(2)
    LbaTrdRed  As String       '       20 판독지소견입력(3)
    LbaForRed  As String       '       21 판독지소견입력(4)
    LbaFifRed  As String       '       22 판독지소견입력(5)
End Type
    
    '--------------------------------------------------------------------
    ' 차트반출 기타관리  - 1996.5.20 -김연수
    '--------------------------------------------------------------------
Type ChtIOInfRec
    ChtChtNum    As String * 8  ' Key   차트번호
    ChtDepCod    As String      ' Key   대출하려는 환자의 과 (통합일때는 "ALL",과별일때는 과목코드)
    ChtAcpDtm    As String      ' Key   대출일시
    ChtOrXray    As String      ' 1     차트=0 or X-Ray=1
    ChtOutDep    As String      ' 2     대출과
    ChtResDtm    As String      ' 3     반납일시
    ChtDlvNam    As String      ' 4     대출자성명
    ChtOutUid    As String      ' 5     반출자
    ChtRcvUid    As String      ' 6     반입자
    ChtOutGrp    As String      ' 7     대출자부서명
    ChtMemo      As String      ' 8     메모
    ChtMemDep    As String      ' 9     대출과(과별챠트와 통합챠트를 모두 적용하기 위한 필드)
                                '       과별일때는 ChtDepCod와 ChtMemDep가 같다.
End Type
    
    '**************************************
    '   Consult 등록
    '**************************************
Type CltInfRec
    CltOcmNum   As String * 10  'CltInfKey  내원번호
    CltAcpDte   As String       'CltInfKey  접수일자
    CltDepCod   As String       'CltInfKey  진료과
    CltUidCod   As String       '1          입력담당자
    CltComStt   As String       '2          외래(O) /입원(I) /응급(E)
End Type
    
    '**************************************
    '   검사 결과
    '**************************************
Type RstInfRec
    RstSotCod   As String       'RstInfKey  항목분류
    RstSpmNum   As String * 10  'RstInfKey  검체번호
    RstSpmSeq   As String * 2   'RstInfKey  검체순서
    RstLabCod   As String       '1          검사코드
    RstSeq      As String * 2   '2          슬립순서
    RstSlpCod   As String       '3          슬립코드
    RstSpmNam   As String       '4          슬립명칭
    RstOcmNum   As String * 10  '5          내원번호
    RstAcpDte   As String       '6          접수일자
    RstSplCmt   As String       '7          특기사항
    RstMzhMax   As String       '8          상한치
    RstMzhLow   As String       '9          하한치      '방사선소견작성여부
    RstMzhMnt   As String       '10         결과치      '방사선소견
    RstMzhUnt   As String       '11         결과단위    ,방사선부위
    RstJugCod   As String       '12         판정코드
    RstSlpDep   As String       '13         처방전달부서
    RstUidCod   As String       '14         수정담당
    RstUpdDtm   As String       '15         수정일시
    RstOdrCod   As String       '16         처방코드
    RstOdrNum   As String       '17         처방번호
End Type
    
    '--------------------------------------------------------------------
    ' 수술 Schedule 정보
    ' Index                     OprInfManDte = D-14,D-5
    '                           OprInfOprDte = D-3,D-5
    '                           OprInfRqtDte = D-16,D-5
    '                           OprInfOcmTms = D-5,D-7
    '--------------------------------------------------------------------
Type OprInfRec
    OprNum      As String * 10  '1Key    수술번호
    OprChtNum   As String * 8   '1       챠트번호
    OprOcmNum   As String * 10  '2       내원번호
    OprCod      As String       '3       수술코드
    OprGbnCod   As String       '4       구분코드("O","I","E")
    OprDte      As String       '5       수술일자
    OprTms      As String       '6       수술시각 (코드값 : 상세정보 사용)
    OprCfmYon   As String       '7       확인여부
    OprActYon   As String       '8       실행여부
    OprNarCod   As String       '9       마취코드
    OprIcdCod   As String       '10      상병코드
    OprDepCod   As String       '11      진료과
    OprNssRom   As String       '12      병동/병실
    OprEmgYon   As String       '13      응급여부
    OprManDtr   As String       '14      집도의
    OprDtrCod   As String       '15      입력의사
    OprRqtEqp   As String       '16      외부기구요청
    OprSplCmt   As String       '17      특기사항
    OprRomCod   As String       '18      수술방
    OprNarDtr   As String       '19      마취의
    OprNrsCod   As String       '20      수술간호사
    OprUpdDtm   As String       '21      수정일시
    OprUidCod   As String       '22      담당자코드(수정담당)
    OprOldCod   As String       '23      이전수술코드
    OprUseTms   As String       '24      수술예상시간
    OprPbsQty   As String       '25      환자 몸무게
    OprGazYon   As String       '26      OR Gauze 사용여부
    OprPanCtr   As String       '27      수술후 통증치료여부
    OprIcdNam   As String       '28      상병명
    OprStrTim   As String       '29      수술시작시간
    OprEndTim   As String       '30      수술종료시간
    OprPrtYon   As String       '31      출력여부
    OprBlood    As String       '32      혈액준비
    OprRstCxr   As String       '33      Chest x-ray
    OprRstEkg   As String       '34      EKG   판독결과
    OprNpoTms   As String       '35      EKG   판독결과
    OprCodNam   As String       '수술명칭
    
End Type
    
    '--------------------------------------------------------------------
    '   수술 Type 정보
    '--------------------------------------------------------------------
Type OprTypRec
    OprNum      As String * 10 'Key    수술번호
    OprCodTyp   As String      'Key    수술코드 Type(Position:PT,Equipments:EQ,Dtr:DR,Nrs:NR)
    OprTypCod   As String      'Key    수술Typ코드
    OprTypNam   As String      '       수술Typ명
End Type
    
    '--------------------------------------------------------------------
    '   퇴원 요약지 정보
    '--------------------------------------------------------------------
Type LslInfRec
    LslOcmNum   As String       'Key    내원번호
    LslChtNum   As String       '       챠트번호
    LslLevDtm   As String       '       퇴원일시
    LslManStt   As String       '       주증상
    LslFnlIcd   As String       '       최종진단명
    LslIcdSum   As String       '       병력요약
    LslAlgYon   As String       '       allergy유무
    LslAlgDtl   As String       '       allergy조치
    LslLevStt   As String       '       퇴원시조치
    LslDtrCod   As String       '       의사코드
End Type
    
    '--------------------------------------------------------------------
    '   I/O 등록
    '--------------------------------------------------------------------
Type InoInfRec
    InoOcmNum   As String       'Key    내원번호
    InoSrtDte   As String       '       시작일자
    InoEndDte   As String       '       종료일자
End Type
    
    '-------------------------------------------------------
    ' 공급실 관리 요청 정보
    '-------------------------------------------------------
Type CsdInfRec
    CsdAcpDte   As String       'Key    접수일시
    CsdDgsDEN   As String       'Key    "D"ay, "E"vning,"N"ight
    CsdUsgPrt   As String       'Key    요청부서(3W,4W,ER,OPN... UidMst의 담당부서별)
    CsdDgsTbl   As String       'Key    입력창( "1" : Routine, "2" : 추가, "3": 소독신청, "4": 물품대여)
    CsdSeq      As String * 3   'Key    요청순서
    CsdDepTyp   As String       '       입력부서(공급실:CSR, 병동:WRD, ...)
    CsdCsrcod   As String       '       분류코드(Key)
    CsdCsmYon   As String       '       소모품여부
    CsdRqtQty   As String       '       요청용량
    CsdOutQty   As String       '       불출용량
    CsdUidCod   As String       '       요청담당
    CsdUdpDtm   As String       '       신청일시
    CsdOutYon   As String       '       불출여부(Y:공급실에서 해당부서로 부출,N:공급요청만 있는 상태)
    CsdOspYon As String         '20030101 lek 소모마감에서 저장한 것인지 여부
End Type
    
    '--------------------------------------------------------------------
    ' 병동 Consult 정보
    '--------------------------------------------------------------------
Type CstInfRec
    CstDepCod  As String      ' 접수과 코드  Key
    CstOcmNum  As String * 10 ' 입원 번호    Key
    CstAdpDte  As String      ' 적용 일자    Key
    CstChtNum  As String * 8  ' 차트 번호
    CstPatNam  As String      ' 환자명
    CstNssCod  As String      ' 병동
    CstRomCod  As String      ' 병실
    CstExpDte  As String      ' 종료일
    CstFrmDep  As String      ' 의뢰과
    CstCurStt  As String      ' 현재상태      ("OA" : 등록, "OT" : 완료)
    CstCrtYon  As String      ' 재진코드생성여부(시작일로부터 30일 경과시 재진코드 하나씩 생성)
    CstGrpDep  As String      ' 그룹총괄과
    CstSplCmt  As String      ' 특기사항
    CstRetCmt  As String      ' 회신내용
End Type
    
    '--------------------------------------------------------------------
    ' 입원 발생진료비 정보
    '--------------------------------------------------------------------
Type SrpInfRec
    SrpEndDte  As String       'SrpInfKey 마감일자
    SrpOcmNum  As String * 10  'SrpInfKey 입원번호
    SrpOcmSeq  As String * 2   'SrpInfKey 내원변경순서
    SrpDupSeq  As String * 2   'SrpInfKey 이중보험순서
    SrpDepCod  As String       'SrpInfKey 진료과
    SrpChtNum  As String * 8   '1         챠트번호
    SrpDtrCod  As String       '2         주치의
    SrpInsCod  As String       '3         보험유형
    SrpInsSeq  As String * 2   '4         보험유형순서
    SrpTotAmt  As String       '5         진료비총액
    SrpCorAmt  As String       '6         조합청구액
    SrpNonAmt  As String       '7         비급여총액          NonOwn
    SrpOwnAmt  As String       '8         급여 본인부담액     InsOwn
    SrpTotOwn  As String       '9         본인부담총액
    SrpInsAmt  As String       '10        급여총액            InsTot
    SrpSpcAmt  As String       '11        특진료
    SrpAskAmt  As String       '12        환자청구액
    SrpDisAmt  As String       '13        할인액
    SrpFutAmt  As String       '14        후불액
    SrpOldAmt  As String       '15        기수납액
    SrpNewAmt  As String       '16        수납액
    SrpGrnAmt  As String       '17        보증금중 계산에 적용된 금액
    SrpRcpNum  As String * 10  '18        영수증번호
    SrpOldNum  As String * 10  '19        이전영수증번호
    SrpCalDte  As String       '20        중간계산일
    SrpUpdDtm  As String       '21        계산일시(system time)
    SrpUidCod  As String       '22        담당자코드
    SrpDimAmt  As String       '23        감면액
    SrpChgYon  As String       '24
End Type
    
    '--------------------------------------------------------------------
    ' 입원 발생 상세 정보
    '--------------------------------------------------------------------
Type SdlInfRec
    SdlEndDte  As String       'SdlInfKey 마감일자
    SdlOcmNum  As String * 10  'SdlInfKey 입원번호
    SdlOcmSeq  As String * 2   'SdlInfKey 내원변경순서
    SdlDupSeq  As String * 2   'SdlInfKey 이중보험순서
    SdlDepCod  As String       'SdlInfKey 진료과
    SdlIncCod  As String * 2   'SdlInfKey 수입원
    SdlChtNum  As String * 8   '          챠트번호
    SdlInsCod  As String       '          보험유형
    SdlInsSeq  As String * 2   '          유형순서
    SdlDtrCod  As String       '          주치의
    SdlInsAct  As String       '          급여행위
    SdlInsMat  As String       '          급여재료
    SdlNonAct  As String       '          비급여행위
    SdlNonMat  As String       '          비급여재료
    SdlInsAmt  As String       '          급여총액
    SdlNonAmt  As String       '          비급여액
    SdlInsOwn  As String       '          급여 본인부담액
    SdlTotOwn  As String       '          본인부담액
    SdlSpcAmt  As String       '          특진료
End Type
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' 식이코드 ImlInf
    ''''''''''''''''''''''''''''''''''''''''''''''
Type ImlInfRec
    ImlOcmNum As String * 10     'ImlInfKey 입원번호
    ImlAdpDte As String      'ImlInfKey 시작일자
    ImlPatTyp As String      'ImlInfKey 환자구분 ( I: 본인, F:환자가족,보호자)
    ImlExpDte As String      '1         종료일자
    ImlBrfCod As String      '2         아침식사코드
    ImlBrfQty As String      '3         아침식사용량  default 1
    ImlBrfCal As String      '4         아침식사칼로리  default ""
    ImlBrfcc  As String      '5         아침식사 cc     default ""
    ImlLnhCod As String      '6         점심식사코드
    ImlLnhQty As String      '7         점식식사용량  default 1
    ImlLnhCal As String      '8         점심식사칼로리  default ""
    ImlLnhcc As String       '9         점심식사 cc     default ""
    ImlDnrCod As String      '10        저녁식사코드
    ImlDnrQty As String      '11        저녁식사용량   default 1
    ImlDnrCal As String      '12        저녁식사칼로리  default ""
    ImlDnrcc  As String      '13        저녁식사 cc     default ""
    ImlSplCmt As String      '14        특기사항
    ImlWhyCod As String      '15        변경사유
    ImlUidCod As String      '16        입력담당자
    ImlEntDtm As String      '17        입력일시
    ImlEtcCmt As String      '18        기타 사유(특기사항)
End Type
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' 식이코드 ImlInfHst
    ''''''''''''''''''''''''''''''''''''''''''''''
    
Type ImlHstRec
    ImlOcmNum As String * 10    'ImlInfHstKey 입원번호
    ImlUpdDtm As String         'ImlInfHstKey Update time
    ImlSerNum As String * 3     'ImlInfHstKey 순차적 번호
    ImlAdpDte As String         'ImlInfHstKey 시작일자
    
    ImlPatTyp As String      '1 환자구분 ( I: 본인, F:환자가족,보호자)
    ImlExpDte As String      '2         종료일자
    ImlBrfCod As String      '3         아침식사코드
    ImlBrfQty As String      '4         아침식사용량    default 1
    ImlBrfCal As String      '5         아침식사칼로리  default ""
    ImlBrfcc  As String      '6         아침식사 cc     default ""
    ImlLnhCod As String      '7         점심식사코드
    ImlLnhQty As String      '8         점식식사용량    default 1
    ImlLnhCal As String      '9         점심식사칼로리  default ""
    ImlLnhcc As String       '10        점심식사 cc     default ""
    ImlDnrCod As String      '11        저녁식사코드
    ImlDnrQty As String      '12        저녁식사용량    default 1
    ImlDnrCal As String      '13        저녁식사칼로리  default ""
    ImlDnrcc  As String      '14        저녁식사 cc     default ""
    ImlSplCmt As String      '15        특기사항
    ImlWhyCod As String      '16        변경사유
    ImlUidCod As String      '17        입력담당자
    ImlEntDtm As String      '18        입력일시
    ImlEtcCmt As String      '19        기타 사유(특기사항)
    ImlUpdUid As String      '20        고친사람 아이디
End Type
    
    '--------------------------------------------------------------------
    '   하루 동일두과이상 조제료,수기료,수렬료 자동계산 프로그램
    '--------------------------------------------------------------------
Type MthInfRec
    MthChtNum As String     'Key Value  1         챠트번호
    MthOdrDte As String     'Key Value  2         처방일자
    MthOcmNum As String     'Key Value  3         내원번호
    MthOdrCod As String     'Key Value  4         자동산정코드    KK010,KK020...,J1000,J2000,X1000
    MthOdrStt As String     'Data Value 0         처방상태'OE'OC
    MthOdrQty As String     'Data Value 1         처방수량
    MthOdrDay As String     'Data Value 2         처방일수
    MthOdrTms As String     'Data Value 3         처방횟수
    MthOdrAmt As String     'Data Value 3         원래조제료
    MthAdpAmt As String     'Data Value 3         산정금액
End Type
    
Type ImgInfRec
    ImgStrDtm As String     'Key                조정마감시작일시
    ImgEndDtm As String     '                   조정마감종료일시
    ImgPatTyp As String     '                   조정환자구분(O'외래, I'입원)
    ImgSavYon As String     '                   (입원)조정저장구분
    ImgCalYon As String     '                   (입원)조정정산구분
    ImgDrgYon As String     '                   (입원)조정DRG정산구분
    ImgAccYon As String     '                   (입원)조정계정정산구분
    ImgSavOut As String     '                   (외래)조정저장구분
    ImgCalOut As String     '                   (외래)조정정산구분
    ImgAccOut As String     '                   (외래)조정계정구분
End Type
    
    '외래 수납과 외래OCS 사이의 ChtNum Locking 971216
Type LocChtRec
    LocChtNum As String * 8     'Key
    LocLevCod As String         'Key        (프로그램에 Level을준다.)
    LocExeNam As String         '실행화일이름
    LocUidCod As String         '담당자
    LocIpAddr As String         '담당IP Address
    LocChtDtm As String         '일시
End Type
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' 프로그램을 실행한 컴퓨터의 정보를 담아둔다.
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Type TcpInfRec
    TcpIp As String         'Key - IP Address
    TcpExeNam As String     'Key - 실행화일명
    TcpPath As String       '      실행경로
    TcpExeVer As String     '      실행화일버젼
    TcpComNam As String     '      컴퓨터 이름
    TcpAcpDtm As String     '      접속일자
    'TcpUidCod As String     '      담당자 ID
    TcpPortNum As String    '      WinSock Port Number
End Type

    '--------------------------------------------------------------------
    ' 영안실관리 프로그램(Death)
    '--------------------------------------------------------------------
Type DthInfRec
    DthChtNum  As String   ' DthInfKey 챠트번호
    DthIoODgs  As String   '           내부,외부사암
    DthWatStr  As String   '           대기실사용시작일시
    DthWatEnd  As String   '           대기실사용종료일시
    DthDthStr  As String   '           안치실사용시작일시
    DthDthEnd  As String   '           안치실사용종료작일시
    DthRefCmd  As String   '           참조사항
End Type
    
    'Not Billing ... 미청구관리
Type NblInfRec
    NblBilDte As String         'Key 1  청구년월
    NblChtNum As String         'Key 2  챠트번호
    NblInsCod As String         'Key 3  보험유형
    NblSeqNum As String         'Key 4  일련번호
    NblStrDte As String         '1      시작일자
    NblEndDte As String         '2      종료일자
    NblDtlDte As String         '3      상세일자
    NblTotAmt As String         '4      총진료비
    NblAskAmt As String         '5      본인부담
    NblCorAmt As String         '6      청구금액
    NblFutAmt As String         '7      후불액
    NblEndFlg As String         '8      종료여부
    NblCmtRef As String         '9      Memo Field
    NblIotFlg As String         '10     입원/외래
    NblAccCod As String         '11     계정코드
    NblRcpNum As String         '12     영수증번호
End Type
    
    '--------------------------------------------------------------------
    ' 미청구 기본 정보
    '--------------------------------------------------------------------
Type NbsInfRec
    NbsFrmDte  As String       'NbsInfKey  통계시작일자
    NbsEndDte  As String       '1          통계종료일자
    NbsPrcDte  As String       '2          마감처리일자
    NbsUidCod  As String       '3          마감담당자
    NbsCloCnt  As String       '4          마감횟수
End Type
    
    '--------------------------------------------------------------------
    '대출/불출을 여기서 하나로 관리하기 위해서(980624박남작성)
    '--------------------------------------------------------------------
Type DlyInfRec
    DlyChtNum    As String * 8  ' Key   차트번호
    DlyDepCod    As String      ' Key   대출하려는 환자의 과 (통합일때는 "ALL",과별일때는 과목코드)
    DlyAcpDtm    As String      ' Key   대출일시
    DlyOrXray    As String      ' 1     차트=0 or X-Ray=1
    DlyOutDep    As String      ' 2     대출과
    DlyResDtm    As String      ' 3     반납일시
    DlyDlvNam    As String      ' 4     대출자성명
    DlyOutUid    As String      ' 5     반출자
    DlyRcvUid    As String      ' 6     반입자
    DlyOutGrp    As String      ' 7     대출자부서명
    DlyMemo      As String      ' 8     메모
    DlyMemDep    As String      ' 9     대출과(과별챠트와 통합챠트를 모두 적용하기 위한 필드)
                                '       과별일때는 ChtDepCod와 ChtMemDep가 같다.
    DlyOcmNum    As String      '10     내원번호
    DlyRelFlg    As String      '11
    DlyRemark    As String      '12
    
End Type


'차트불출용 파일
'기존에 DlyInf / StbInf / DlyTmp 를 사용하던 것을 하나로 관리한다.
Type CdeInfRec  'Chart Delibery Information
    CdeChtNum    As String * 8  ' Key   차트번호
    CdeAskDte    As String      ' Key   의뢰일자 ("" 인경우 현재 챠트의 상태, 일자가 있는 경우는 History)
    CdeAskHMS    As String      ' Key   의뢰시간 ("" 인경우 현재 챠트의 상태, 일자가 있는 경우는 History)
    CdeChtStt    As String      '       챠트의 현재 상태 (보관중 "E" / 대출중 "O" / 대출요청 "A" / 의뢰하였으나 타과에 차트가 있는경우 "W")
    CdeOutDep    As String      '       대출하려는 환자의 과
    CdeAcpDtm    As String      '       접수일시
    CdeAskUid    As String      '       의뢰자
    CdeOutDtm    As String      '       반출일시
    CdeOutUid    As String      '       반출자
    CdeResDtm    As String      '       반납일시
    CdeResUid    As String      '       반입자
    CdeDgsNfs    As String      '       초재진 구분 (신환을 구분하기 위해)
    CdeFlgStt    As String      '       전과로 인한 불출인지... 변경으로 인한 불출인지....
    CdePrnYon    As String      '       가이드지 출력여부
    CdeMemo      As String      '       메모
End Type

    '지원부서 보류정보
Type LbbInfRec
    LbbSotCod   As String        'LbbInfKey   항목분류
    LbbOcmNum   As String        'LbbInfKey   내원번호
    LbbOdrDte   As String        'LbbInfKey   처방일자
    
    LbbChtNum   As String        '챠트번호
    LbbLabDte   As String        '시행일자
    LbbLabTim   As String        '시행시간
    LbbCanFlg   As String        '처리상태   "OB":보류, "OC":취소, "OE" :완료
    LbbOcmStt   As String        '입원/외래
    LbbPreYon   As String        '예약여부
End Type
    
    '*************************************
    ' 검사결과 기본파일   (정도관리를 하기위한 파일중 일부)
    '*************************************
Type RsbInfRec 'Result Basic Information
    RsbAcpDte   As String   '접수일자(K-1)
    RsbAcpCod   As String   '접수코드(K-2)
    RsbAcpNum   As String   '접수번호(K-3)
    RsbSpmCod   As String   '검체코드(K-4)
    
    RsbOcmNum   As String   '1  내원번호(D-1)
    RsbChtNum   As String   '2  차트번호
    RsbItfYon   As String   '3  Interface 여부
    RsbPrnYon   As String   '4 결과지 출력 여부
    RsbPrnUid   As String   '5 결과지 출력 ID
    RsbPrnDtm   As String   '6 결과지 출력일시
    RsbAcpTim   As String   '7 접수시간
    RsbSpcCmt   As String   '8 특기사항
    RsbOkSw     As String   '9 승인여부(N:없음,S:일부승인,A:전체승인)
    RsbOspIsp   As String   '10 외래냐, 입원이냐.      2002.01.15 sebal 파트접수
    RsbTryNum   As String   '11 파트접수 검체별 번호   2002.01.15 sebal 파트접수
    RsbWrdCod   As String   '12 병동                   2002.01.17 sebal 파트접수
    RsbParTms   As String   '13 학부접수횟수           2002.01.18 sebal 파트접수
    RsbParDte   As String   '14 파트접수일자
    RsbParTim   As String   '15 파트접수시간
    RsbSpmNum   As String   '16 검체번호
    RsbSpmDte   As String   '17 검체 접수 일자
    RsbSpmTim   As String   '18 검체 접수 일시
    RsbSpmUid   As String   '19 검체 접수 ID
End Type
    
    
    '*************************************
    '   검사결과
    '*************************************
Type ResInfRec
    ResAcpDte   As String   '접수일자(K-1)
    ResAcpCod   As String   '접수코드(K-2)
    ResAcpNum   As String   '접수번호(K-3)
    ResSpmCod   As String   '검체코드(K-4)
    ResLabCod   As String   '검사코드(K-5)

    ResOcmNum   As String   '1  내원번호(D-1)
    ResChtNum   As String   '2  차트번호
    ResSotCod1  As String   '3  항목대분류
    ResSotCod2  As String   '4  항목소분류
    ResJbsSeq   As String   '5  접수화면순서(D-5)
    ResRltSeq   As String   '6  결과조회화면순서
    ResMzhMin   As String   '7  결과최소값
    ResMzhMax   As String   '8  결과최대값
    ResMzhRef   As String   '9  결과표준값
    ResMzhUnt   As String   '10 결과단위(D-10)
    ResMzhMnt   As String   '11 검사결과
    ResSplCmt   As String   '12 코멘트('의뢰과에서 입력한 코멘트)
    ResOdrDtm   As String   '13 의뢰일시
    ResAcpDtm   As String   '14 접수일시
    ResTstDtm   As String   '15 검사일시(D-15)
    ResUpdDtm   As String   '16 결과최종수정일시
    ResOdrUid   As String   '17 의뢰자ID
    ResAcpUid   As String   '18 접수자ID
    ResTstUid   As String   '19 검사자ID
    ResUpdUid   As String   '20 결과최종수정자ID(D-20)
    ResSeeYon   As String   '21 의사조회여부
    ResConLvl   As String   '22 결과신뢰도
    ResMzhTyp   As String   '23 결과 유형
    ResMzhLin   As String   '24 결과의 최대라인수
    ResShtNam   As String   '25 검사종목의 약칭(D-25)
    ResSclCod   As String   '26 의뢰검사처 코드
    ResPrtYon   As String   '27 워크리스트에 찍을항목인가?
    ResOspIsp   As String   '28 외래인가 입원인가
    ResMadYon   As String   '29 모 코드로부터 파생된 것인가 아닌가?
    ResJbsMth   As String   '30 접수방식 (검사종류별로 접수했나, 아니면 검체종류별로 접수 했나?)(D-30)
    ResJbsQty   As String   '31 접수 수량
    ResCasYon   As String   '32 수납여부
    ResEmgYon   As String   '33 응급여부
    ResOdrNum   As String   '34 처방번호 (From OspInf or IspInf)
    ResOdrSeq   As String   '35 Seq No. (From OspInf or IspInf) (D-35)
    ResOdrDep   As String   '36 처방과
    ResWrdCod   As String   '37 병동
    ResStaYon   As String   '38 통계여부
    ResOkYon    As String   '39 임상병리사 승인여부
    ResOkUid    As String   '40            승인ID
    ResOkDtm    As String   '41            승인일시
    ResRedYon   As String   '42 결과조회여부
    ResRedUid   As String   '43 결과조회 ID
    ResRedDtm   As String   '44 결과조회일시
    ResPanMax   As String   '45 Panic 상한값
    ResPanMin   As String   '46 Panic 하한값
    ResRepTyp   As String   '47 Report type(최종보고_F, 중간보고_I)
    ResWrdPrn   As String   '48 병동에서 출력여부를 기록...(출력:Y, 미출력:Null)    => 2001/12/10 james
    ResMchCod   As String   '49 02.04.27 sebal 검사할 검사장비 코드
'    ResBtlCod   As String   '50 용기코드
'    ResSpmNum   As String   '51 검체번호
'    ResTryNum   As String   '52 파트접수 검체별 번호
'    ResMicTyp   As String   '53 Micr labNum(작업번호)의 작업type 1,2,3,4,GS,AS...
'    ResLabNum   As String   '54 lab 작업번호
'    ResGroYon   As String   '55 Growth 여부(배양여부) - G:Growth , NG:No Growth
'    ResGroDte   As String   '56 배양일자(Afb Culture(AC), Fungus Culture(FC), blood culture(5))
    
End Type

Type MorInfRec  '미생물 배양 성상정보
    MorAcpDte   As String   '접수일자(K-1)
    MorAcpCod   As String   '접수코드(K-2)
    MorAcpNum   As String   '접수번호(K-3)
    MorSpmCod   As String   '검체코드(K-4)
    MorLabCod   As String   '검사코드(K-5)
    MorColTyp   As String   'Color(D-1)
    MorSurTyp   As String   '표면(D-2)
    MorEdgTyp   As String   '가장자리(D-3)
    MorHemTyp   As String   '용혈성(D-4)
    MorExiTyp   As String   '유주성(D-5)
    MorThiTyp   As String   '두께(D-6)
    MorLabNum   As String   '작업번호(D-7)
    MorMicTyp   As String   '작업번호head(D-8)
    
End Type

Type BacInfRec  '미생물 균정보
    BacAcpDte   As String   '접수일자(K-1)
    BacAcpCod   As String   '접수코드(K-2)
    BacAcpNum   As String   '접수번호(K-3)
    BacSpmCod   As String   '검체코드(K-4)
    BacLabCod   As String   '검사코드(K-5)
    BacBacCod   As String   '균코드(K-6)
    BacLabNum   As String   '작업번호(D-1)
    BacMicTyp   As String   '작업번호Head(D-2)
    
End Type

Type GroInfRec  '미생물 Growth결과 정보
    GroAcpDte   As String   '접수일자(K-1)
    GroAcpCod   As String   '접수코드(K-2)
    GroAcpNum   As String   '접수번호(K-3)
    GroSpmCod   As String   '검체코드(K-4)
    GroLabCod   As String   '검사코드(K-5)
    GroMicTyp   As String   '작업번호Head(D-1)
    GroGroYon   As String   'Growth여부:G,NG(D-2)
    GroRecCod   As String   'Growth결과코드(D-3)
    GroLabNum   As String   '작업번호(D-4)
    
End Type

Type StnInfRec  '미생물 Stain 결과정보

    StnAcpDte   As String   '접수일자(K-1)
    StnAcpCod   As String   '접수코드(K-2)
    StnAcpNum   As String   '접수번호(K-3)
    StnSpmCod   As String   '검체코드(K-4)
    StnLabCod   As String   '검사코드(K-5)
    StnStnCod   As String   'Stain 결과코드(K-6)
    StnLabNum   As String   '작업번호(D-1)
    StnMicTyp   As String   '작업번호Head(D-2)
    
End Type

Type AntInfRec  '미생물 항생제 정보
    AntAcpDte   As String   '접수일자(K-1)
    AntAcpCod   As String   '접수코드(K-2)
    AntAcpNum   As String   '접수번호(K-3)
    AntSpmCod   As String   '검체코드(K-4)
    AntLabCod   As String   '검사코드(K-5)
    AntBacCod   As String   '균코드(K-6)
    AntBioTyp   As String   '항생제코드(K-7)
    AntLabNum   As String   '작업번호(D-1)
    AntMicTyp   As String   '작업번호head(D-2)
    AntMicRes   As String   'Mic 결과(D-3)
    AntMicDan   As String   'Mic 단위(D-4)
    AntRisRes   As String   '판정 RIS(D-5)
    
End Type

    
Type EegInfRec
    EegChtNum As String * 8  'Key
    EegEegNum As String * 10
    EegUpdDte As String
End Type
    
'--------------------------
'- 카드를 현금처럼 만든다
'--------------------------
Type CrdInfRec
    CrdRcpNum As String * 10    '영수증번호 Key
    CrdCrdSeq As String * 2     '카드 순서  Key
    CrdOcmNum As String * 10    '내원번호
    CrdChtNum As String * 8     '차트번호
    CrdCrdCod As String         'Card 회사 코드
    CrdCrdNum As String         'Card Number
    CrdExpDte As String         'Card Expired Date - 카드 적용 년월
    CrdCtfnum As String         '승인번호
    CrdNewAmt As String         '결재금액
    CrdDivMth As String         '일시불 / 할부개월수
    CrdUseCod As String         '카드 사용자(본인, 배우자, 자녀, 기타)
    CrdCanYon As String         '결재취소 여부
    CrdUidCod As String         '취급자
    CrdUpdDtm As String         '입력일시
End Type
    
'================================================
'원외처방 발생건에 관한 자료...
'Index OutInfChtDteNum      'D-1, K-1, K-2
'================================================
Type OutInfRec
    OutOdrDte  As String       'OutInfKey  처방일자
    OutNum     As String * 5   'OutInfKey  교부번호
    OutOcmNum  As String * 10  'OutInfKey  OspInf의 내원번호
    OutOdrNum  As String * 4   'OutInfKey  OspInf의 처방번호
    OutOdrSeq  As String * 5   'OutInfKey  OspInf의 처방순서
    OutChtNum  As String * 8   '           챠트번호
    OutOdrStt  As String       '           처방전 상태(E:진행, C:취소, P:출력)
    OutUpdTms  As String       '           수정시간
    OutPatNam  As String       '           환자이름
    OutDepCod  As String       '           진료과목
    OutResNum  As String       '           주민번호
    OutCanNum  As String       '           취소교부번호                          <=추가
End Type


'////////////////////////////////////////////////////////////////
'/// 에... 말레이지아에서만 사용하는 프로그램인 "산전관리" 와 ///
'/// "예방접종"에 사용하는 마스타입네다. 010427 오성연        ///
'////////////////////////////////////////////////////////////////

'예방접종(Immunization Master)
Type ImmMstRec
    ImmOdrCod As String 'Key 접종코드
    
    ImmBasTms As String '접종기본횟수(1차,2차,3차,기초,추가)
    ImmUsgCod As String '접종용법(피내주사,피하주사,근육주사,경구투여)
    ImmRegCod As String '접종부위(왼팔삼부근,대퇴부전외측,허벅지전의측)
    ImmBasYon As String '기본접종여부
    ImmPrnSeq As String '화면출력순서

End Type

'예방접종(Immunization Infomation)
Type ImiInfRec
    ImiChtNum As String * 8 'Key 챠트번호
    ImiOdrCod As String     'Key 접종코드
    ImiAdpDte As String     'Key 접종일자
    
    ImiBasTms As String '접종차수(1차)
    ImiUsgCod As String '접종용법(피내주사,피하주사,근육주사,경구투여)
    ImiRegCod As String '접종부위(왼팔삼부근,대퇴부전외측,허벅지전의측)
    ImiWroYon As String '이상여부
    ImiSplCmt As String '비고
    ImiUntQty As String '투여량
    ImiLotNum As String '백신로트번호
    ImiOcmNum As String * 10    '내원번호
    ImiOdrNum As String * 4     '처방번호
    ImiOdrSeq As String * 5     '처방Seq

End Type

'산전관리(Postpartum Care)
Type PpcInfRec
    PpcChtNum As String     'Key 챠트번호
    PpcAdpDte As String     'Key 적용일자

    PpcWeight As String     '몸무게
    PpcBPStr As String      '혈압시작
    PpcBPEnd As String      '혈압끝
    PpcCyeCar As String     '태아심음
    PpcUlrYon As String     '초음파 측정여부
    PpcSplCmt As String     '기타사항
End Type

'병동검사 라벨출력여부(학부별)
Type BarInfRec
    BarOcmNum As String     'Key 내원번호
    BarOdrDte As String     'Key 적용일자
    BarLabTst As String     'Key 검사학부

    BarPrnDtm As String     '    출력일시
    BarPrnCnt As String     '    출력매수
    BarPrnUid As String     '    출력자 ID
End Type

'병동 카덱스
'임시로 이곳에 선언하지만 추후 EPR 의 기본 맵 BAS가 정해지면 그곳으로 옮겨야징
Type CdxInfRec
    CdxChtNum As String * 8         'key
    CdxOdrDte As String * 8         'Key
    CdxFreNot As String             'Free Note
    CdxCmtNot As String             '기타
    CdxFreNt2 As String             'Free Note
End Type
'
''원래의 예약어기능.
Type CrvInfRec
    CrvMotCod As String
    CrvChrCod As String
    CrvCodNam As String
End Type


'새로 만든 예약어 템플릿 글로벌.

Type CrvCtpInfRec
    CrvCtpWrdNam       As String        'K-1    병동
    CrvCtpMotCod       As String        'K-2    모코드
    CrvCtpChdCod       As String        'K-3    자식코드
    CrvCtpMotCodNam    As String        'D-1    모코드이름
    CrvCtpChdCodNam    As String        'D-2    자식코드이름
    CrvCtpCodCon       As String        'D-3    예약어(자식코드의 내용)
End Type

'--------------------------
''' 신생아 인적정보 관리
'--------------------------
Type BabInfRec
    BabChtNum As String         'K-1    아기 차트
    BabIcmNum As String         'D-1    내원번호
    BabBabNam As String         'D-2    아기 이름
    BabBonDtm As String         'D-3    출생일시
    BabSexTyp As String         'D-4    아기 성별
    BabBabHet As String         'D-5    아기 신장
    BabBabWet As String         'D-6    아기 몸무게
    BabBabAbo As String         'D-7    아기 혈액형
    BabMomCht As String         'D-8    어머니 차트
    BabMomIcm As String         'D-9    어머니 내원번호
    BabFatNam As String         'D-10   아버지 성함
    BabMomPrd1 As String        'D-11   산모 임신기간(주수)
    BabMomPrd2 As String        'D-12   산모 임신기간(일수)
    BabBabTyp As String         'D-13   출생아 상태(1명, 쌍태, 삼태, 다태)
    BabBabSeq As String         'D-14   다태일 경우 출산 순위
    BabDtrCod As String         'D-15   담당 의사
    BabUidCod As String         'D-16   담당자
    BabAcpDtm As String         'D-17   입력일시
    BabSpcCmt As String         'D-18   특기사항
End Type

'===================================
'Nurse Duty Schedule 정보
'===================================
Type NrsInfRec
     NrsEmpNum As String       'NrsInfKey 사원번호
     NrsRqtDte As String       'NrsInfKey 근무일자
     NrsRqtCod As String       '1         근무코드
     NrsWrdCod As String       '2         병동코드
     NrsCodStt As String       '3         코드상태("W":신청, "O":확정,"T"근무확인)
     NrsWrkTms As String       '4         초과근무시간
     NrsWrkCmt As String       '5         특기사항
     NrsOldCod As String       '6         요청당시의 코드
     NrsMstUid As String       '7         UidMst 의 UidCod
     NrsUpdUid As String       '8         입력담당
     NrsUpdDtm As String       '9         수정일시
End Type

'===================================
'Nurse Display Sequenct 정보
'===================================
Type NrsSeqRec
     NrsWrdCod As String       'NrsInfKey 병동코드
     NrsMstUid As String       'NrsInfKey UidMst 의 UidCod
     NrsAdpDte As String       'NrsInfKey 적용일자
     NrsSeq    As String * 3   '1         순서
End Type

'''마일리지 관리
Type MilInfRec
    MilChtNum As String     ' 챠트번호
    MilAcpDte As String     ' 발생일자
    MilOcrTyp As String     ' 발생구분(차감구분)  A 발생, D 차감(혜택을 받았을때만)
    MilOcmNum As String     ' 내원번호
    
    MilRsn    As String     ' 마일리지 부여사유
    MilDgsCnt As String     ' 외래,입원 접수 마일리지
    MilCnt    As String     ' 취득 및 사용마일리지
End Type


'2003-05-12 corebrain :성삼에선 있었던  Global
'--------------------------------------------------------------------
' TPM 코드 결과정보 TpmInf
'--------------------------------------------------------------------
Type TpmInfRec
    TpmCytoNum  As String   'K-1     '세포/조직번호
    TpmCodSeq   As String   'K-2     '입력번호
    TpmCodDat   As String            'T.P.M. 코드
End Type

Type EdcInfRec
    EdcLmpDte As String     '최종생리일
    EdcChtNum As String     '차트번호
    EdcEdcDte As String     '분만예정일
    EdcAbrYon As String     '유산여부
    EdcAbrNum As String     '유산횟수
    EdcLbrYon As String     '분만여부(타병원포함)
    EdcLbrSeq As String     '분만순서
    EdcHspLbr As String     '우리병원에서 분만 여부
End Type
'---------------------------------------------------------------------

'narcotic Information (마약/향정 관리정보)
Type NarInfRec
    NarOdrDte   As String        'Key 처방일자
    NarMdcNum   As String * 10   'Key 마약/향정 관리번호
    NarOcmNum   As String * 10   'Key 내원번호
    NarOdrNum   As String * 4    'Key 처방번호
    NarOdrSeq   As String * 5    'Key 처방순서
    NarIOsw     As String        '    외래(O)/입원(I)
    NarInpCan   As String        '    취소처방(CANCEL)인지 입력된 처방(INPUT)인지
    NarPrtDtm   As String        '    출력일시
    NarPrtID    As String        '    출력ID
    NarOutDtm   As String        '    불출일시 및 수령자의 수령일시 (취소처방의 경우 반납일시)
    NarOutID    As String        '    불출ID
    NarRcvID    As String        '    수령자 (취소처방의 경우 반납한 사람)
    NarInpPrt   As String        '    입력부서 (외래의 경우 진료과, ER은 ER, 입원은 입력병동)
    NarEmgYon   As String        '    응급여부
    NarInpDtm   As String        '    입력일시
End Type

'처방 출력 정보 저장
 Type PrnInfRec
        
    PrnOcmNum As String * 10 'Key 입원번호
    PrnSeq As String * 10   'Key 순서
    
    PrnPatTyp As String     '입원/외래 -> I/O
    PrnUsrID As String      '사용자 ID
    PrnDate As String       '출력일
    PrnFromDate As String   '처방 조회 시작일
    PrnToDate As String     '처방 조회 종료일
    
End Type

'KG정보
Type PkgInfRec
    PkgChtNum   As String * 8   'Key 챠트번호
    PkgKG       As String       '    KG
    PkgChkDtm   As String       '    측정일자
End Type

'입원환자 외래 내원예약처리
Type OrvInfRec
    OrvIcmNum   As String * 10  'Key 입원내원번호
    OrvSeq      As String * 2   'Key
    OrvDepCod   As String       '    예약 진료과
    OrvDtrCod   As String       '    예약 주치의
    OrvRsvDte   As String       '    예약일자
    OrvRsvTim   As String       '    예약시간
    OrvRsvOcm   As String * 10  '    예약된 내원번호
End Type

Public Sub PkgInfLoad(sPrmValue As String, ptPkgData As PkgInfRec)
    
    On Error GoTo PkgInfLoad

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 10)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    ptPkgData.PkgChtNum = vVal(i)
    i = i + 1
    ptPkgData.PkgKG = vVal(i)
    i = i + 1
    ptPkgData.PkgChkDtm = vVal(i)
    
    Exit Sub

PkgInfLoad:
    Resume Next
    
End Sub

Public Sub PkgInfStore(sPrmKey As String, sPrmValue As String, ptPkgData As PkgInfRec)

    sPrmKey = Format((ptPkgData.PkgChtNum), "@@@@@@@@") & Chr(5)

    sPrmValue = ptPkgData.PkgKG & Chr(5)
    sPrmValue = sPrmValue & ptPkgData.PkgChkDtm & Chr(5)
    
End Sub


Public Sub PrnInfLoad(sPrmValue As String, tPrmPrnData As PrnInfRec)
    
    tPrmPrnData.PrnOcmNum = piece(sPrmValue, Chr(5), 1) 'Key 입원번호
    tPrmPrnData.PrnSeq = piece(sPrmValue, Chr(5), 2)  'Key 순서
    
    tPrmPrnData.PrnPatTyp = piece(sPrmValue, Chr(5), 3)    '입원/외래 -> I/O
    tPrmPrnData.PrnUsrID = piece(sPrmValue, Chr(5), 4)      '사용자 ID
    tPrmPrnData.PrnDate = piece(sPrmValue, Chr(5), 5)       '출력일
    tPrmPrnData.PrnFromDate = piece(sPrmValue, Chr(5), 6)   '처방 조회 시작일
    tPrmPrnData.PrnToDate = piece(sPrmValue, Chr(5), 7)    '처방 조회 종료일
    
End Sub

Public Sub PrnInfStore(sPrmKey As String, sPrmValue As String, tPrmData As PrnInfRec)

    sPrmKey = Format((tPrmData.PrnOcmNum), "@@@@@@@@@@") & Chr(5) 'Key 챠트번호
    sPrmKey = sPrmKey & Format((tPrmData.PrnSeq), "@@@@@@@@@@") & Chr(5)             'Key 접용일자

    sPrmValue = tPrmData.PrnPatTyp & Chr(5)                         '입원/외래 -> I/O
    sPrmValue = sPrmValue & tPrmData.PrnUsrID & Chr(5)             '사용자 ID
    sPrmValue = sPrmValue & tPrmData.PrnDate & Chr(5)              '출력일
    sPrmValue = sPrmValue & tPrmData.PrnFromDate & Chr(5)              '처방 조회 시작일
    sPrmValue = sPrmValue & tPrmData.PrnToDate & Chr(5)              '처방 조회 종료일
    
End Sub

Public Sub MilInfLoad(sPrmValue As String, tPrmMilData As MilInfRec)
    On Error GoTo MilInfLoad

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmMilData.MilChtNum = vVal(i)
    i = i + 1
    tPrmMilData.MilAcpDte = vVal(i)
    i = i + 1
    tPrmMilData.MilOcrTyp = vVal(i)
    i = i + 1
    tPrmMilData.MilOcmNum = vVal(i)
    i = i + 1
    tPrmMilData.MilRsn = vVal(i)
    i = i + 1
    tPrmMilData.MilDgsCnt = vVal(i)
    i = i + 1
    tPrmMilData.MilCnt = vVal(i)
    
    Exit Sub

MilInfLoad:
    Resume Next

End Sub
Public Sub MilInfStore(sPrmKey As String, sPrmValue As String, tPrmMilData As MilInfRec)
    
    sPrmKey = Format(CDouble(tPrmMilData.MilChtNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmMilData.MilAcpDte & Chr(5)
    sPrmKey = sPrmKey & tPrmMilData.MilOcrTyp & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmMilData.MilOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmMilData.MilRsn & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmMilData.MilDgsCnt), "@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmMilData.MilCnt), "@@@@@@") & Chr(5)
        
End Sub


Sub NrsInfLoad(sPrmValue As String, tPrmData As NrsInfRec)
    Dim i As Integer

    i = 1
    tPrmData.NrsEmpNum = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsRqtDte = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsRqtCod = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsWrdCod = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsCodStt = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsWrkTms = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsWrkCmt = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsOldCod = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsMstUid = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsUpdUid = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsUpdDtm = piece(sPrmValue, Chr(5), i)

End Sub

Sub NrsInfStore(sPrmKey As String, sPrmValue As String, tPrmData As NrsInfRec)
    
    sPrmKey = tPrmData.NrsEmpNum & Chr(5)
    sPrmKey = sPrmKey & tPrmData.NrsRqtDte & Chr(5)

    sPrmValue = tPrmData.NrsRqtCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NrsWrdCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NrsCodStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NrsWrkTms & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NrsWrkCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NrsOldCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NrsMstUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NrsUpdUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NrsUpdDtm

End Sub

Sub NrsSeqLoad(sPrmValue As String, tPrmData As NrsSeqRec)
    
    Dim i As Integer

    i = 1
    tPrmData.NrsWrdCod = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsMstUid = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsAdpDte = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmData.NrsSeq = piece(sPrmValue, Chr(5), i)

End Sub

Sub NrsSeqStore(sPrmKey As String, sPrmValue As String, tPrmData As NrsSeqRec)
    
    sPrmKey = tPrmData.NrsWrdCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.NrsMstUid & Chr(5)
    sPrmKey = sPrmKey & tPrmData.NrsAdpDte & Chr(5)

    sPrmValue = Format(CDouble(tPrmData.NrsSeq), "@@@") & Chr(5)

End Sub

Sub BabInfLoad(sPrmValue As String, tPrmData As BabInfRec)

On Error GoTo BabInfLoad_ErrorTraping

    Dim vVal()  As String
    Dim i As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.BabChtNum = vVal(i)            '아기차트
    i = i + 1
    tPrmData.BabIcmNum = vVal(i)            '내원번호
    i = i + 1
    tPrmData.BabBabNam = vVal(i)            '아기 이름
    i = i + 1
    tPrmData.BabBonDtm = vVal(i)            '출생일자 및 시간
    i = i + 1
    tPrmData.BabSexTyp = vVal(i)            '성별
    i = i + 1
    tPrmData.BabBabHet = vVal(i)            '신장
    i = i + 1
    tPrmData.BabBabWet = vVal(i)            '몸무게
    i = i + 1
    tPrmData.BabBabAbo = vVal(i)            '혈액형
    i = i + 1
    tPrmData.BabMomCht = vVal(i)            '어머니차트
    i = i + 1
    tPrmData.BabMomIcm = vVal(i)            '어머니 내원번호
    i = i + 1
    tPrmData.BabFatNam = vVal(i)            '아버지 성함
    i = i + 1
    tPrmData.BabMomPrd1 = vVal(i)           '산모 임신기간(주수)
    i = i + 1
    tPrmData.BabMomPrd2 = vVal(i)           '산모 임신기간(일수)
    i = i + 1
    tPrmData.BabBabTyp = vVal(i)            '출생아상태
    i = i + 1
    tPrmData.BabBabSeq = vVal(i)            '다태 출산순위
    i = i + 1
    tPrmData.BabDtrCod = vVal(i)            '담당의사
    i = i + 1
    tPrmData.BabUidCod = vVal(i)            '담당자
    i = i + 1
    tPrmData.BabAcpDtm = vVal(i)            '입력일시
    i = i + 1
    tPrmData.BabSpcCmt = vVal(i)            '특기사항
    
    Exit Sub

BabInfLoad_ErrorTraping:
    Resume Next

End Sub

Sub BabInfStore(sPrmKey As String, sPrmValue As String, tPrmData As BabInfRec)

    sPrmKey = Format((tPrmData.BabChtNum), "@@@@@@@@") & Chr(5) 'Key 챠트번호

    sPrmValue = tPrmData.BabIcmNum & Chr(5)                     '내원번호
    sPrmValue = sPrmValue & tPrmData.BabBabNam & Chr(5)         '아기 이름
    sPrmValue = sPrmValue & tPrmData.BabBonDtm & Chr(5)         '출생일시
    sPrmValue = sPrmValue & tPrmData.BabSexTyp & Chr(5)         '성별
    sPrmValue = sPrmValue & tPrmData.BabBabHet & Chr(5)         '신장
    sPrmValue = sPrmValue & tPrmData.BabBabWet & Chr(5)         '몸무게
    sPrmValue = sPrmValue & tPrmData.BabBabAbo & Chr(5)         '혈액형
    sPrmValue = sPrmValue & tPrmData.BabMomCht & Chr(5)         '어머니차트
    sPrmValue = sPrmValue & tPrmData.BabMomIcm & Chr(5)         '어머니내원번호
    sPrmValue = sPrmValue & tPrmData.BabFatNam & Chr(5)         '아버지성함
    sPrmValue = sPrmValue & tPrmData.BabMomPrd1 & Chr(5)        '산모임신기간(주수)
    sPrmValue = sPrmValue & tPrmData.BabMomPrd2 & Chr(5)        '산모임신기간(일수)
    sPrmValue = sPrmValue & tPrmData.BabBabTyp & Chr(5)         '출생아상태
    sPrmValue = sPrmValue & tPrmData.BabBabSeq & Chr(5)         '다태 출산순위
    sPrmValue = sPrmValue & tPrmData.BabDtrCod & Chr(5)         '담당의사
    sPrmValue = sPrmValue & tPrmData.BabUidCod & Chr(5)         '담당자
    sPrmValue = sPrmValue & tPrmData.BabAcpDtm & Chr(5)         '입력일시
    sPrmValue = sPrmValue & tPrmData.BabSpcCmt & Chr(5)         '특기사항

End Sub

Sub PpcInfLoad(sPrmValue As String, tPrmImmData As PpcInfRec)
        
    tPrmImmData.PpcChtNum = piece(sPrmValue, Chr(5), 1)
    tPrmImmData.PpcAdpDte = piece(sPrmValue, Chr(5), 2)
    tPrmImmData.PpcWeight = piece(sPrmValue, Chr(5), 3)
    tPrmImmData.PpcBPStr = piece(sPrmValue, Chr(5), 4)
    tPrmImmData.PpcBPEnd = piece(sPrmValue, Chr(5), 5)
    tPrmImmData.PpcCyeCar = piece(sPrmValue, Chr(5), 6)
    tPrmImmData.PpcUlrYon = piece(sPrmValue, Chr(5), 7)
    tPrmImmData.PpcSplCmt = piece(sPrmValue, Chr(5), 8)
    
End Sub

Sub PpcInfStore(sPrmKey As String, sPrmValue As String, tPrmData As PpcInfRec)

    sPrmKey = Format((tPrmData.PpcChtNum), "@@@@@@@@") & Chr(5) 'Key 챠트번호
    sPrmKey = sPrmKey & tPrmData.PpcAdpDte & Chr(5)             'Key 접용일자

    sPrmValue = tPrmData.PpcWeight & Chr(5)                         '몸무게
    sPrmValue = sPrmValue & tPrmData.PpcBPStr & Chr(5)             '혈압시작
    sPrmValue = sPrmValue & tPrmData.PpcBPEnd & Chr(5)             '혈압끝
    sPrmValue = sPrmValue & tPrmData.PpcCyeCar & Chr(5)              '태아심음
    sPrmValue = sPrmValue & tPrmData.PpcUlrYon & Chr(5)              '초음파 측정여부
    sPrmValue = sPrmValue & tPrmData.PpcSplCmt & Chr(5)              '기타사항
End Sub


'/// 말레이지아에서만 사용하는 예방접종 프로그램에서 사용하는 모듈
Public Sub ImmMstLoad(sPrmValue As String, tPrmImmData As ImmMstRec)
        
    tPrmImmData.ImmOdrCod = piece(sPrmValue, Chr(5), 1)
    tPrmImmData.ImmBasTms = piece(sPrmValue, Chr(5), 2)
    tPrmImmData.ImmUsgCod = piece(sPrmValue, Chr(5), 3)
    tPrmImmData.ImmRegCod = piece(sPrmValue, Chr(5), 4)
    tPrmImmData.ImmBasYon = piece(sPrmValue, Chr(5), 5)
    tPrmImmData.ImmPrnSeq = piece(sPrmValue, Chr(5), 6)
                 
End Sub
'/// 말레이지아에서만 사용하는 예방접종 프로그램에서 사용하는 모듈
Sub ImiInfload(sPrmValue As String, tPrmImiData As ImiInfRec)
    Dim i As Integer

    i = 1
    tPrmImiData.ImiChtNum = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiOdrCod = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiAdpDte = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiBasTms = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiUsgCod = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiRegCod = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiWroYon = piece(sPrmValue, Chr(5), i)                 '
    i = i + 1
    tPrmImiData.ImiSplCmt = piece(sPrmValue, Chr(5), i)                 '
    i = i + 1
    tPrmImiData.ImiUntQty = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiLotNum = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiOcmNum = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiOdrNum = piece(sPrmValue, Chr(5), i)                '
    i = i + 1
    tPrmImiData.ImiOdrSeq = piece(sPrmValue, Chr(5), i)                '

End Sub

Public Sub ChtInfLoad(sPrmValue As String, tPrmData As ChtInfRec)

    On Error GoTo ChtInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.ChtNum = vVal(i)
    i = i + 1
    tPrmData.ChtDscCc = vVal(i)
    i = i + 1
    tPrmData.ChtDscPhx = vVal(i)
    i = i + 1
    tPrmData.ChtDscSpc = vVal(i)
    i = i + 1
    tPrmData.ChtOldDscCc = vVal(i)
    i = i + 1
    tPrmData.ChtOldDscPhx = vVal(i)
    i = i + 1
    tPrmData.ChtOldDscSpc = vVal(i)
    i = i + 1
    tPrmData.ChtUidCod = vVal(i)
    i = i + 1
    tPrmData.ChtEntDtm = vVal(i)
    
    Exit Sub

ChtInfLoad_ErrorTraping:
    Resume Next

End Sub
    
    
Public Sub ChtInfStore(sPrmKey As String, sPrmValue As String, tPrmData As ChtInfRec)

    
    sPrmKey = Format((tPrmData.ChtNum), "@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.ChtDscCc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtDscPhx & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtDscSpc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtOldDscCc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtOldDscPhx & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtOldDscSpc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtEntDtm & Chr(5)
    
End Sub

    
Public Sub ChtIOInfLoad(sPrmValue As String, tPrmData As ChtIOInfRec)

    On Error GoTo ChtIOInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.ChtChtNum = vVal(i)
    i = i + 1
    tPrmData.ChtDepCod = vVal(i)
    i = i + 1
    tPrmData.ChtAcpDtm = vVal(i)
    i = i + 1
    tPrmData.ChtOrXray = vVal(i)
    i = i + 1
    tPrmData.ChtOutDep = vVal(i)
    i = i + 1
    tPrmData.ChtResDtm = vVal(i)
    i = i + 1
    tPrmData.ChtDlvNam = vVal(i)
    i = i + 1
    tPrmData.ChtOutUid = vVal(i)
    i = i + 1
    tPrmData.ChtRcvUid = vVal(i)
    i = i + 1
    tPrmData.ChtOutGrp = vVal(i)
    i = i + 1
    tPrmData.ChtMemo = vVal(i)
    i = i + 1
    tPrmData.ChtMemDep = vVal(i)
                
    Exit Sub

ChtIOInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ChtIOInfStore(sPrmKey As String, sPrmValue As String, tPrmData As ChtIOInfRec)

    
    sPrmKey = Format(tPrmData.ChtChtNum, "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ChtDepCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ChtAcpDtm & Chr(5)
    
    sPrmValue = tPrmData.ChtOrXray & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtOutDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtResDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtDlvNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtOutUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtRcvUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtOutGrp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtMemo & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ChtMemDep & Chr(5)

End Sub

    
Public Sub CltInfLoad(sPrmValue As String, tPrmData As CltInfRec)

    On Error GoTo CltInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.CltOcmNum = vVal(i)
    i = i + 1
    tPrmData.CltAcpDte = vVal(i)
    i = i + 1
    tPrmData.CltDepCod = vVal(i)
    i = i + 1
    tPrmData.CltUidCod = vVal(i)
    i = i + 1
    tPrmData.CltComStt = vVal(i)
    
    
    Exit Sub

CltInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CltInfStore(sPrmKey As String, sPrmValue As String, tPrmData As CltInfRec)

    
    sPrmKey = Format((tPrmData.CltOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CltAcpDte & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CltDepCod & Chr(5)
    
    sPrmValue = tPrmData.CltUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CltComStt & Chr(5)
    
End Sub

    
Public Sub CrdInfLoad(sPrmValue As String, tPrmData As CrdInfRec)

On Error GoTo CrdInfLoad_ErrorTraping

    Dim vVal()  As String
    Dim i As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.CrdRcpNum = vVal(i)
    i = i + 1
    tPrmData.CrdCrdSeq = vVal(i)
    i = i + 1
    tPrmData.CrdOcmNum = vVal(i)
    i = i + 1
    tPrmData.CrdChtNum = vVal(i)
    i = i + 1
    tPrmData.CrdCrdCod = vVal(i)
    i = i + 1
    tPrmData.CrdCrdNum = vVal(i)
    i = i + 1
    tPrmData.CrdExpDte = vVal(i)
    i = i + 1
    tPrmData.CrdCtfnum = vVal(i)
    i = i + 1
    tPrmData.CrdNewAmt = vVal(i)
    i = i + 1
    tPrmData.CrdDivMth = vVal(i)
    i = i + 1
    tPrmData.CrdUseCod = vVal(i)
    i = i + 1
    tPrmData.CrdCanYon = vVal(i)
    i = i + 1
    tPrmData.CrdUidCod = vVal(i)
    i = i + 1
    tPrmData.CrdUpdDtm = vVal(i)

    Exit Sub

CrdInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CrdInfStore(sPrmKey As String, sPrmValue As String, tPrmData As CrdInfRec)

    sPrmKey = Format(CDouble(tPrmData.CrdRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.CrdCrdSeq), "@@") & Chr(5)

    sPrmValue = Format(CDouble(tPrmData.CrdOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.CrdChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdCrdCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdCrdNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdCtfnum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdNewAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdDivMth & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdUseCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdCanYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrdUpdDtm & Chr(5)


End Sub

    
Public Sub CsdInfLoad(sPrmValue As String, tPrmData As CsdInfRec)

    On Error GoTo CsdInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.CsdAcpDte = vVal(i)
    i = i + 1
    tPrmData.CsdDgsDEN = vVal(i)
    i = i + 1
    tPrmData.CsdUsgPrt = vVal(i)
    i = i + 1
    tPrmData.CsdDgsTbl = vVal(i)
    i = i + 1
    tPrmData.CsdSeq = vVal(i)
    i = i + 1
    tPrmData.CsdDepTyp = vVal(i)
    i = i + 1
    tPrmData.CsdCsrcod = vVal(i)
    i = i + 1
    tPrmData.CsdCsmYon = vVal(i)
    i = i + 1
    tPrmData.CsdRqtQty = vVal(i)
    i = i + 1
    tPrmData.CsdOutQty = vVal(i)
    i = i + 1
    tPrmData.CsdUidCod = vVal(i)
    i = i + 1
    tPrmData.CsdUdpDtm = vVal(i)
    i = i + 1
    tPrmData.CsdOutYon = vVal(i)
    
    i = i + 1   '20030101 lek add
    tPrmData.CsdOspYon = vVal(i)
    
    
    
    Exit Sub

CsdInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CsdInfStore(sPrmKey As String, sPrmValue As String, tPrmData As CsdInfRec)

    
    sPrmKey = tPrmData.CsdAcpDte & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CsdDgsDEN & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CsdUsgPrt & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CsdDgsTbl & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.CsdSeq), "@@@") & Chr(5)
    
    sPrmValue = tPrmData.CsdDepTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsdCsrcod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsdCsmYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsdRqtQty & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsdOutQty & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsdUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsdUdpDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsdOutYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsdOspYon & Chr(5) '20030101 lek add
    
End Sub

    
Public Sub CstInfLoad(sPrmValue As String, tPrmData As CstInfRec)

    On Error GoTo CstInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.CstDepCod = vVal(i)
    i = i + 1
    tPrmData.CstOcmNum = vVal(i)
    i = i + 1
    tPrmData.CstAdpDte = vVal(i)
    i = i + 1
    tPrmData.CstChtNum = vVal(i)
    i = i + 1
    tPrmData.CstPatNam = vVal(i)
    i = i + 1
    tPrmData.CstNssCod = vVal(i)
    i = i + 1
    tPrmData.CstRomCod = vVal(i)
    i = i + 1
    tPrmData.CstExpDte = vVal(i)
    i = i + 1
    tPrmData.CstFrmDep = vVal(i)
    i = i + 1
    tPrmData.CstCurStt = vVal(i)
    i = i + 1
    tPrmData.CstCrtYon = vVal(i)
    i = i + 1
    tPrmData.CstGrpDep = vVal(i)
    i = i + 1
    tPrmData.CstSplCmt = vVal(i)
    i = i + 1
    tPrmData.CstRetCmt = vVal(i)
    
    Exit Sub

CstInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CstInfStore(sPrmKey As String, sPrmValue As String, tPrmData As CstInfRec)

    
    sPrmKey = tPrmData.CstDepCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.CstOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CstAdpDte & Chr(5)
    
    sPrmValue = Format((tPrmData.CstChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstNssCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstRomCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstFrmDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstCurStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstCrtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstGrpDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CstRetCmt & Chr(5)
    
End Sub

    
Public Sub DctInfLoad(sPrmValue As String, tPrmData As DctInfRec)

    On Error GoTo DctInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    '
    i = i + 1
    tPrmData.DctChtNum = vVal(i)
    i = i + 1
    tPrmData.DctUpdTim = vVal(i)
    i = i + 1
    tPrmData.DctUidCod = vVal(i)
    '
    Exit Sub

DctInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub DctInfStore(sPrmKey As String, sPrmValue As String, tPrmData As DctInfRec)

    
    sPrmKey = Format((tPrmData.DctChtNum), "@@@@@@@@") & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.DctUpdTim & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DctUidCod & Chr(5)
    
End Sub

    
Public Sub DlyInfLoad(sPrmValue As String, tPrmData As DlyInfRec)

    On Error GoTo DlyInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.DlyChtNum = vVal(i)
    i = i + 1
    tPrmData.DlyDepCod = vVal(i)
    i = i + 1
    tPrmData.DlyAcpDtm = vVal(i)
    i = i + 1
    tPrmData.DlyOrXray = vVal(i)
    i = i + 1
    tPrmData.DlyOutDep = vVal(i)
    i = i + 1
    tPrmData.DlyResDtm = vVal(i)
    i = i + 1
    tPrmData.DlyDlvNam = vVal(i)
    i = i + 1
    tPrmData.DlyOutUid = vVal(i)
    i = i + 1
    tPrmData.DlyRcvUid = vVal(i)
    i = i + 1
    tPrmData.DlyOutGrp = vVal(i)
    i = i + 1
    tPrmData.DlyMemo = vVal(i)
    i = i + 1
    tPrmData.DlyMemDep = vVal(i)
    i = i + 1
    tPrmData.DlyOcmNum = vVal(i)
    i = i + 1
    tPrmData.DlyRelFlg = vVal(i)
    i = i + 1
    tPrmData.DlyRemark = vVal(i)
    
    
    Exit Sub

DlyInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub DlyInfStore(sPrmKey As String, sPrmValue As String, tPrmData As DlyInfRec)

    
    sPrmKey = Format(tPrmData.DlyChtNum, "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.DlyDepCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.DlyAcpDtm & Chr(5)
    
    sPrmValue = tPrmData.DlyOrXray & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyOutDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyResDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyDlvNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyOutUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyRcvUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyOutGrp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyMemo & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyMemDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyOcmNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyRelFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DlyRemark & Chr(5)
    
End Sub

Public Sub DthInfLoad(sPrmValue As String, tPrmData As DthInfRec)

    On Error GoTo DthInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.DthChtNum = vVal(i)
    i = i + 1
    tPrmData.DthIoODgs = vVal(i)
    i = i + 1
    tPrmData.DthWatStr = vVal(i)
    i = i + 1
    tPrmData.DthWatEnd = vVal(i)
    i = i + 1
    tPrmData.DthDthStr = vVal(i)
    i = i + 1
    tPrmData.DthDthEnd = vVal(i)
    i = i + 1
    tPrmData.DthRefCmd = vVal(i)
    
    Exit Sub

DthInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub DthInfStore(sPrmKey As String, sPrmValue As String, tPrmData As DthInfRec)

    
    sPrmKey = Format(Trim(tPrmData.DthChtNum), "@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.DthIoODgs & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DthWatStr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DthWatEnd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DthDthStr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DthDthEnd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DthRefCmd & Chr(5)
    
End Sub

    
Public Sub DutInfLoad(sPrmValue As String, tPrmData As DutInfRec)

    On Error GoTo DutInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.DutDtrCod = vVal(i)
    i = i + 1
    tPrmData.DutSttDtm = vVal(i)
    
    i = i + 1
    tPrmData.DutEndDtm = vVal(i)
    i = i + 1
    tPrmData.DutOdrCmt = vVal(i)
    
    Exit Sub

DutInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub DutInfStore(sPrmKey As String, sPrmValue As String, tPrmData As DutInfRec)

    
    sPrmKey = tPrmData.DutDtrCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.DutSttDtm & Chr(5)
    
    sPrmValue = tPrmData.DutEndDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.DutOdrCmt & Chr(5)
    
End Sub

    

Public Sub EdcInfLoad(sPrmValue As String, tPrmData As EdcInfRec)

    On Error GoTo EdcInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.EdcLmpDte = vVal(i)
    i = i + 1
    tPrmData.EdcChtNum = vVal(i)
    i = i + 1
    tPrmData.EdcEdcDte = vVal(i)
    i = i + 1
    tPrmData.EdcAbrYon = vVal(i)
    i = i + 1
    tPrmData.EdcLbrYon = vVal(i)
    i = i + 1
    tPrmData.EdcLbrSeq = vVal(i)
    i = i + 1
    tPrmData.EdcHspLbr = vVal(i)
    
    Exit Sub

EdcInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub EdcInfStore(sPrmKey As String, sPrmValue As String, tPrmData As EdcInfRec)

    sPrmKey = tPrmData.EdcEdcDte & Chr(5)
    sPrmKey = sPrmKey & tPrmData.EdcChtNum & Chr(5)
    
    sPrmValue = tPrmData.EdcLmpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EdcAbrYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EdcLbrYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EdcLbrSeq & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EdcHspLbr & Chr(5)
        
End Sub


Public Sub EegInfLoad(sPrmValue As String, tPrmEegData As EegInfRec)

    On Error GoTo EegInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmEegData.EegChtNum = vVal(i)
    i = i + 1
    tPrmEegData.EegEegNum = vVal(i)
    i = i + 1
    tPrmEegData.EegUpdDte = vVal(i)
    
    Exit Sub

EegInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub EegInfStore(sPrmKey As String, sPrmValue As String, tPrmEegData As EegInfRec)

    
    sPrmKey = Format(Trim(tPrmEegData.EegChtNum), "@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(Trim(tPrmEegData.EegEegNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmEegData.EegUpdDte & Chr(5)
    
End Sub

    
Public Sub EtcInfLoad(sPrmValue As String, tPrmData As EtcInfRec)

    On Error GoTo EtcInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.EtcOcmNum = vVal(i)
    i = i + 1
    tPrmData.EtcChtNum = vVal(i)
    i = i + 1
    tPrmData.EtcPatNam = vVal(i)
    i = i + 1
    tPrmData.EtcResNum = vVal(i)
    i = i + 1
    tPrmData.EtcInsCod = vVal(i)
    i = i + 1
    tPrmData.EtcInsSeq = vVal(i)
    i = i + 1
    tPrmData.EtcDepCod = vVal(i)
    i = i + 1
    tPrmData.EtcAssCod = vVal(i)
    i = i + 1
    tPrmData.EtcOdrDtm = vVal(i)
    i = i + 1
    tPrmData.EtcEntNam = vVal(i)
    i = i + 1
    tPrmData.EtcSplCmt = vVal(i)
    i = i + 1
    tPrmData.EtcUidCod = vVal(i)
    i = i + 1
    tPrmData.EtcGbnCod = vVal(i)
    i = i + 1
    tPrmData.EtcSpcYon = vVal(i)
    i = i + 1
    tPrmData.EtcTotAmt = vVal(i)
    i = i + 1
    tPrmData.EtcMmsEtc = vVal(i)
    i = i + 1
    tPrmData.EtcGbnDtl = vVal(i)
    
    '2002.03.04 sebal
    i = i + 1
    tPrmData.EtcTelNum = vVal(i)
    i = i + 1
    tPrmData.EtcZipCod = vVal(i)
    i = i + 1
    tPrmData.EtcAddRes = vVal(i)
    i = i + 1
    tPrmData.EtcCalTel = vVal(i)
    i = i + 1
    tPrmData.EtcCalZip = vVal(i)
    i = i + 1
    tPrmData.EtcCalAdd = vVal(i)
    i = i + 1
    tPrmData.EtcJobNam = vVal(i)
    i = i + 1
    tPrmData.EtcHndPhn = vVal(i)
    i = i + 1
    tPrmData.EtcE_Mail = vVal(i)
    Exit Sub

EtcInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub EtcInfStore(sPrmKey As String, sPrmValue As String, tPrmData As EtcInfRec)

    
    sPrmKey = Format(Trim(tPrmData.EtcOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(Trim(tPrmData.EtcChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcResNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcInsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcInsSeq & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcAssCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcOdrDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcEntNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcGbnCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcTotAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcMmsEtc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcGbnDtl & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.EtcTelNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcZipCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcAddRes & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcCalTel & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcCalZip & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcCalAdd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcJobNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcHndPhn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.EtcE_Mail & Chr(5)
    
End Sub

    
Public Sub FhtInfLoad(sPrmValue As String, tPrmData As FhtInfRec)

    On Error GoTo FhtInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.FhtRcpNum = vVal(i)
    i = i + 1
    tPrmData.FhtFutTyp = vVal(i)
    i = i + 1
    tPrmData.FhtChtNum = vVal(i)
    i = i + 1
    tPrmData.FhtFutOld = vVal(i)
    i = i + 1
    tPrmData.FhtOcmNum = vVal(i)
    i = i + 1
    tPrmData.FhtRvnTyp = vVal(i)
    i = i + 1
    tPrmData.FhtOcmSeq = vVal(i)
    i = i + 1
    tPrmData.FhtDupSeq = vVal(i)
    i = i + 1
    tPrmData.FhtFutSts = vVal(i)
    i = i + 1
    tPrmData.FhtPatTyp = vVal(i)
    i = i + 1
    tPrmData.FhtInsCod = vVal(i)
    i = i + 1
    tPrmData.FhtDepCod = vVal(i)
    i = i + 1
    tPrmData.FhtSerNum = vVal(i)
    i = i + 1
    tPrmData.FhtOcrAmt = vVal(i)
    i = i + 1
    tPrmData.FhtDisAmt = vVal(i)
    i = i + 1
    tPrmData.FhtPayAmt = vVal(i)
    i = i + 1
    tPrmData.FhtRemAmt = vVal(i)
    i = i + 1
    tPrmData.FhtCorAmt = vVal(i)
    i = i + 1
    tPrmData.FhtRcpOld = vVal(i)
    i = i + 1
    tPrmData.FhtOldNum = vVal(i)
    i = i + 1
    tPrmData.FhtOcrRsn = vVal(i)
    i = i + 1
    tPrmData.FhtPrcDtm = vVal(i)
    i = i + 1
    tPrmData.FhtUidCod = vVal(i)
    i = i + 1
    tPrmData.FhtDisYon = vVal(i)
    
    Exit Sub

FhtInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub FhtInfStore(sPrmKey As String, sPrmValue As String, tPrmData As FhtInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.FhtRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.FhtFutTyp & Chr(5)
    
    sPrmValue = Format((tPrmData.FhtChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtFutOld & Chr(5)
    sPrmValue = sPrmValue & Format((tPrmData.FhtOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtRvnTyp & Chr(5)
    sPrmValue = sPrmValue & Format(tPrmData.FhtOcmSeq, "@@") & Chr(5)
    sPrmValue = sPrmValue & Format(tPrmData.FhtDupSeq, "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtFutSts & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtPatTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtInsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtSerNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtOcrAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtDisAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtPayAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtRemAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtCorAmt & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.FhtRcpOld), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.FhtOldNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtOcrRsn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtPrcDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FhtDisYon & Chr(5)
    
End Sub

    
Public Sub ForInfLoad(sPrmValue As String, tPrmData As ForInfRec)

    On Error GoTo ForInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.ForChtNum = vVal(i)
    i = i + 1
    tPrmData.ForFutTyp = vVal(i)
    i = i + 1
    tPrmData.ForOcmNum = vVal(i)
    i = i + 1
    tPrmData.ForRvnTyp = vVal(i)
    i = i + 1
    tPrmData.ForOcmSeq = vVal(i)
    i = i + 1
    tPrmData.ForDupSeq = vVal(i)
    i = i + 1
    tPrmData.ForFutSts = vVal(i)
    i = i + 1
    tPrmData.ForPatTyp = vVal(i)
    i = i + 1
    tPrmData.ForInsCod = vVal(i)
    i = i + 1
    tPrmData.ForDepCod = vVal(i)
    i = i + 1
    tPrmData.ForSerNum = vVal(i)
    i = i + 1
    tPrmData.ForOcrAmt = vVal(i)
    i = i + 1
    tPrmData.ForDisAmt = vVal(i)
    i = i + 1
    tPrmData.ForPayAmt = vVal(i)
    i = i + 1
    tPrmData.ForRemAmt = vVal(i)
    i = i + 1
    tPrmData.ForCorAmt = vVal(i)
    i = i + 1
    tPrmData.ForRcpNum = vVal(i)
    i = i + 1
    tPrmData.ForOldNum = vVal(i)
    i = i + 1
    tPrmData.ForOcrRsn = vVal(i)
    i = i + 1
    tPrmData.ForPrcDtm = vVal(i)
    i = i + 1
    tPrmData.ForUidCod = vVal(i)
    i = i + 1
    tPrmData.ForDisYon = vVal(i)
    
    Exit Sub

ForInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ForInfStore(sPrmKey As String, sPrmValue As String, tPrmData As ForInfRec)

    
    sPrmKey = Format((tPrmData.ForChtNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ForFutTyp & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.ForOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ForRvnTyp & Chr(5)
    sPrmKey = sPrmKey & Format((tPrmData.ForOcmSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format((tPrmData.ForDupSeq), "@@") & Chr(5)
    
    sPrmValue = tPrmData.ForFutSts & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForPatTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForInsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForSerNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForOcrAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForDisAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForPayAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForRemAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForCorAmt & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.ForRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.ForOldNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForOcrRsn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForPrcDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ForDisYon & Chr(5)
    
End Sub

    
Public Sub FpaInfLoad(sPrmValue As String, tPrmData As FpaInfRec)

    On Error GoTo FpaInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.FpaChtNum = vVal(i)
    i = i + 1
    tPrmData.FpaPaySeq = vVal(i)
    i = i + 1
    tPrmData.FpaPayAmt = vVal(i)
    i = i + 1
    tPrmData.FpaRemDsc = vVal(i)
    i = i + 1
    tPrmData.FpaPrcDtm = vVal(i)
    i = i + 1
    tPrmData.FpaUidCod = vVal(i)
    i = i + 1
    tPrmData.FpaRcpNum = vVal(i)
    
    Exit Sub

FpaInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub FpaInfStore(sPrmKey As String, sPrmValue As String, tPrmFpaData As FpaInfRec)

    
    sPrmKey = tPrmFpaData.FpaChtNum & Chr(5)
    sPrmKey = sPrmKey & tPrmFpaData.FpaPaySeq & Chr(5)
    
    sPrmValue = tPrmFpaData.FpaPayAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFpaData.FpaRemDsc & Chr(5)
    sPrmValue = sPrmValue & tPrmFpaData.FpaPrcDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmFpaData.FpaUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmFpaData.FpaRcpNum & Chr(5)
    
End Sub

    
Public Sub FutInfLoad(sPrmValue As String, tPrmData As FutInfRec)

    On Error GoTo FutInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.FutChtNum = vVal(i)
    i = i + 1
    tPrmData.FutCurSts = vVal(i)
    i = i + 1
    tPrmData.FutFutAmt = vVal(i)
    i = i + 1
    tPrmData.FutPayAmt = vVal(i)
    i = i + 1
    tPrmData.FutDisAmt = vVal(i)
    i = i + 1
    tPrmData.FutRemAmt = vVal(i)
    i = i + 1
    tPrmData.FutEmpNum = vVal(i)
    i = i + 1
    tPrmData.FutStrDte = vVal(i)
    i = i + 1
    tPrmData.FutEndDte = vVal(i)
    i = i + 1
    tPrmData.FutFutRsn = vVal(i)
    i = i + 1
    tPrmData.FutExpDte = vVal(i)
    
    Exit Sub

FutInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub FutInfStore(sPrmKey As String, sPrmValue As String, tPrmData As FutInfRec)

    
    sPrmKey = Format((tPrmData.FutChtNum), "@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.FutCurSts & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutFutAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutPayAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutDisAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutRemAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutEmpNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutStrDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutEndDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutFutRsn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.FutExpDte & Chr(5)
End Sub

    
Public Sub IcmInfLoad(sPrmValue As String, tPrmData As IcmInfRec)

    On Error GoTo IcmInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    With tPrmData
        i = i + 1
        .IcmOcmNum = vVal(i)
        i = i + 1
        .IcmChtNum = vVal(i)
        i = i + 1
        .IcmAcpStt = vVal(i)
        i = i + 1
        .IcmAcpDtm = vVal(i)
        i = i + 1
        .IcmArrPat = vVal(i)
        i = i + 1
        .IcmAcpRut = vVal(i)
        i = i + 1
        .IcmDepCod = vVal(i)
        i = i + 1
        .IcmDtrCod = vVal(i)
        i = i + 1
        .IcmInsCod = vVal(i)
        i = i + 1
        .IcmInsSeq = vVal(i)
        i = i + 1
        .IcmDupYon = vVal(i)
        i = i + 1
        .IcmConYon = vVal(i)
        i = i + 1
        .IcmNssCod = vVal(i)
        i = i + 1
        .IcmRomCod = vVal(i)
        i = i + 1
        .IcmBedCod = vVal(i)
        i = i + 1
        .IcmLevCnt = vVal(i)
        i = i + 1
        .IcmLevDtm = vVal(i)
        i = i + 1
        .IcmNtcDtm = vVal(i)
        i = i + 1
        .IcmOutDtm = vVal(i)
        i = i + 1
        .IcmRtnDtm = vVal(i)
        i = i + 1
        .IcmDgsNfs = vVal(i)
        i = i + 1
        .IcmDgsDnh = vVal(i)
        i = i + 1
        .IcmSpcYon = vVal(i)
        i = i + 1
        .IcmUpdDtm = vVal(i)
        i = i + 1
        .icmUidCod = vVal(i)
        i = i + 1
        .IcmNonIns = vVal(i)
        i = i + 1
        .IcmImgYon = vVal(i)
        i = i + 1
        .IcmRemark = vVal(i)
        i = i + 1
        .IcmRcpCmt = vVal(i)
        i = i + 1
        .IcmMomCht = vVal(i)
        i = i + 1
        .IcmRptDte = vVal(i)
        i = i + 1
        .IcmPreSts = vVal(i)
        i = i + 1
        .IcmPreDtm = vVal(i)
        i = i + 1
        .IcmOdrDtm = vVal(i)
        i = i + 1
        .IcmCfmYon = vVal(i)
        i = i + 1
        .IcmIspBak = vVal(i)
        i = i + 1
        .IcmSimDtm = vVal(i)
        i = i + 1
        .IcmPedOcm = vVal(i)
        i = i + 1
        .IcmPedDtm = vVal(i)
    
    End With
    
    Exit Sub

IcmInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IcmInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IcmInfRec)

    With tPrmData
        sPrmKey = Format(CDouble(.IcmOcmNum), "@@@@@@@@@@") & Chr(5)
            
        sPrmValue = Format((.IcmChtNum), "@@@@@@@@") & Chr(5)
        sPrmValue = sPrmValue & .IcmAcpStt & Chr(5)
        sPrmValue = sPrmValue & .IcmAcpDtm & Chr(5)
        sPrmValue = sPrmValue & .IcmArrPat & Chr(5)
        sPrmValue = sPrmValue & .IcmAcpRut & Chr(5)
        sPrmValue = sPrmValue & .IcmDepCod & Chr(5)
        sPrmValue = sPrmValue & .IcmDtrCod & Chr(5)
        sPrmValue = sPrmValue & .IcmInsCod & Chr(5)
        sPrmValue = sPrmValue & Format(CDouble(.IcmInsSeq), "@@") & Chr(5)
        sPrmValue = sPrmValue & .IcmDupYon & Chr(5)
        sPrmValue = sPrmValue & .IcmConYon & Chr(5)
        sPrmValue = sPrmValue & .IcmNssCod & Chr(5)
        sPrmValue = sPrmValue & .IcmRomCod & Chr(5)
        sPrmValue = sPrmValue & .IcmBedCod & Chr(5)
        sPrmValue = sPrmValue & .IcmLevCnt & Chr(5)
        sPrmValue = sPrmValue & .IcmLevDtm & Chr(5)
        sPrmValue = sPrmValue & .IcmNtcDtm & Chr(5)
        sPrmValue = sPrmValue & .IcmOutDtm & Chr(5)
        sPrmValue = sPrmValue & .IcmRtnDtm & Chr(5)
        sPrmValue = sPrmValue & .IcmDgsNfs & Chr(5)
        sPrmValue = sPrmValue & .IcmDgsDnh & Chr(5)
        sPrmValue = sPrmValue & .IcmSpcYon & Chr(5)
        sPrmValue = sPrmValue & .IcmUpdDtm & Chr(5)
        sPrmValue = sPrmValue & .icmUidCod & Chr(5)
        sPrmValue = sPrmValue & .IcmNonIns & Chr(5)
        sPrmValue = sPrmValue & .IcmImgYon & Chr(5)
        sPrmValue = sPrmValue & .IcmRemark & Chr(5)
        sPrmValue = sPrmValue & .IcmRcpCmt & Chr(5)
        sPrmValue = sPrmValue & .IcmMomCht & Chr(5)
        sPrmValue = sPrmValue & .IcmRptDte & Chr(5)
        sPrmValue = sPrmValue & .IcmPreSts & Chr(5)
        sPrmValue = sPrmValue & .IcmPreDtm & Chr(5)
        sPrmValue = sPrmValue & .IcmOdrDtm & Chr(5)
        sPrmValue = sPrmValue & .IcmCfmYon & Chr(5)
        sPrmValue = sPrmValue & .IcmIspBak & Chr(5)
        sPrmValue = sPrmValue & .IcmSimDtm & Chr(5)
        sPrmValue = sPrmValue & Format(CDouble(.IcmPedOcm), "@@@@@@@@@@") & Chr(5)
        sPrmValue = sPrmValue & .IcmPedDtm & Chr(5)
    End With
    
End Sub

    
Public Sub IcrInfLoad(sPrmValue As String, tPrmData As IcrInfRec)

    On Error GoTo IcrInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.IcrHopDtm = vVal(i)
    i = i + 1
    tPrmData.IcrChtNum = vVal(i)
    
    i = i + 1
    tPrmData.IcrOcmNum = vVal(i)
    
    i = i + 1
    tPrmData.IcrCurDep = vVal(i)
    i = i + 1
    tPrmData.IcrCurDtr = vVal(i)
    i = i + 1
    tPrmData.IcrCurNss = vVal(i)
    i = i + 1
    tPrmData.IcrCurRom = vVal(i)
    i = i + 1
    tPrmData.IcrCurBed = vVal(i)
    i = i + 1
    tPrmData.IcrCurGrd = vVal(i)
    
    i = i + 1
    tPrmData.IcrHopDep = vVal(i)
    i = i + 1
    tPrmData.IcrHopDtr = vVal(i)
    i = i + 1
    tPrmData.IcrHopNss = vVal(i)
    i = i + 1
    tPrmData.IcrHopRom = vVal(i)
    i = i + 1
    tPrmData.IcrHopBed = vVal(i)
    i = i + 1
    tPrmData.IcrHopGrd = vVal(i)
    
    i = i + 1
    tPrmData.IcrTrsDtm = vVal(i)
    
    i = i + 1
    tPrmData.IcrNssYon = vVal(i)
    i = i + 1
    tPrmData.IcrNssUid = vVal(i)
    
    i = i + 1
    tPrmData.IcrWonYon = vVal(i)
    i = i + 1
    tPrmData.IcrWonUid = vVal(i)
    
    Exit Sub

IcrInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IcrInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IcrInfRec)

    
    sPrmKey = tPrmData.IcrHopDtm & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmData.IcrChtNum, "@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(tPrmData.IcrOcmNum, "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.IcrCurDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrCurDtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrCurNss & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrCurRom & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrCurBed & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrCurGrd & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.IcrHopDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrHopDtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrHopNss & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrHopRom & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrHopBed & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrHopGrd & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.IcrTrsDtm & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.IcrNssYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrNssUid & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.IcrWonYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IcrWonUid & Chr(5)
    
End Sub

    
Public Sub IdiInfLoad(sPrmValue As String, tPrmIdiData As IdiInfRec)

    On Error GoTo IdiInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmIdiData.IdiOcmNum = vVal(i)
    i = i + 1
    tPrmIdiData.IdiFrmDte = vVal(i)
    i = i + 1
    tPrmIdiData.IdiFrmTyp = vVal(i)
    i = i + 1
    tPrmIdiData.IdiFeeCod = vVal(i)
    i = i + 1
    tPrmIdiData.IdiCalTyp = vVal(i)
    i = i + 1
    tPrmIdiData.IdiEndDte = vVal(i)
    i = i + 1
    tPrmIdiData.IdiEndTyp = vVal(i)
    i = i + 1
    tPrmIdiData.IdiNssCod = vVal(i)
    i = i + 1
    tPrmIdiData.IdiRomCod = vVal(i)
    i = i + 1
    tPrmIdiData.IdiBedCod = vVal(i)
    i = i + 1
    tPrmIdiData.IdiAddCod = vVal(i)
    i = i + 1
    tPrmIdiData.IdiDepCod = vVal(i)
    i = i + 1
    tPrmIdiData.IdiDtrCod = vVal(i)
    i = i + 1
    tPrmIdiData.IdiWhyCod = vVal(i)
    i = i + 1
    tPrmIdiData.IdiUpdDte = vVal(i)
    i = i + 1
    tPrmIdiData.IdiUidCod = vVal(i)
    
    Exit Sub

IdiInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IdiInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IdiInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.IdiOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.IdiFrmDte & Chr(5)
    sPrmKey = sPrmKey & tPrmData.IdiFrmTyp & Chr(5)
    
    sPrmValue = tPrmData.IdiFeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiCalTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiEndDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiEndTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiNssCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiRomCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiBedCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiAddCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiWhyCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiUpdDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdiUidCod & Chr(5)
End Sub

    
Public Sub IdaInfLoad(sPrmValue As String, tPrmData As IdaInfRec)

    On Error GoTo IdlInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.IdaOcmNum = vVal(i)
    i = i + 1
    tPrmData.IdaDupSeq = vVal(i)
    i = i + 1
    tPrmData.IdaOdrDte = vVal(i)
    i = i + 1
    tPrmData.IdaDepCod = vVal(i)
    i = i + 1
    tPrmData.IdaItmCod = vVal(i)
    i = i + 1
    tPrmData.IdaAstCod = vVal(i)
    i = i + 1
    tPrmData.IdaChtNum = vVal(i)
    i = i + 1
    tPrmData.IdaInsMat = vVal(i)
    i = i + 1
    tPrmData.IdaInsAct = vVal(i)
    i = i + 1
    tPrmData.IdaNonMat = vVal(i)
    i = i + 1
    tPrmData.IdaNonAct = vVal(i)
    i = i + 1
    tPrmData.IdaSpcAmt = vVal(i)
    i = i + 1
    tPrmData.IdaUpdDtm = vVal(i)
    i = i + 1
    tPrmData.IdaUidCod = vVal(i)
    
    Exit Sub

IdlInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IdaInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IdaInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.IdaOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IdaDupSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.IdaOdrDte & Chr(5)
    sPrmKey = sPrmKey & tPrmData.IdaDepCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.IdaItmCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.IdaAstCod & Chr(5)
    
    sPrmValue = Format(tPrmData.IdaChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdaInsMat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdaInsAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdaNonMat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdaNonAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdaSpcAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdaUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdaUidCod & Chr(5)
    
End Sub
    
Public Sub IdlInfLoad(sPrmValue As String, tPrmData As IdlInfRec)

    On Error GoTo IdlInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.IdlRcpNum = vVal(i)
    i = i + 1
    tPrmData.IdlIncCod = vVal(i)
    i = i + 1
    tPrmData.IdlChtNum = vVal(i)
    i = i + 1
    tPrmData.IdlInsCod = vVal(i)
    i = i + 1
    tPrmData.IdlInsSeq = vVal(i)
    i = i + 1
    tPrmData.IdlDepCod = vVal(i)
    i = i + 1
    tPrmData.IdlDtrCod = vVal(i)
    i = i + 1
    tPrmData.IdlInsAct = vVal(i)
    i = i + 1
    tPrmData.IdlInsMat = vVal(i)
    i = i + 1
    tPrmData.IdlNonAct = vVal(i)
    i = i + 1
    tPrmData.IdlNonMat = vVal(i)
    i = i + 1
    tPrmData.IdlInsAmt = vVal(i)
    i = i + 1
    tPrmData.IdlNonAmt = vVal(i)
    i = i + 1
    tPrmData.IdlInsOwn = vVal(i)
    i = i + 1
    tPrmData.IdlTotOwn = vVal(i)
    i = i + 1
    tPrmData.IdlSpcAmt = vVal(i)
'--------------------------------------> 추가
    '20040102..HTS..add
    i = i + 1
    tPrmData.IdlNinAct = vVal(i)
    i = i + 1
    tPrmData.IdlNinMat = vVal(i)
    i = i + 1
    tPrmData.IdlNinAmt = vVal(i)
'--------------------------------------> 추가
    Exit Sub

IdlInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IdlInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IdlInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.IdlRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IdlIncCod), "@@") & Chr(5)
    
    sPrmValue = Format(tPrmData.IdlChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IdlInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlInsAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlInsMat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlNonAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlNonMat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlNonAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlInsOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlTotOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlSpcAmt & Chr(5)
'--------------------------------------> 추가
    '20040102..HTS..add
    sPrmValue = sPrmValue & tPrmData.IdlNinAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlNinMat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IdlNinAmt & Chr(5)
'--------------------------------------> 추가
End Sub

    
Public Sub IhtInfLoad(sPrmValue As String, tPrmData As IhtInfRec)

    On Error GoTo IhtInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.IhtRcpNum = vVal(i)
    i = i + 1
    tPrmData.IhtOcmNum = vVal(i)
    i = i + 1
    tPrmData.IhtOcmSeq = vVal(i)
    i = i + 1
    tPrmData.IhtDupSeq = vVal(i)
    i = i + 1
    tPrmData.IhtDepcod = vVal(i)
    i = i + 1
    tPrmData.IhtChtNum = vVal(i)
    i = i + 1
    tPrmData.IhtDtrCod = vVal(i)
    i = i + 1
    tPrmData.IhtInsCod = vVal(i)
    i = i + 1
    tPrmData.IhtInsSeq = vVal(i)
    i = i + 1
    tPrmData.IhtTotAmt = vVal(i)
    i = i + 1
    tPrmData.IhtCorAmt = vVal(i)
    i = i + 1
    tPrmData.IhtNonAmt = vVal(i)
    i = i + 1
    tPrmData.IhtOwnAmt = vVal(i)
    i = i + 1
    tPrmData.IhtTotOwn = vVal(i)
    i = i + 1
    tPrmData.IhtInsAmt = vVal(i)
    i = i + 1
    tPrmData.IhtSpcAmt = vVal(i)
    i = i + 1
    tPrmData.IhtAskAmt = vVal(i)
    i = i + 1
    tPrmData.IhtDisAmt = vVal(i)
    i = i + 1
    tPrmData.IhtFutAmt = vVal(i)
    i = i + 1
    tPrmData.IhtOldAmt = vVal(i)
    i = i + 1
    tPrmData.IhtNewAmt = vVal(i)
    i = i + 1
    tPrmData.IhtGrnAmt = vVal(i)
    i = i + 1
    tPrmData.IhtRcpCur = vVal(i)
    i = i + 1
    tPrmData.IhtOldNum = vVal(i)
    i = i + 1
    tPrmData.IhtCalDte = vVal(i)
    i = i + 1
    tPrmData.IhtUpdDtm = vVal(i)
    i = i + 1
    tPrmData.IhtUidCod = vVal(i)
    i = i + 1
    tPrmData.IhtOvrAmt = vVal(i)
    i = i + 1
    tPrmData.IhtRcpFlg = vVal(i)
    i = i + 1
    tPrmData.IhtDimAmt = vVal(i)
    i = i + 1
    tPrmData.IhtNonIns = vVal(i)
    i = i + 1
    tPrmData.IhtCarFut = vVal(i)
    i = i + 1
    tPrmData.IhtFodAmt = vVal(i)
'--------------------------------------> 추가
    '20040102..HTS..Add
    i = i + 1
    tPrmData.IhtNinAmt = vVal(i)
'--------------------------------------> 추가
    Exit Sub

IhtInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IhtInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IhtInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.IhtRcpNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(CDouble(tPrmData.IhtOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IhtOcmSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IhtDupSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtDepcod & Chr(5)
    sPrmValue = sPrmValue & Format(tPrmData.IhtChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IhtInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtTotAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtCorAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtNonAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtOwnAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtTotOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtSpcAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtAskAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtDisAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtFutAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtOldAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtNewAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtGrnAmt & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IhtRcpCur), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IhtOldNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtCalDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtOvrAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtRcpFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtDimAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtNonIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtCarFut & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IhtFodAmt & Chr(5)
'--------------------------------------> 추가
    '20040102..HTS..add
    sPrmValue = sPrmValue & tPrmData.IhtNinAmt & Chr(5)
'--------------------------------------> 추가
End Sub

    
Public Sub IisHstLoad(sPrmValue As String, tPrmData As IisHstRec)

    On Error GoTo IisHstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.IisOcmNum = vVal(i)
    i = i + 1
    tPrmData.IisOcmSeq = vVal(i)
    i = i + 1
    tPrmData.IisDupSeq = vVal(i)
    i = i + 1
    tPrmData.IisDelDtm = vVal(i)
    i = i + 1
    tPrmData.IisChtNum = vVal(i)
    i = i + 1
    tPrmData.IisInsCod = vVal(i)
    i = i + 1
    tPrmData.IisInsSeq = vVal(i)
    i = i + 1
    tPrmData.IisSpcYon = vVal(i)
    i = i + 1
    tPrmData.IisArtYon = vVal(i)
    i = i + 1
    tPrmData.IisAdpDte = vVal(i)
    i = i + 1
    tPrmData.IisExpDte = vVal(i)
    i = i + 1
    tPrmData.IisAcpDay = vVal(i)
    i = i + 1
    tPrmData.IisIcuDay = vVal(i)
    i = i + 1
    tPrmData.IisRcpYon = vVal(i)
    i = i + 1
    tPrmData.IisRcpDtm = vVal(i)
    i = i + 1
    tPrmData.IisDepCod = vVal(i)
    i = i + 1
    tPrmData.IisDtrCod = vVal(i)
    i = i + 1
    tPrmData.IisNonIns = vVal(i)
    i = i + 1
    tPrmData.IisBilYon = vVal(i)
    i = i + 1
    tPrmData.IisDrgYon = vVal(i)
    i = i + 1
    tPrmData.IisUidCod = vVal(i)
    i = i + 1
    tPrmData.IisDrgCod = vVal(i)
    i = i + 1
    tPrmData.IisDrgDay = vVal(i)
    i = i + 1
    tPrmData.IisUpdDtm = vVal(i)
    i = i + 1
    tPrmData.IisActTyp = vVal(i)
    i = i + 1
    tPrmData.IisCstYon = vVal(i)
    i = i + 1
    tPrmData.IisLmtAmt = vVal(i)
    
    
        
    Exit Sub

IisHstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IisHstStore(sPrmKey As String, sPrmValue As String, tPrmData As IisHstRec)

    sPrmKey = Format(CDouble(tPrmData.IisOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IisOcmSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IisDupSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IisDelDtm), "@@") & Chr(5)
    
    sPrmValue = Format(tPrmData.IisChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IisInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisArtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisAcpDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisIcuDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisRcpYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisRcpDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisNonIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisBilYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDrgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDrgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDrgDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisActTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisCstYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisLmtAmt & Chr(5)
End Sub

    
Public Sub IisInfLoad(sPrmValue As String, tPrmData As IisInfRec)

    On Error GoTo IisInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.IisOcmNum = vVal(i)
    i = i + 1
    tPrmData.IisOcmSeq = vVal(i)
    i = i + 1
    tPrmData.IisDupSeq = vVal(i)
    i = i + 1
    tPrmData.IisChtNum = vVal(i)
    i = i + 1
    tPrmData.IisInsCod = vVal(i)
    i = i + 1
    tPrmData.IisInsSeq = vVal(i)
    i = i + 1
    tPrmData.IisSpcYon = vVal(i)
    i = i + 1
    tPrmData.IisArtYon = vVal(i)
    i = i + 1
    tPrmData.IisAdpDte = vVal(i)
    i = i + 1
    tPrmData.IisExpDte = vVal(i)
    i = i + 1
    tPrmData.IisAcpDay = vVal(i)
    i = i + 1
    tPrmData.IisIcuDay = vVal(i)
    i = i + 1
    tPrmData.IisRcpYon = vVal(i)
    i = i + 1
    tPrmData.IisRcpDtm = vVal(i)
    i = i + 1
    tPrmData.IisDepCod = vVal(i)
    i = i + 1
    tPrmData.IisDtrCod = vVal(i)
    i = i + 1
    tPrmData.IisNonIns = vVal(i)
    i = i + 1
    tPrmData.IisBilYon = vVal(i)
    i = i + 1
    tPrmData.IisDrgYon = vVal(i)
    i = i + 1
    tPrmData.IisUidCod = vVal(i)
    i = i + 1
    tPrmData.IisDrgCod = vVal(i)
    i = i + 1
    tPrmData.IisDrgDay = vVal(i)
    i = i + 1
    tPrmData.IisUpdDtm = vVal(i)
    i = i + 1
    tPrmData.IisCstYon = vVal(i)
    i = i + 1
    tPrmData.IisLmtAmt = vVal(i)
    
    
    Exit Sub

IisInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IisInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IisInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.IisOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IisOcmSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IisDupSeq), "@@") & Chr(5)
    
    sPrmValue = Format(tPrmData.IisChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IisInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisArtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisAcpDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisIcuDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisRcpYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisRcpDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisNonIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisBilYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDrgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDrgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisDrgDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisCstYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IisLmtAmt & Chr(5)
End Sub

    
Public Sub IloInfLoad(sPrmValue As String, tPrmData As IloInfRec)

    On Error GoTo IloInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.IloUpdDtm = vVal(i)
    i = i + 1
    tPrmData.IloSavSeq = vVal(i)
    i = i + 1
    tPrmData.IloOcmNum = vVal(i)
    i = i + 1
    tPrmData.IloChtNum = vVal(i)
    i = i + 1
    tPrmData.IloFrmIns = vVal(i)
    i = i + 1
    tPrmData.IloFrmDep = vVal(i)
    i = i + 1
    tPrmData.IloFrmDtr = vVal(i)
    i = i + 1
    tPrmData.IloFrmNss = vVal(i)
    i = i + 1
    tPrmData.IloFrmRom = vVal(i)
    i = i + 1
    tPrmData.IloFrmBed = vVal(i)
    i = i + 1
    tPrmData.IloToIns = vVal(i)
    i = i + 1
    tPrmData.IloToDep = vVal(i)
    i = i + 1
    tPrmData.IloToDtr = vVal(i)
    i = i + 1
    tPrmData.IloToNss = vVal(i)
    i = i + 1
    tPrmData.IloToRom = vVal(i)
    i = i + 1
    tPrmData.IloToBed = vVal(i)
    i = i + 1
    tPrmData.IloFunCod = vVal(i)
    
    Exit Sub

IloInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IloInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IloInfRec)

    
    sPrmKey = tPrmData.IloUpdDtm & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IloSavSeq), "@@") & Chr(5)
    
    sPrmValue = Format(CDouble(tPrmData.IloOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IloChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloFrmIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloFrmDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloFrmDtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloFrmNss & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloFrmRom & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloFrmBed & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloToIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloToDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloToDtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloToNss & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloToRom & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloToBed & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IloFunCod & Chr(5)
    
End Sub

    
Public Sub ImgInfLoad(sPrmValue As String, tPrmData As ImgInfRec)

    On Error GoTo ImgInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.ImgStrDtm = vVal(i)
    i = i + 1
    tPrmData.ImgEndDtm = vVal(i)
    i = i + 1
    tPrmData.ImgPatTyp = vVal(i)
    i = i + 1
    tPrmData.ImgSavYon = vVal(i)
    i = i + 1
    tPrmData.ImgCalYon = vVal(i)
    i = i + 1
    tPrmData.ImgDrgYon = vVal(i)
    i = i + 1
    tPrmData.ImgAccYon = vVal(i)
    i = i + 1
    tPrmData.ImgSavOut = vVal(i)
    i = i + 1
    tPrmData.ImgCalOut = vVal(i)
    i = i + 1
    tPrmData.ImgAccOut = vVal(i)
        
    Exit Sub

ImgInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ImgInfStore(sPrmKey As String, sPrmValue As String, tPrmData As ImgInfRec)

    
    sPrmKey = tPrmData.ImgStrDtm & Chr(5)
    
    sPrmValue = tPrmData.ImgEndDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ImgPatTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ImgSavYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ImgCalYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ImgDrgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ImgAccYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ImgSavOut & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ImgCalOut & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ImgAccOut & Chr(5)
    
End Sub

    
Public Sub ImlInfHstLoad(sPrmValue As String, tPrmImlData As ImlHstRec)

    On Error GoTo ImlInfHstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmImlData.ImlOcmNum = vVal(i)
    i = i + 1
    tPrmImlData.ImlUpdDtm = vVal(i)
    i = i + 1
    tPrmImlData.ImlSerNum = vVal(i)
    i = i + 1
    tPrmImlData.ImlAdpDte = vVal(i)
    i = i + 1
    tPrmImlData.ImlPatTyp = vVal(i)
    i = i + 1
    tPrmImlData.ImlExpDte = vVal(i)
    i = i + 1
    tPrmImlData.ImlBrfCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlBrfQty = vVal(i)
    i = i + 1
    tPrmImlData.ImlBrfCal = vVal(i)
    i = i + 1
    tPrmImlData.ImlBrfcc = vVal(i)
    i = i + 1
    tPrmImlData.ImlLnhCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlLnhQty = vVal(i)
    i = i + 1
    tPrmImlData.ImlLnhCal = vVal(i)
    i = i + 1
    tPrmImlData.ImlLnhcc = vVal(i)
    i = i + 1
    tPrmImlData.ImlDnrCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlDnrQty = vVal(i)
    i = i + 1
    tPrmImlData.ImlDnrCal = vVal(i)
    i = i + 1
    tPrmImlData.ImlDnrcc = vVal(i)
    i = i + 1
    tPrmImlData.ImlSplCmt = vVal(i)
    i = i + 1
    tPrmImlData.ImlWhyCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlUidCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlEntDtm = vVal(i)
    i = i + 1
    tPrmImlData.ImlEtcCmt = vVal(i)
    i = i + 1
    tPrmImlData.ImlUpdUid = vVal(i)
    
    
    Exit Sub

ImlInfHstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ImlInfHstStore(sPrmKey As String, sPrmValue As String, tPrmImlData As ImlHstRec)

    
    sPrmKey = tPrmImlData.ImlOcmNum & Chr(5)
    sPrmKey = sPrmKey & tPrmImlData.ImlUpdDtm & Chr(5)
    sPrmKey = sPrmKey & tPrmImlData.ImlSerNum & Chr(5)
    sPrmKey = sPrmKey & tPrmImlData.ImlAdpDte & Chr(5)
    
    sPrmValue = tPrmImlData.ImlPatTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlBrfCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlBrfQty & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlBrfCal & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlBrfcc & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlLnhCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlLnhQty & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlLnhCal & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlLnhcc & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlDnrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlDnrQty & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlDnrCal & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlDnrcc & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlWhyCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlEntDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlEtcCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlUpdUid & Chr(5)
    
    
End Sub

    
Public Sub ImlInfLoad(sPrmValue As String, tPrmImlData As ImlInfRec)

    On Error GoTo ImlInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmImlData.ImlOcmNum = vVal(i)
    i = i + 1
    tPrmImlData.ImlAdpDte = vVal(i)
    i = i + 1
    tPrmImlData.ImlPatTyp = vVal(i)
    i = i + 1
    tPrmImlData.ImlExpDte = vVal(i)
    i = i + 1
    tPrmImlData.ImlBrfCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlBrfQty = vVal(i)
    i = i + 1
    tPrmImlData.ImlBrfCal = vVal(i)
    i = i + 1
    tPrmImlData.ImlBrfcc = vVal(i)
    i = i + 1
    tPrmImlData.ImlLnhCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlLnhQty = vVal(i)
    i = i + 1
    tPrmImlData.ImlLnhCal = vVal(i)
    i = i + 1
    tPrmImlData.ImlLnhcc = vVal(i)
    i = i + 1
    tPrmImlData.ImlDnrCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlDnrQty = vVal(i)
    i = i + 1
    tPrmImlData.ImlDnrCal = vVal(i)
    i = i + 1
    tPrmImlData.ImlDnrcc = vVal(i)
    i = i + 1
    tPrmImlData.ImlSplCmt = vVal(i)
    i = i + 1
    tPrmImlData.ImlWhyCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlUidCod = vVal(i)
    i = i + 1
    tPrmImlData.ImlEntDtm = vVal(i)
    i = i + 1
    tPrmImlData.ImlEtcCmt = vVal(i)
    
    Exit Sub

ImlInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ImlInfStore(sPrmKey As String, sPrmValue As String, tPrmImlData As ImlInfRec)

    
    sPrmKey = tPrmImlData.ImlOcmNum & Chr(5)
    sPrmKey = sPrmKey & tPrmImlData.ImlAdpDte & Chr(5)
    sPrmKey = sPrmKey & tPrmImlData.ImlPatTyp & Chr(5)
    
    sPrmValue = tPrmImlData.ImlExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlBrfCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlBrfQty & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlBrfCal & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlBrfcc & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlLnhCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlLnhQty & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlLnhCal & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlLnhcc & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlDnrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlDnrQty & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlDnrCal & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlDnrcc & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlWhyCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlEntDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmImlData.ImlEtcCmt & Chr(5)
    
End Sub

    
Public Sub InoInfLoad(sPrmValue As String, tPrmData As InoInfRec)

    On Error GoTo InoInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.InoOcmNum = vVal(i)
    i = i + 1
    tPrmData.InoSrtDte = vVal(i)
    i = i + 1
    tPrmData.InoEndDte = vVal(i)
                                                                
    Exit Sub

InoInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub InoInfStore(sPrmKey As String, sPrmValue As String, tPrmData As InoInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.InoOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.InoSrtDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.InoEndDte & Chr(5)
    
End Sub

    
Public Sub IrcInfLoad(sPrmValue As String, tPrmData As IrcInfRec)

    On Error GoTo IrcInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.IrcOcmNum = vVal(i)
    i = i + 1
    tPrmData.IrcOcmSeq = vVal(i)
    i = i + 1
    tPrmData.IrcDupSeq = vVal(i)
    i = i + 1
    tPrmData.IrcRcpNum = vVal(i)
    i = i + 1
    tPrmData.IrcChtNum = vVal(i)
    i = i + 1
    tPrmData.IrcIrpNum = vVal(i)
    i = i + 1
    tPrmData.IrcRcpTyp = vVal(i)
    i = i + 1
    tPrmData.IrcDepCod = vVal(i)
    i = i + 1
    tPrmData.IrcRetYon = vVal(i)
    i = i + 1
    tPrmData.IrcRcpAmt = vVal(i)
    i = i + 1
    tPrmData.IrcRcpDtm = vVal(i)
    i = i + 1
    tPrmData.IrcRetAmt = vVal(i)
    i = i + 1
    tPrmData.IrcRetDtm = vVal(i)
    i = i + 1
    tPrmData.IrcUidCod = vVal(i)
    i = i + 1
    tPrmData.IrcRetUid = vVal(i)
    i = i + 1
    tPrmData.IrcRelCod = vVal(i)
    i = i + 1
    tPrmData.IrcManNam = vVal(i)
    i = i + 1
    tPrmData.IrcPreCas = vVal(i)
    i = i + 1
    tPrmData.IrcInsCod = vVal(i)
    i = i + 1
    tPrmData.IrcDtlNum = vVal(i)
    i = i + 1
    tPrmData.IrcCarFut = vVal(i)
    i = i + 1
    tPrmData.IrcCarRet = vVal(i)
    
    Exit Sub

IrcInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IrcInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IrcInfRec)

    Dim i As Integer
    
    sPrmKey = Format(CDouble(tPrmData.IrcOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IrcOcmSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IrcDupSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IrcRcpNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(tPrmData.IrcChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IrcIrpNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcRcpTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcRetYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcRcpAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcRcpDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcRetAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcRetDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcRetUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcRelCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcManNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcPreCas & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcInsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcDtlNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcCarFut & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrcCarRet & Chr(5)
    
End Sub

    
Public Sub IrpInfLoad(sPrmValue As String, tPrmData As IrpInfRec)

    On Error GoTo IrpInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.IrpOcmNum = vVal(i)
    i = i + 1
    tPrmData.IrpOcmSeq = vVal(i)
    i = i + 1
    tPrmData.IrpDupSeq = vVal(i)
    i = i + 1
    tPrmData.IrpDepCod = vVal(i)
    i = i + 1
    tPrmData.IrpChtNum = vVal(i)
    i = i + 1
    tPrmData.IrpDtrCod = vVal(i)
    i = i + 1
    tPrmData.IrpInsCod = vVal(i)
    i = i + 1
    tPrmData.IrpInsSeq = vVal(i)
    i = i + 1
    tPrmData.IrpTotAmt = vVal(i)
    i = i + 1
    tPrmData.IrpCorAmt = vVal(i)
    i = i + 1
    tPrmData.IrpNonAmt = vVal(i)
    i = i + 1
    tPrmData.IrpOwnAmt = vVal(i)
    i = i + 1
    tPrmData.IrpTotOwn = vVal(i)
    i = i + 1
    tPrmData.IrpInsAmt = vVal(i)
    i = i + 1
    tPrmData.IrpSpcAmt = vVal(i)
    i = i + 1
    tPrmData.IrpAskAmt = vVal(i)
    i = i + 1
    tPrmData.IrpDisAmt = vVal(i)
    i = i + 1
    tPrmData.IrpFutAmt = vVal(i)
    i = i + 1
    tPrmData.IrpOldAmt = vVal(i)
    i = i + 1
    tPrmData.IrpNewAmt = vVal(i)
    i = i + 1
    tPrmData.IrpGrnAmt = vVal(i)
    i = i + 1
    tPrmData.IrpRcpNum = vVal(i)
    i = i + 1
    tPrmData.IrpOldNum = vVal(i)
    i = i + 1
    tPrmData.IrpCalDte = vVal(i)
    i = i + 1
    tPrmData.IrpUpdDtm = vVal(i)
    i = i + 1
    tPrmData.IrpUidCod = vVal(i)
    i = i + 1
    tPrmData.IrpOvrAmt = vVal(i)
    i = i + 1
    tPrmData.IrpRcpFlg = vVal(i)
    i = i + 1
    tPrmData.IrpDimAmt = vVal(i)
    i = i + 1
    tPrmData.IrpNonIns = vVal(i)
    i = i + 1
    tPrmData.IrpCarFut = vVal(i)
    i = i + 1
    tPrmData.IrpFodAmt = vVal(i)
    
    '20040102..HTS..add
    i = i + 1
'--------------------------------------> 추가
    tPrmData.IrpNinAmt = vVal(i)
'--------------------------------------> 추가
    Exit Sub

IrpInfLoad_ErrorTraping:
    Resume Next

End Sub
Public Sub IrpInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IrpInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.IrpOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IrpOcmSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.IrpDupSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.IrpDepCod & Chr(5)
    
    sPrmValue = Format(tPrmData.IrpChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IrpInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpTotAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpCorAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpNonAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpOwnAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpTotOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpSpcAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpAskAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpDisAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpFutAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpOldAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpNewAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpGrnAmt & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IrpRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.IrpOldNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpCalDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpOvrAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpRcpFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpDimAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpNonIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpCarFut & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IrpFodAmt & Chr(5)
'--------------------------------------> 추가
    sPrmValue = sPrmValue & tPrmData.IrpNinAmt & Chr(5) '20040102..HTS..add
'--------------------------------------> 추가
    
End Sub

    
Public Sub IspInfLoad(sPrmValue As String, tPrmData As IspInfRec)

    On Error GoTo IspInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    With tPrmData
    
    i = i + 1
    .IspOcmNum = vVal(i)
    i = i + 1
    .IspOdrNum = vVal(i)
    i = i + 1
    .IspOdrSeq = vVal(i)
    i = i + 1
    .IspOdrCod = vVal(i)
    i = i + 1
    .IspOdrTyp = vVal(i)
    i = i + 1
    .IspSlpDep = vVal(i)
    i = i + 1
    .IspSlpCod = vVal(i)
    i = i + 1
    .IspItmCod = vVal(i)
    i = i + 1
    .IspFeeCod = vVal(i)
    i = i + 1
    .IspOdrPrc = vVal(i)
    i = i + 1
    .IspOdrSib = vVal(i)
    i = i + 1
    .IspOdrQty = vVal(i)
    i = i + 1
    .IspOdrDay = vVal(i)
    i = i + 1
    .IspOdrTms = vVal(i)
    i = i + 1
    .IspInsYon = vVal(i)
    i = i + 1
    .IspInsCod = vVal(i)
    i = i + 1
    .IspInsSeq = vVal(i)
    i = i + 1
    .IspDgsEtc = vVal(i)
    i = i + 1
    .IspOdrDnh = vVal(i)
    i = i + 1
    .IspOprDtm = vVal(i)
    i = i + 1
    .IspDepCod = vVal(i)
    i = i + 1
    .IspOdrDtm = vVal(i)
    i = i + 1
    .IspOdrStt = vVal(i)
    i = i + 1
    .IspStkStt = vVal(i)
    i = i + 1
    .IspEmgYon = vVal(i)
    i = i + 1
    .IspSpcYon = vVal(i)
    i = i + 1
    .IspCmpSym = vVal(i)
    i = i + 1
    .IspUsgCod = vVal(i)
    i = i + 1
    .IspMthCod = vVal(i)
    i = i + 1
    .IspSpmCod = vVal(i)
    i = i + 1
    .IspIncCod = vVal(i)
    i = i + 1
    .IspSotCod = vVal(i)
    i = i + 1
    .IspSlpAmt = vVal(i)
    i = i + 1
    .IspAddCod = vVal(i)
    i = i + 1
    .IspDupYon = vVal(i)
    i = i + 1
    .IspDscMed = vVal(i)
    i = i + 1
    .IspAftYon = vVal(i)
    i = i + 1
    .IspAddAmt = vVal(i)
    i = i + 1
    .IspPreDtm = vVal(i)
    i = i + 1
    .IspEntDtm = vVal(i)
    i = i + 1
    .IspUidCod = vVal(i)
    i = i + 1
    .IspCncDtm = vVal(i)
    i = i + 1
    .IspCncUid = vVal(i)
    i = i + 1
    .IspSplYon = vVal(i)
    i = i + 1
    .IspSplCmt = vVal(i)
    i = i + 1
    .IspChkStt = vVal(i)
    i = i + 1
    .IspDgsRol = vVal(i)
    i = i + 1
    .IspCvtYon = vVal(i)
    i = i + 1
    .IspMdcNum = vVal(i)
    i = i + 1
    .IspCanMdc = vVal(i)
    i = i + 1
    .IspStgCod = vVal(i)
    i = i + 1
    .IspBasUnt = vVal(i)
    i = i + 1
    .IspMntUsg = vVal(i)
    i = i + 1
    .IspXryPtb = vVal(i)
    i = i + 1
    .IspPreStt = vVal(i)
    i = i + 1
    .IspConYon = vVal(i)
    i = i + 1
    .IspUidPrt = vVal(i)
    i = i + 1
    .IspCncPrt = vVal(i)
    i = i + 1
    .IspImgYon = vVal(i)
    i = i + 1
    .IspCanNum = vVal(i)
    i = i + 1
    .IspCanSeq = vVal(i)
    i = i + 1
    .IspAstCod = vVal(i)
    i = i + 1
    .IspMixNum = vVal(i)
    i = i + 1
    .IspDenRgn = vVal(i)        '02.03.21 sebal 코드별 치식 입력.
    i = i + 1
    .IspEodNum = vVal(i)       '62       EodInf order Number
    i = i + 1
    .IspEodSeq = vVal(i)       '63       EodInf Order Sequence
    i = i + 1
    .IspIctNum = vVal(i)       '64       IctInf order Number
    i = i + 1
    .IspIctSeq = vVal(i)       '65       IctInf Order Sequence

    End With

    Exit Sub

IspInfLoad_ErrorTraping:
    Resume Next

End Sub
Public Sub IctInfLoad(sPrmValue As String, tPrmData As IctInfRec)

    On Error GoTo IspInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    With tPrmData
    
    i = i + 1
    .IspOcmNum = vVal(i)
    i = i + 1
    .IspOdrNum = vVal(i)
    i = i + 1
    .IspOdrSeq = vVal(i)
    i = i + 1
    .IspOdrCod = vVal(i)
    i = i + 1
    .IspOdrTyp = vVal(i)
    i = i + 1
    .IspSlpDep = vVal(i)
    i = i + 1
    .IspSlpCod = vVal(i)
    i = i + 1
    .IspItmCod = vVal(i)
    i = i + 1
    .IspFeeCod = vVal(i)
    i = i + 1
    .IspOdrPrc = vVal(i)
    i = i + 1
    .IspOdrSib = vVal(i)
    i = i + 1
    .IspOdrQty = vVal(i)
    i = i + 1
    .IspOdrDay = vVal(i)
    i = i + 1
    .IspOdrTms = vVal(i)
    i = i + 1
    .IspInsYon = vVal(i)
    i = i + 1
    .IspInsCod = vVal(i)
    i = i + 1
    .IspInsSeq = vVal(i)
    i = i + 1
    .IspDgsEtc = vVal(i)
    i = i + 1
    .IspOdrDnh = vVal(i)
    i = i + 1
    .IspOprDtm = vVal(i)
    i = i + 1
    .IspDepCod = vVal(i)
    i = i + 1
    .IspOdrDtm = vVal(i)
    i = i + 1
    .IspOdrStt = vVal(i)
    i = i + 1
    .IspStkStt = vVal(i)
    i = i + 1
    .IspEmgYon = vVal(i)
    i = i + 1
    .IspSpcYon = vVal(i)
    i = i + 1
    .IspCmpSym = vVal(i)
    i = i + 1
    .IspUsgCod = vVal(i)
    i = i + 1
    .IspMthCod = vVal(i)
    i = i + 1
    .IspSpmCod = vVal(i)
    i = i + 1
    .IspIncCod = vVal(i)
    i = i + 1
    .IspSotCod = vVal(i)
    i = i + 1
    .IspSlpAmt = vVal(i)
    i = i + 1
    .IspAddCod = vVal(i)
    i = i + 1
    .IspDupYon = vVal(i)
    i = i + 1
    .IspDscMed = vVal(i)
    i = i + 1
    .IspAftYon = vVal(i)
    i = i + 1
    .IspAddAmt = vVal(i)
    i = i + 1
    .IspPreDtm = vVal(i)
    i = i + 1
    .IspEntDtm = vVal(i)
    i = i + 1
    .IspUidCod = vVal(i)
    i = i + 1
    .IspCncDtm = vVal(i)
    i = i + 1
    .IspCncUid = vVal(i)
    i = i + 1
    .IspSplYon = vVal(i)
    i = i + 1
    .IspSplCmt = vVal(i)
    i = i + 1
    .IspChkStt = vVal(i)
    i = i + 1
    .IspDgsRol = vVal(i)
    i = i + 1
    .IspCvtYon = vVal(i)
    i = i + 1
    .IspMdcNum = vVal(i)
    i = i + 1
    .IspCanMdc = vVal(i)
    i = i + 1
    .IspStgCod = vVal(i)
    i = i + 1
    .IspBasUnt = vVal(i)
    i = i + 1
    .IspMntUsg = vVal(i)
    i = i + 1
    .IspXryPtb = vVal(i)
    i = i + 1
    .IspPreStt = vVal(i)
    i = i + 1
    .IspConYon = vVal(i)
    i = i + 1
    .IspUidPrt = vVal(i)
    i = i + 1
    .IspCncPrt = vVal(i)
    i = i + 1
    .IspImgYon = vVal(i)
    i = i + 1
    .IspCanNum = vVal(i)
    i = i + 1
    .IspCanSeq = vVal(i)
    i = i + 1
    .IspAstCod = vVal(i)
    i = i + 1
    .IspMixNum = vVal(i)
    i = i + 1
    .IspDenRgn = vVal(i)        '02.03.21 sebal 코드별 치식 입력.
    i = i + 1
    .IspEodNum = vVal(i)       '62       EdoInf order Number
    i = i + 1
    .IspEodSeq = vVal(i)       '63       EodInf Order Sequence

    End With

    Exit Sub

IspInfLoad_ErrorTraping:
    Resume Next

End Sub

    
Public Sub IspInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IspInfRec)

    With tPrmData
    
    sPrmKey = Format(CDouble(.IspOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(.IspOdrNum), "@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(.IspOdrSeq), "@@@@@") & Chr(5)
    
    sPrmValue = .IspOdrCod & Chr(5)
    sPrmValue = sPrmValue & .IspOdrTyp & Chr(5)
    sPrmValue = sPrmValue & .IspSlpDep & Chr(5)
    sPrmValue = sPrmValue & .IspSlpCod & Chr(5)
    sPrmValue = sPrmValue & .IspItmCod & Chr(5)
    sPrmValue = sPrmValue & .IspFeeCod & Chr(5)
    sPrmValue = sPrmValue & .IspOdrPrc & Chr(5)
    sPrmValue = sPrmValue & .IspOdrSib & Chr(5)
    sPrmValue = sPrmValue & .IspOdrQty & Chr(5)
    sPrmValue = sPrmValue & .IspOdrDay & Chr(5)
    sPrmValue = sPrmValue & .IspOdrTms & Chr(5)
    sPrmValue = sPrmValue & .IspInsYon & Chr(5)
    sPrmValue = sPrmValue & .IspInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(.IspInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & .IspDgsEtc & Chr(5)
    sPrmValue = sPrmValue & .IspOdrDnh & Chr(5)
    sPrmValue = sPrmValue & .IspOprDtm & Chr(5)
    sPrmValue = sPrmValue & .IspDepCod & Chr(5)
    sPrmValue = sPrmValue & .IspOdrDtm & Chr(5)
    sPrmValue = sPrmValue & .IspOdrStt & Chr(5)
    sPrmValue = sPrmValue & .IspStkStt & Chr(5)
    sPrmValue = sPrmValue & .IspEmgYon & Chr(5)
    sPrmValue = sPrmValue & .IspSpcYon & Chr(5)
    sPrmValue = sPrmValue & .IspCmpSym & Chr(5)
    sPrmValue = sPrmValue & .IspUsgCod & Chr(5)
    sPrmValue = sPrmValue & .IspMthCod & Chr(5)
    sPrmValue = sPrmValue & .IspSpmCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(.IspIncCod), "@@") & Chr(5)
    sPrmValue = sPrmValue & .IspSotCod & Chr(5)
    sPrmValue = sPrmValue & .IspSlpAmt & Chr(5)
    sPrmValue = sPrmValue & .IspAddCod & Chr(5)
    sPrmValue = sPrmValue & .IspDupYon & Chr(5)
    sPrmValue = sPrmValue & .IspDscMed & Chr(5)
    sPrmValue = sPrmValue & .IspAftYon & Chr(5)
    sPrmValue = sPrmValue & .IspAddAmt & Chr(5)
    sPrmValue = sPrmValue & .IspPreDtm & Chr(5)
    sPrmValue = sPrmValue & .IspEntDtm & Chr(5)
    sPrmValue = sPrmValue & .IspUidCod & Chr(5)
    sPrmValue = sPrmValue & .IspCncDtm & Chr(5)
    sPrmValue = sPrmValue & .IspCncUid & Chr(5)
    sPrmValue = sPrmValue & .IspSplYon & Chr(5)
    sPrmValue = sPrmValue & .IspSplCmt & Chr(5)
    sPrmValue = sPrmValue & .IspChkStt & Chr(5)
    sPrmValue = sPrmValue & .IspDgsRol & Chr(5)
    sPrmValue = sPrmValue & .IspCvtYon & Chr(5)
    sPrmValue = sPrmValue & .IspMdcNum & Chr(5)
    sPrmValue = sPrmValue & .IspCanMdc & Chr(5)
    sPrmValue = sPrmValue & .IspStgCod & Chr(5)
    sPrmValue = sPrmValue & .IspBasUnt & Chr(5)
    sPrmValue = sPrmValue & .IspMntUsg & Chr(5)
    sPrmValue = sPrmValue & .IspXryPtb & Chr(5)
    sPrmValue = sPrmValue & .IspPreStt & Chr(5)
    sPrmValue = sPrmValue & .IspConYon & Chr(5)
    sPrmValue = sPrmValue & .IspUidPrt & Chr(5)
    sPrmValue = sPrmValue & .IspCncPrt & Chr(5)
    sPrmValue = sPrmValue & .IspImgYon & Chr(5)
    sPrmValue = sPrmValue & .IspCanNum & Chr(5)
    sPrmValue = sPrmValue & .IspCanSeq & Chr(5)
    sPrmValue = sPrmValue & .IspAstCod & Chr(5)
    sPrmValue = sPrmValue & .IspMixNum & Chr(5)
    sPrmValue = sPrmValue & .IspDenRgn & Chr(5)         '02.03.21 sebal 코드별 치식 입력.
    sPrmValue = sPrmValue & .IspEodNum & Chr(5)
    sPrmValue = sPrmValue & .IspEodSeq & Chr(5)
    sPrmValue = sPrmValue & .IspIctNum & Chr(5)
    sPrmValue = sPrmValue & .IspIctSeq & Chr(5)
    
    End With

End Sub

    Public Sub IctInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IctInfRec)

    With tPrmData
    
    sPrmKey = Format(CDouble(.IspOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(.IspOdrNum), "@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(.IspOdrSeq), "@@@@@") & Chr(5)
    
    sPrmValue = .IspOdrCod & Chr(5)
    sPrmValue = sPrmValue & .IspOdrTyp & Chr(5)
    sPrmValue = sPrmValue & .IspSlpDep & Chr(5)
    sPrmValue = sPrmValue & .IspSlpCod & Chr(5)
    sPrmValue = sPrmValue & .IspItmCod & Chr(5)
    sPrmValue = sPrmValue & .IspFeeCod & Chr(5)
    sPrmValue = sPrmValue & .IspOdrPrc & Chr(5)
    sPrmValue = sPrmValue & .IspOdrSib & Chr(5)
    sPrmValue = sPrmValue & .IspOdrQty & Chr(5)
    sPrmValue = sPrmValue & .IspOdrDay & Chr(5)
    sPrmValue = sPrmValue & .IspOdrTms & Chr(5)
    sPrmValue = sPrmValue & .IspInsYon & Chr(5)
    sPrmValue = sPrmValue & .IspInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(.IspInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & .IspDgsEtc & Chr(5)
    sPrmValue = sPrmValue & .IspOdrDnh & Chr(5)
    sPrmValue = sPrmValue & .IspOprDtm & Chr(5)
    sPrmValue = sPrmValue & .IspDepCod & Chr(5)
    sPrmValue = sPrmValue & .IspOdrDtm & Chr(5)
    sPrmValue = sPrmValue & .IspOdrStt & Chr(5)
    sPrmValue = sPrmValue & .IspStkStt & Chr(5)
    sPrmValue = sPrmValue & .IspEmgYon & Chr(5)
    sPrmValue = sPrmValue & .IspSpcYon & Chr(5)
    sPrmValue = sPrmValue & .IspCmpSym & Chr(5)
    sPrmValue = sPrmValue & .IspUsgCod & Chr(5)
    sPrmValue = sPrmValue & .IspMthCod & Chr(5)
    sPrmValue = sPrmValue & .IspSpmCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(.IspIncCod), "@@") & Chr(5)
    sPrmValue = sPrmValue & .IspSotCod & Chr(5)
    sPrmValue = sPrmValue & .IspSlpAmt & Chr(5)
    sPrmValue = sPrmValue & .IspAddCod & Chr(5)
    sPrmValue = sPrmValue & .IspDupYon & Chr(5)
    sPrmValue = sPrmValue & .IspDscMed & Chr(5)
    sPrmValue = sPrmValue & .IspAftYon & Chr(5)
    sPrmValue = sPrmValue & .IspAddAmt & Chr(5)
    sPrmValue = sPrmValue & .IspPreDtm & Chr(5)
    sPrmValue = sPrmValue & .IspEntDtm & Chr(5)
    sPrmValue = sPrmValue & .IspUidCod & Chr(5)
    sPrmValue = sPrmValue & .IspCncDtm & Chr(5)
    sPrmValue = sPrmValue & .IspCncUid & Chr(5)
    sPrmValue = sPrmValue & .IspSplYon & Chr(5)
    sPrmValue = sPrmValue & .IspSplCmt & Chr(5)
    sPrmValue = sPrmValue & .IspChkStt & Chr(5)
    sPrmValue = sPrmValue & .IspDgsRol & Chr(5)
    sPrmValue = sPrmValue & .IspCvtYon & Chr(5)
    sPrmValue = sPrmValue & .IspMdcNum & Chr(5)
    sPrmValue = sPrmValue & .IspCanMdc & Chr(5)
    sPrmValue = sPrmValue & .IspStgCod & Chr(5)
    sPrmValue = sPrmValue & .IspBasUnt & Chr(5)
    sPrmValue = sPrmValue & .IspMntUsg & Chr(5)
    sPrmValue = sPrmValue & .IspXryPtb & Chr(5)
    sPrmValue = sPrmValue & .IspPreStt & Chr(5)
    sPrmValue = sPrmValue & .IspConYon & Chr(5)
    sPrmValue = sPrmValue & .IspUidPrt & Chr(5)
    sPrmValue = sPrmValue & .IspCncPrt & Chr(5)
    sPrmValue = sPrmValue & .IspImgYon & Chr(5)
    sPrmValue = sPrmValue & .IspCanNum & Chr(5)
    sPrmValue = sPrmValue & .IspCanSeq & Chr(5)
    sPrmValue = sPrmValue & .IspAstCod & Chr(5)
    sPrmValue = sPrmValue & .IspMixNum & Chr(5)
    sPrmValue = sPrmValue & .IspDenRgn & Chr(5)         '02.03.21 sebal 코드별 치식 입력.
    sPrmValue = sPrmValue & .IspEodNum & Chr(5)
    sPrmValue = sPrmValue & .IspEodSeq & Chr(5)
    
    End With

End Sub
Public Sub ItrHstLoad(sPrmValue As String, tPrmData As ItrHstRec)

    On Error GoTo ItrHstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.ItrOcmNum = vVal(i)
    i = i + 1
    tPrmData.ItrStrDtm = vVal(i)
    i = i + 1
    tPrmData.ItrDelDtm = vVal(i)
    i = i + 1
    tPrmData.ItrSttFlg = vVal(i)
    i = i + 1
    tPrmData.ItrEndDtm = vVal(i)
    i = i + 1
    tPrmData.ItrChtNum = vVal(i)
    i = i + 1
    tPrmData.ItrDepCod = vVal(i)
    i = i + 1
    tPrmData.ItrDtrCod = vVal(i)
    i = i + 1
    tPrmData.ItrNssCod = vVal(i)
    i = i + 1
    tPrmData.ItrRomCod = vVal(i)
    i = i + 1
    tPrmData.ItrBedCod = vVal(i)
    i = i + 1
    tPrmData.ItrBedGrd = vVal(i)
    i = i + 1
    tPrmData.ItrSpcYon = vVal(i)
    i = i + 1
    tPrmData.ItrWhyCod = vVal(i)
    i = i + 1
    tPrmData.ItrUpdDtm = vVal(i)
    i = i + 1
    tPrmData.ItrUidCod = vVal(i)
    
    Exit Sub

ItrHstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ItrHstStore(sPrmKey As String, sPrmValue As String, tPrmData As ItrHstRec)

    
    sPrmKey = Format(CDouble(tPrmData.ItrOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ItrStrDtm & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ItrDelDtm & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ItrSttFlg & Chr(5)
    
    sPrmValue = tPrmData.ItrEndDtm & Chr(5)
    sPrmValue = sPrmValue & Format(tPrmData.ItrChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrNssCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrRomCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrBedCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrBedGrd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrWhyCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrUidCod & Chr(5)
    
End Sub

    
Public Sub ItrInfLoad(sPrmValue As String, tPrmData As ItrInfRec)

    On Error GoTo ItrInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.ItrOcmNum = vVal(i)
    i = i + 1
    tPrmData.ItrStrDtm = vVal(i)
    i = i + 1
    tPrmData.ItrEndDtm = vVal(i)
    i = i + 1
    tPrmData.ItrChtNum = vVal(i)
    i = i + 1
    tPrmData.ItrDepCod = vVal(i)
    i = i + 1
    tPrmData.ItrDtrCod = vVal(i)
    i = i + 1
    tPrmData.ItrNssCod = vVal(i)
    i = i + 1
    tPrmData.ItrRomCod = vVal(i)
    i = i + 1
    tPrmData.ItrBedCod = vVal(i)
    i = i + 1
    tPrmData.ItrBedGrd = vVal(i)
    i = i + 1
    tPrmData.ItrSpcYon = vVal(i)
    i = i + 1
    tPrmData.ItrWhyCod = vVal(i)
    i = i + 1
    tPrmData.ItrUpdDtm = vVal(i)
    i = i + 1
    tPrmData.ItrUidCod = vVal(i)
    
    Exit Sub

ItrInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ItrInfStore(sPrmKey As String, sPrmValue As String, tPrmData As ItrInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.ItrOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ItrStrDtm & Chr(5)
    
    sPrmValue = tPrmData.ItrEndDtm & Chr(5)
    sPrmValue = sPrmValue & Format(tPrmData.ItrChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrNssCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrRomCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrBedCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrBedGrd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrWhyCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ItrUidCod & Chr(5)
    
End Sub

    
Public Sub IwlInfLoad(sPrmValue As String, tPrmData As IwlInfRec)

    On Error GoTo IwlInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.IwlTypCod = vVal(i)
    i = i + 1
    tPrmData.IwlAcpNum = vVal(i)
    i = i + 1
    tPrmData.IwlChtNum = vVal(i)
    i = i + 1
    tPrmData.IwlAcpDte = vVal(i)
    i = i + 1
    tPrmData.IwlReqNam = vVal(i)
    i = i + 1
    tPrmData.IwlActDte = vVal(i)
    i = i + 1
    tPrmData.IwlEntDte = vVal(i)
    i = i + 1
    tPrmData.IwlNssCod = vVal(i)
    i = i + 1
    tPrmData.IwlFstPhn = vVal(i)
    i = i + 1
    tPrmData.IwlSndPhn = vVal(i)
    i = i + 1
    tPrmData.IwlSplCmt = vVal(i)
    i = i + 1
    tPrmData.IwlStrUid = vVal(i)
    i = i + 1
    tPrmData.IwlStrDtm = vVal(i)
    i = i + 1
    tPrmData.IwlUpdUid = vVal(i)
    i = i + 1
    tPrmData.IwlUpdDtm = vVal(i)
    
    Exit Sub

IwlInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IwlInfStore(sPrmKey As String, sPrmValue As String, tPrmData As IwlInfRec)

    
    sPrmKey = tPrmData.IwlTypCod & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmData.IwlAcpNum, "@@@@") & Chr(5)
    
    sPrmValue = Format(tPrmData.IwlChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlAcpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlReqNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlActDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlEntDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlNssCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlFstPhn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlSndPhn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlStrUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlStrDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlUpdUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.IwlUpdDtm & Chr(5)
    
End Sub

    
Public Sub LbaInfLoad(sPrmValue As String, tPrmData As LbaInfRec)

    On Error GoTo LbaInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.LbaSotCod = vVal(i)
    i = i + 1
    tPrmData.LbaOcmNum = vVal(i)
    i = i + 1
    tPrmData.LbaSeq = vVal(i)
    i = i + 1
    tPrmData.LbaSlpCod = vVal(i)
    i = i + 1
    tPrmData.LbaChtNum = vVal(i)
    i = i + 1
    tPrmData.LbaOdrDte = vVal(i)
    i = i + 1
    tPrmData.LbaComStt = vVal(i)
    i = i + 1
    tPrmData.LbaDepCod = vVal(i)
    i = i + 1
    tPrmData.LbaDtrCod = vVal(i)
    i = i + 1
    tPrmData.LbaRomCod = vVal(i)
    i = i + 1
    tPrmData.LbaEmgYon = vVal(i)
    i = i + 1
    tPrmData.LbaOdrStt = vVal(i)
    i = i + 1
    tPrmData.LbaSpmNum = vVal(i)
    i = i + 1
    tPrmData.LbaAcpDtm = vVal(i)
    i = i + 1
    tPrmData.LbaAcpUid = vVal(i)
    i = i + 1
    tPrmData.LbaRstDte = vVal(i)
    i = i + 1
    tPrmData.LbaRptDte = vVal(i)
    i = i + 1
    tPrmData.LbaRptUid = vVal(i)
    i = i + 1
    tPrmData.LbaSlpDep = vVal(i)
    i = i + 1
    tPrmData.LbaSplCmt = vVal(i)
    i = i + 1
    tPrmData.LbaUidCod = vVal(i)
    i = i + 1
    tPrmData.LbaFstRed = vVal(i)
    i = i + 1
    tPrmData.LbaSndRed = vVal(i)
    i = i + 1
    tPrmData.LbaTrdRed = vVal(i)
    i = i + 1
    tPrmData.LbaForRed = vVal(i)
    i = i + 1
    tPrmData.LbaFifRed = vVal(i)
    
    Exit Sub

LbaInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LbaInfStore(sPrmKey As String, sPrmValue As String, tPrmData As LbaInfRec)

    
    sPrmKey = tPrmData.LbaSotCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.LbaOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.LbaSeq & Chr(5)
    sPrmKey = sPrmKey & tPrmData.LbaSlpCod & Chr(5)
    
    sPrmValue = Format((tPrmData.LbaChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaOdrDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaComStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaRomCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaEmgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaOdrStt & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.LbaSpmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaAcpDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaAcpUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaRstDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaRptDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaRptUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaSlpDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaFstRed & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaSndRed & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaTrdRed & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaForRed & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbaFifRed & Chr(5)
    
End Sub

    
Public Sub LbbInfLoad(sPrmValue As String, tPrmData As LbbInfRec)

    On Error GoTo LbbInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.LbbSotCod = vVal(i)
    i = i + 1
    tPrmData.LbbOcmNum = vVal(i)
    i = i + 1
    tPrmData.LbbOdrDte = vVal(i)
    i = i + 1
    tPrmData.LbbChtNum = vVal(i)
    i = i + 1
    tPrmData.LbbLabDte = vVal(i)
    i = i + 1
    tPrmData.LbbLabTim = vVal(i)
    i = i + 1
    tPrmData.LbbCanFlg = vVal(i)
    i = i + 1
    tPrmData.LbbOcmStt = vVal(i)
    i = i + 1
    tPrmData.LbbPreYon = vVal(i)
    
    Exit Sub

LbbInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LbbInfStore(sPrmKey As String, sPrmValue As String, tPrmData As LbbInfRec)

    
    sPrmKey = tPrmData.LbbSotCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.LbbOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.LbbOdrDte & Chr(5)
    
    sPrmValue = Format((tPrmData.LbbChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbbLabDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbbLabTim & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbbCanFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbbOcmStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbbPreYon & Chr(5)
    
End Sub

    
Public Sub LbqInfLoad(sPrmValue As String, tPrmData As LbqInfRec)

    On Error GoTo LbqInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.LbqSotCod = vVal(i)
    i = i + 1
    tPrmData.LbqAcpDte = vVal(i)
    i = i + 1
    tPrmData.LbqOcmNum = vVal(i)
    i = i + 1
    tPrmData.LbqChtNum = vVal(i)
    i = i + 1
    tPrmData.LbqPatNam = vVal(i)
    i = i + 1
    tPrmData.LbqDepCod = vVal(i)
    i = i + 1
    tPrmData.LbqWrdCod = vVal(i)
    i = i + 1
    tPrmData.LbqRomCod = vVal(i)
    i = i + 1
    tPrmData.LbqAcpStt = vVal(i)
    i = i + 1
    tPrmData.LbqCodCnt = vVal(i)
    i = i + 1
    tPrmData.LbqEmgCnt = vVal(i)
    i = i + 1
    tPrmData.LbqCasCnt = vVal(i)
    i = i + 1
    tPrmData.LbqCanCnt = vVal(i)
    i = i + 1
    tPrmData.LbqPreDay = vVal(i)
    i = i + 1
    tPrmData.LbqRsvYon = vVal(i)
    i = i + 1
    tPrmData.LbqAcpTms = vVal(i)
    i = i + 1
    tPrmData.LbqCfmDte = vVal(i)
    i = i + 1
    tPrmData.LbqOdrDte = vVal(i)
    i = i + 1
    tPrmData.LbqCasDtm = vVal(i)
    i = i + 1
    tPrmData.LbqRstMak = vVal(i)
    i = i + 1
    tPrmData.LbqStdPat = vVal(i)
                
    Exit Sub

LbqInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LbqInfStore(sPrmKey As String, sPrmValue As String, tPrmData As LbqInfRec)

    
    sPrmKey = tPrmData.LbqSotCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.LbqAcpDte & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.LbqOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format((tPrmData.LbqChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqWrdCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqRomCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqAcpStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqCodCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqEmgCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqCasCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqCanCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqPreDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqRsvYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqAcpTms & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqCfmDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqOdrDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqCasDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqRstMak & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LbqStdPat & Chr(5)

End Sub

    
Public Sub LhrInfLoad(sPrmValue As String, tPrmData As LhrInfRec)

    On Error GoTo LhrInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    '
    i = i + 1
    tPrmData.LhrOcmNum = vVal(i)
    i = i + 1
    tPrmData.LhrLevDte = vVal(i)
    i = i + 1
    tPrmData.LhrEduYer = vVal(i)
    i = i + 1
    tPrmData.LhrJobCod = vVal(i)
    i = i + 1
    tPrmData.LhrJobCmt = vVal(i)
    i = i + 1
    tPrmData.LhrEcnStt = vVal(i)
    i = i + 1
    tPrmData.LhrMrgStt = vVal(i)
    i = i + 1
    tPrmData.LhrRlgCod = vVal(i)
    i = i + 1
    tPrmData.LhrFstAge = vVal(i)
    i = i + 1
    tPrmData.LhrWrsDte = vVal(i)
    i = i + 1
    tPrmData.LhrInhTyp = vVal(i)
    i = i + 1
    tPrmData.LhrOthCnt = vVal(i)
    i = i + 1
    tPrmData.LhrManPbm = vVal(i)
    i = i + 1
    tPrmData.LhrFmlHst = vVal(i)
    i = i + 1
    tPrmData.LhrFstFml = vVal(i)
    i = i + 1
    tPrmData.LhrFstDgn = vVal(i)
    i = i + 1
    tPrmData.LhrFstCmt = vVal(i)
    i = i + 1
    tPrmData.LhrSndFml = vVal(i)
    i = i + 1
    tPrmData.LhrSndDgn = vVal(i)
    i = i + 1
    tPrmData.LhrSndCmt = vVal(i)
    i = i + 1
    tPrmData.LhrTrdFml = vVal(i)
    i = i + 1
    tPrmData.LhrTrdDgn = vVal(i)
    i = i + 1
    tPrmData.LhrTrdCmt = vVal(i)
    i = i + 1
    tPrmData.LhrFthFml = vVal(i)
    i = i + 1
    tPrmData.LhrFthDgn = vVal(i)
    i = i + 1
    tPrmData.LhrFthCmt = vVal(i)
    i = i + 1
    tPrmData.LhrClnDgn = vVal(i)
    i = i + 1
    tPrmData.LhrTrbDgn = vVal(i)
    i = i + 1
    tPrmData.LhrPhyDgn = vVal(i)
    i = i + 1
    tPrmData.LhrLmtCls = vVal(i)
    i = i + 1
    tPrmData.LhrFunInh = vVal(i)
    i = i + 1
    tPrmData.LhrFunLeh = vVal(i)
    i = i + 1
    tPrmData.LhrFunBin = vVal(i)
    i = i + 1
    tPrmData.LhrEstAct = vVal(i)
    i = i + 1
    tPrmData.LhrEegLab = vVal(i)
    i = i + 1
    tPrmData.LhrPsyLab = vVal(i)
    i = i + 1
    tPrmData.LhrLevJud = vVal(i)
    i = i + 1
    tPrmData.LhrTrtRst = vVal(i)
    i = i + 1
    tPrmData.LhrLevTrt = vVal(i)
    i = i + 1
    tPrmData.LhrSpcDtr = vVal(i)
    i = i + 1
    tPrmData.LhrDtrCod = vVal(i)
    '
    Exit Sub

LhrInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LhrinfStore(sPrmKey As String, sPrmValue As String, tPrmData As LhrInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.LhrOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.LhrLevDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrEduYer & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrJobCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrJobCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrEcnStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrMrgStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrRlgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFstAge & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrWrsDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrInhTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrOthCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrManPbm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFmlHst & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFstFml & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFstDgn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFstCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrSndFml & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrSndDgn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrSndCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrTrdFml & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrTrdDgn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrTrdCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFthFml & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFthDgn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFthCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrClnDgn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrTrbDgn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrPhyDgn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrLmtCls & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFunInh & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFunLeh & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrFunBin & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrEstAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrEegLab & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrPsyLab & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrLevJud & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrTrtRst & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrLevTrt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrSpcDtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LhrDtrCod & Chr(5)
    
End Sub

    
Public Sub LocChtLoad(sPrmValue As String, tPrmData As LocChtRec)

    On Error GoTo LocChtLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.LocChtNum = vVal(i)
    i = i + 1
    tPrmData.LocLevCod = vVal(i)
    i = i + 1
    tPrmData.LocExeNam = vVal(i)
    i = i + 1
    tPrmData.LocUidCod = vVal(i)
    i = i + 1
    tPrmData.LocIpAddr = vVal(i)
    i = i + 1
    tPrmData.LocChtDtm = vVal(i)
    Exit Sub

LocChtLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LocChtStore(sPrmKey As String, sPrmValue As String, tPrmData As LocChtRec)

    sPrmKey = Format(tPrmData.LocChtNum, "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.LocLevCod & Chr(5)
    
    sPrmValue = tPrmData.LocExeNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LocUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LocIpAddr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.LocChtDtm & Chr(5)
End Sub

    
Public Sub LslInfLoad(sPrmValue As String, tPrmData As LslInfRec)

    On Error GoTo LslInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.LslOcmNum = vVal(i)
    i = i + 1
    tPrmData.LslChtNum = vVal(i)
    i = i + 1
    tPrmData.LslLevDtm = vVal(i)
    i = i + 1
    tPrmData.LslManStt = vVal(i)
    i = i + 1
    tPrmData.LslFnlIcd = vVal(i)
    i = i + 1
    tPrmData.LslIcdSum = vVal(i)
    i = i + 1
    tPrmData.LslAlgYon = vVal(i)
    i = i + 1
    tPrmData.LslAlgDtl = vVal(i)
    i = i + 1
    tPrmData.LslLevStt = vVal(i)
    i = i + 1
    tPrmData.LslDtrCod = vVal(i)
    
    Exit Sub

LslInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LslInfStore(sPrmKey As String, sPrmValue As String, tPrmData As LslInfRec)

    
    sPrmKey = tPrmData.LslOcmNum & Chr(5)
    
    sPrmValue = tPrmData.LslChtNum & Chr(5)
    sPrmValue = sPrmValue & Chr(5) & tPrmData.LslLevDtm & Chr(5)
    sPrmValue = sPrmValue & Chr(5) & tPrmData.LslManStt & Chr(5)
    sPrmValue = sPrmValue & Chr(5) & tPrmData.LslFnlIcd & Chr(5)
    sPrmValue = sPrmValue & Chr(5) & tPrmData.LslIcdSum & Chr(5)
    sPrmValue = sPrmValue & Chr(5) & tPrmData.LslAlgYon & Chr(5)
    sPrmValue = sPrmValue & Chr(5) & tPrmData.LslAlgDtl & Chr(5)
    sPrmValue = sPrmValue & Chr(5) & tPrmData.LslLevStt & Chr(5)
    sPrmValue = sPrmValue & Chr(5) & tPrmData.LslDtrCod & Chr(5)
    
End Sub

    
Public Sub MalInfLoad(sPrmValue As String, tPrmData As MalInfRec)

    On Error GoTo MalInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.MalRcvUid = vVal(i)
    i = i + 1
    tPrmData.MalSndDtm = vVal(i)
    i = i + 1
    tPrmData.MalSndUid = vVal(i)
    i = i + 1
    tPrmData.MalCfmYon = vVal(i)
    i = i + 1
    tPrmData.MalSndSts = vVal(i)
    i = i + 1
    tPrmData.MalMsgDtl = vVal(i)
    i = i + 1
    tPrmData.MalMsgSbj = vVal(i)
    i = i + 1
    tPrmData.MalApdFle = vVal(i)
    
    Exit Sub

MalInfLoad_ErrorTraping:
    Resume Next

End Sub
    
    
Public Sub MalGrpLoad(sPrmValue As String, tPrmData As MalGrpRec)

    On Error GoTo MalGrpLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.MalGrpUid = vVal(i)
    i = i + 1
    tPrmData.MalGrpCod = vVal(i)
    i = i + 1
    tPrmData.MalGrpNam = vVal(i)
    
    Exit Sub

MalGrpLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub MalDtlLoad(sPrmValue As String, tPrmData As MalDtlRec)

    On Error GoTo MalDtlLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.MalGrpUid = vVal(i)
    i = i + 1
    tPrmData.MalGrpCod = vVal(i)
    i = i + 1
    tPrmData.MalDtlUid = vVal(i)
    
    Exit Sub

MalDtlLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub MalInfStore(sPrmKey As String, sPrmValue As String, tPrmData As MalInfRec)

    
    sPrmKey = tPrmData.MalRcvUid & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MalSndDtm & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MalSndUid & Chr(5)
    
    sPrmValue = tPrmData.MalCfmYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MalSndSts & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MalMsgDtl & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MalMsgSbj & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MalApdFle & Chr(5)
    
End Sub

Public Sub MalGrpStore(sPrmKey As String, sPrmValue As String, tPrmData As MalGrpRec)

    
    sPrmKey = tPrmData.MalGrpUid & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MalGrpCod & Chr(5)
    
    sPrmValue = tPrmData.MalGrpNam & Chr(5)
    
    
End Sub
Public Sub MalDtlStore(sPrmKey As String, sPrmValue As String, tPrmData As MalDtlRec)

    
    sPrmKey = tPrmData.MalGrpUid & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MalGrpCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MalDtlUid & Chr(5)
    
    
End Sub

    
Public Sub MthInfLoad(sPrmValue As String, tPrmData As MthInfRec)

    On Error GoTo MthInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.MthChtNum = vVal(i)
    i = i + 1
    tPrmData.MthOdrDte = vVal(i)
    i = i + 1
    tPrmData.MthOcmNum = vVal(i)
    i = i + 1
    tPrmData.MthOdrCod = vVal(i)
    i = i + 1
    tPrmData.MthOdrStt = vVal(i)
    i = i + 1
    tPrmData.MthOdrQty = vVal(i)
    i = i + 1
    tPrmData.MthOdrDay = vVal(i)
    i = i + 1
    tPrmData.MthOdrTms = vVal(i)
    i = i + 1
    tPrmData.MthOdrAmt = vVal(i)
    i = i + 1
    tPrmData.MthAdpAmt = vVal(i)
    
    Exit Sub

MthInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub MthInfStore(sPrmKey As String, sPrmValue As String, tPrmData As MthInfRec)

    
    sPrmKey = Format((tPrmData.MthChtNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MthOdrDte & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MthOcmNum & Chr(5)
    sPrmKey = sPrmKey & tPrmData.MthOdrCod & Chr(5)
    
    sPrmValue = tPrmData.MthOdrStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MthOdrQty & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MthOdrDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MthOdrTms & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MthOdrAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.MthAdpAmt & Chr(5)
    
    
End Sub

    
Public Sub NblInfLoad(sPrmValue As String, tPrmData As NblInfRec)

    On Error GoTo NblInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.NblBilDte = vVal(i)
    i = i + 1
    tPrmData.NblChtNum = vVal(i)
    i = i + 1
    tPrmData.NblInsCod = vVal(i)
    i = i + 1
    tPrmData.NblSeqNum = vVal(i)
    i = i + 1
    tPrmData.NblStrDte = vVal(i)
    i = i + 1
    tPrmData.NblEndDte = vVal(i)
    i = i + 1
    tPrmData.NblDtlDte = vVal(i)
    i = i + 1
    tPrmData.NblTotAmt = vVal(i)
    i = i + 1
    tPrmData.NblAskAmt = vVal(i)
    i = i + 1
    tPrmData.NblCorAmt = vVal(i)
    i = i + 1
    tPrmData.NblFutAmt = vVal(i)
    i = i + 1
    tPrmData.NblEndFlg = vVal(i)
    i = i + 1
    tPrmData.NblCmtRef = vVal(i)
    i = i + 1
    tPrmData.NblIotFlg = vVal(i)
    i = i + 1
    tPrmData.NblAccCod = vVal(i)
    i = i + 1
    tPrmData.NblRcpNum = vVal(i)
    
    Exit Sub

NblInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub NblInfStore(sPrmKey As String, sPrmValue As String, tPrmData As NblInfRec)

    
    sPrmKey = tPrmData.NblBilDte & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmData.NblChtNum, "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.NblInsCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.NblSeqNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.NblStrDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblEndDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblDtlDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblTotAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblAskAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblCorAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblFutAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblEndFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblCmtRef & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblIotFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblAccCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NblRcpNum & Chr(5)
    
End Sub

    
Public Sub NbsInfLoad(sPrmValue As String, tPrmData As NbsInfRec)

    On Error GoTo NbsInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.NbsFrmDte = vVal(i)
    i = i + 1
    tPrmData.NbsEndDte = vVal(i)
    i = i + 1
    tPrmData.NbsPrcDte = vVal(i)
    i = i + 1
    tPrmData.NbsUidCod = vVal(i)
    i = i + 1
    tPrmData.NbsCloCnt = vVal(i)
    
    Exit Sub

NbsInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub NbsInfStore(sPrmKey As String, sPrmValue As String, tPrmData As NbsInfRec)

    
    sPrmKey = tPrmData.NbsFrmDte & Chr(5)
    
    sPrmValue = tPrmData.NbsEndDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NbsPrcDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NbsUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NbsCloCnt & Chr(5)
    
End Sub

    
Public Sub OacInfLoad(sPrmValue As String, tPrmData As OacInfRec)

    On Error GoTo OacInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.OacRcpNum = vVal(i)
    i = i + 1
    tPrmData.OacAccCod = vVal(i)
    i = i + 1
    tPrmData.OacChtNum = vVal(i)
    i = i + 1
    tPrmData.OacOcmNum = vVal(i)
    i = i + 1
    tPrmData.OacAccAmt = vVal(i)
    i = i + 1
    tPrmData.OacAccDsc = vVal(i)
    i = i + 1
    tPrmData.OacAccRat = vVal(i)
    i = i + 1
    tPrmData.OacAccDgs = vVal(i)
    i = i + 1
    tPrmData.OacFdgRat = vVal(i)
    i = i + 1
    tPrmData.OacFdgAmt = vVal(i)
    i = i + 1
    tPrmData.OacSdgRat = vVal(i)
    i = i + 1
    tPrmData.OacSdgAmt = vVal(i)
    i = i + 1
    tPrmData.OacCalRat = vVal(i)
    i = i + 1
    tPrmData.OacCalAmt = vVal(i)
    i = i + 1
    tPrmData.OacCalSeq = vVal(i)
    i = i + 1
    tPrmData.OacEmpCod = vVal(i)
    i = i + 1
    tPrmData.OacRelCod = vVal(i)
    '''yk : 계정에서 카드는 뺀다.
    'i = i + 1
    'tPrmData.OacCrdMax = vVal(i)
    'Call CrdInfLoad(sPrmValue, tPrmData.OacCrdDat(), i)
    
    Exit Sub

OacInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OacInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OacInfRec)

    
    Dim i As Integer
    
    sPrmKey = Format(CDouble(tPrmData.OacRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.OacAccCod & Chr(5)
    
    sPrmValue = Format((tPrmData.OacChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OacOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacAccAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacAccDsc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacAccRat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacAccDgs & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacFdgRat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacFdgAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacSdgRat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacSdgAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacCalRat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacCalAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacCalSeq & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacEmpCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OacRelCod & Chr(5)
    
    '''yk : 계정에 카드를 뺀다.
    'sPrmValue = sPrmValue & tPrmData.OacCrdMax & Chr(5)
    
    'For i = 1 To 10
    '    sPrmValue = sPrmValue & tPrmData.OacCrdDat(i).CrdCrdNum & Chr(5)
    '    sPrmValue = sPrmValue & tPrmData.OacCrdDat(i).CrdCrdApp & Chr(5)
    '    sPrmValue = sPrmValue & tPrmData.OacCrdDat(i).CrdAdpAmt & Chr(5)
    '    sPrmValue = sPrmValue & tPrmData.OacCrdDat(i).CrdUidCod & Chr(5)
    'Next
    
End Sub

    
    
    '======================================================
    ' "외래 내원 환자 정보" 의 자료를 OcmInfRec자료형으로 복사한다.
    '======================================================
Public Sub OcmInfLoad(sPrmValue As String, tPrmData As OcmInfRec)

    On Error GoTo OcmInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.OcmNum = vVal(i)
    i = i + 1
    tPrmData.OcmChtNum = vVal(i)
    i = i + 1
    tPrmData.OcmComStt = vVal(i)
    i = i + 1
    tPrmData.OcmDepCod = vVal(i)
    i = i + 1
    tPrmData.OcmDtrCod = vVal(i)
    i = i + 1
    tPrmData.OcmComRut = vVal(i)
    i = i + 1
    tPrmData.OcmAcpDtm = vVal(i)
    i = i + 1
    tPrmData.OcmInsCod = vVal(i)
    i = i + 1
    tPrmData.OcmInsSeq = vVal(i)
    i = i + 1
    tPrmData.OcmDgsNfs = vVal(i)
    i = i + 1
    tPrmData.OcmDgsDnh = vVal(i)
    i = i + 1
    tPrmData.OcmFreRsn = vVal(i)
    i = i + 1
    tPrmData.OcmSpcYon = vVal(i)
    i = i + 1
    tPrmData.OcmArtYon = vVal(i)
    i = i + 1
    tPrmData.OcmRsuYon = vVal(i)
    i = i + 1
    tPrmData.OcmDgsCht = vVal(i)
    i = i + 1
    tPrmData.OcmMdcDay = vVal(i)
    i = i + 1
    tPrmData.OcmNul = vVal(i)
    i = i + 1
    tPrmData.OcmArrStt = vVal(i)
    i = i + 1
    tPrmData.OcmLevTim = vVal(i)
    i = i + 1
    tPrmData.OcmLevRst = vVal(i)
    i = i + 1
    tPrmData.OcmEmgCod = vVal(i)
    i = i + 1
    tPrmData.OcmUpdTim = vVal(i)
    i = i + 1
    tPrmData.OcmUidCod = vVal(i)
    i = i + 1
    tPrmData.OcmNonIns = vVal(i)
    i = i + 1
    tPrmData.OcmIcmNum = vVal(i)
    i = i + 1
    tPrmData.OcmMdcTyp = vVal(i)
    i = i + 1
    tPrmData.OcmTrmDtr = vVal(i)
    i = i + 1
    tPrmData.OcmArrDtm = vVal(i)
    i = i + 1
    tPrmData.OcmImgYon = vVal(i)
    i = i + 1
    tPrmData.OcmEmgKnd = vVal(i)
    i = i + 1
    tPrmData.OcmActFlg = vVal(i)
    i = i + 1
    tPrmData.OcmEndStt = vVal(i)
    i = i + 1
    tPrmData.OcmHanAmt = vVal(i)
    i = i + 1
    tPrmData.OcmHanCmt = vVal(i)
    i = i + 1
    tPrmData.OcmPhyRev = vVal(i)
    i = i + 1
    tPrmData.OcmRomYon = vVal(i)
    i = i + 1
    tPrmData.OcmFutDay = vVal(i)
    i = i + 1
    tPrmData.OcmOutCod = vVal(i)
    i = i + 1
    tPrmData.OcmOutNum = vVal(i)
    i = i + 1
    tPrmData.OcmRcpCmt = vVal(i)
    i = i + 1
    tPrmData.OcmCvtYon = vVal(i)
    i = i + 1
    tPrmData.OcmCvtDtm = vVal(i)
    i = i + 1
    tPrmData.OcmCasStb = vVal(i)
    Exit Sub

OcmInfLoad_ErrorTraping:
    Resume Next

End Sub
    
    '=====================================================================
    ' 외래 내원 환자 정보
    '---------------------------------------------------------------------
    '   OcmInfRec 자료구조의 값으로 저장 형태인 Key, Value 생성
    '=====================================================================
Public Sub OcmInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OcmInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.OcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(CDouble(tPrmData.OcmChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmComStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmComRut & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmAcpDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OcmInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmDgsNfs & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmDgsDnh & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmFreRsn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmArtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmRsuYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmDgsCht & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmMdcDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmNul & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmArrStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmLevTim & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmLevRst & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmEmgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmUpdTim & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmNonIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmIcmNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmMdcTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmTrmDtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmArrDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmImgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmEmgKnd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmActFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmEndStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmHanAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmHanCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmPhyRev & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmRomYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmFutDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmOutCod & Chr(5)
    sPrmValue = sPrmValue & Format(tPrmData.OcmOutNum, "@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmRcpCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmCvtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmCvtDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OcmCasStb & Chr(5)
    
End Sub

    
Public Sub OdlInfLoad(sPrmValue As String, tPrmData As OdlInfRec)

    On Error GoTo OdlInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.OdlRcpNum = vVal(i)
    i = i + 1
    tPrmData.OdlIncCod = vVal(i)
    i = i + 1
    tPrmData.OdlChtNum = vVal(i)
    i = i + 1
    tPrmData.OdlInsCod = vVal(i)
    i = i + 1
    tPrmData.OdlInsSeq = vVal(i)
    i = i + 1
    tPrmData.OdlDepCod = vVal(i)
    i = i + 1
    tPrmData.OdlInsAct = vVal(i)
    i = i + 1
    tPrmData.OdlInsStf = vVal(i)
    i = i + 1
    tPrmData.OdlNonAct = vVal(i)
    i = i + 1
    tPrmData.OdlNonStf = vVal(i)
    i = i + 1
    tPrmData.OdlInsAmt = vVal(i)
    i = i + 1
    tPrmData.OdlNonAmt = vVal(i)
    i = i + 1
    tPrmData.OdlOwnAmt = vVal(i)
    i = i + 1
    tPrmData.OdlTotOwn = vVal(i)
    i = i + 1
    tPrmData.OdlSpcAmt = vVal(i)
    
    '20040101..HTS..add
    i = i + 1
    tPrmData.OdlNinAct = vVal(i)
    i = i + 1
    tPrmData.OdlNinStf = vVal(i)
    i = i + 1
    tPrmData.OdlNinAmt = vVal(i)
    
    Exit Sub

OdlInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OdlInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OdlInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.OdlRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.OdlIncCod), "@@") & Chr(5)
    
    sPrmValue = Format((tPrmData.OdlChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OdlInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlInsAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlInsStf & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlNonAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlNonStf & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlNonAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlOwnAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlTotOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlSpcAmt & Chr(5)
    
    '20040101..HTS..add
    sPrmValue = sPrmValue & tPrmData.OdlNinAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlNinStf & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OdlNinAmt & Chr(5)
    
End Sub

    
Public Sub OffInfLoad(sPrmValue As String, tPrmData As OffInfRec)

    On Error GoTo OffInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.OffChtNum = vVal(i)
    i = i + 1
    tPrmData.OffRelTyp = vVal(i)
    i = i + 1
    tPrmData.OffRelEmp = vVal(i)
    i = i + 1
    tPrmData.OffEmpNam = vVal(i)
    i = i + 1
    tPrmData.OffUidCod = vVal(i)
    
    Exit Sub

OffInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OffInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OffInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.OffChtNum), "@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.OffRelTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OffRelEmp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OffEmpNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OffUidCod & Chr(5)
    
End Sub

    
Public Sub OhtInfLoad(sPrmValue As String, tPrmData As OhtInfRec)

    On Error GoTo OhtInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.OhtRcpNum = vVal(i)
    i = i + 1
    tPrmData.OhtOcmNum = vVal(i)
    i = i + 1
    tPrmData.OhtRvnTyp = vVal(i)
    i = i + 1
    tPrmData.OhtChtNum = vVal(i)
    i = i + 1
    tPrmData.OhtDepCod = vVal(i)
    i = i + 1
    tPrmData.OhtDtrCod = vVal(i)
    i = i + 1
    tPrmData.OhtInsCod = vVal(i)
    i = i + 1
    tPrmData.OhtInsSeq = vVal(i)
    i = i + 1
    tPrmData.OhtRcpStt = vVal(i)
    i = i + 1
    tPrmData.OhtTotAmt = vVal(i)
    i = i + 1
    tPrmData.OhtInsAmt = vVal(i)
    i = i + 1
    tPrmData.OhtNonAmt = vVal(i)
    i = i + 1
    tPrmData.OhtCorAmt = vVal(i)
    i = i + 1
    tPrmData.OhtOwnAmt = vVal(i)
    i = i + 1
    tPrmData.OhtTotOwn = vVal(i)
    i = i + 1
    tPrmData.OhtSpcAmt = vVal(i)
    i = i + 1
    tPrmData.OhtDisAmt = vVal(i)
    i = i + 1
    tPrmData.OhtFutAmt = vVal(i)
    i = i + 1
    tPrmData.OhtAskAmt = vVal(i)
    i = i + 1
    tPrmData.OhtOldAmt = vVal(i)
    i = i + 1
    tPrmData.OhtNewAmt = vVal(i)
    i = i + 1
    tPrmData.OhtRcpYon = vVal(i)
    i = i + 1
    tPrmData.OhtRetRsn = vVal(i)
    i = i + 1
    tPrmData.OhtPubYon = vVal(i)
    i = i + 1
    tPrmData.OhtOldRcp = vVal(i)
    i = i + 1
    tPrmData.OhtOldNum = vVal(i)
    i = i + 1
    tPrmData.OhtManNum = vVal(i)
    i = i + 1
    tPrmData.OhtMdcNum = vVal(i)
    i = i + 1
    tPrmData.OhtBknDtm = vVal(i)
    i = i + 1
    tPrmData.OhtUpdDtm = vVal(i)
    i = i + 1
    tPrmData.OhtUidCod = vVal(i)
    i = i + 1
    tPrmData.OhtPrcFun = vVal(i)
    i = i + 1
    tPrmData.OhtMdcTyp = vVal(i)
    i = i + 1
    tPrmData.OhtDimAmt = vVal(i)
    i = i + 1
    tPrmData.OhtNonIns = vVal(i)
    i = i + 1
    tPrmData.OhtEtcDtl = vVal(i)
    i = i + 1
    tPrmData.OhtCarFut = vVal(i)
    i = i + 1
    tPrmData.OhtOutNum = vVal(i)
    
    i = i + 1
    tPrmData.OhtAccDte = vVal(i)
    
    i = i + 1
    tPrmData.OhtFodAmt = vVal(i)
    
    '20040101..HTS..add
    i = i + 1
    tPrmData.OhtNinAmt = vVal(i)
    
    Exit Sub

OhtInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OhtInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OhtInfRec)

    '''yk...20011031 : 3.0부터 접수 취소 영수증은 하나도 안써졌겠네...쩝
    'sPrmKey = Format(CDouble(tPrmData.OhtRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = Format(CDouble(tPrmData.OhtOldRcp), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(CDouble(tPrmData.OhtOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtRvnTyp & Chr(5)
    sPrmValue = sPrmValue & Format((tPrmData.OhtChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OhtInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtRcpStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtTotAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtNonAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtCorAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtOwnAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtTotOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtSpcAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtDisAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtFutAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtAskAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtOldAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtNewAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtRcpYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtRetRsn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtPubYon & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OhtOldRcp), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OhtOldNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OhtManNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtMdcNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtBknDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtPrcFun & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtMdcTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtDimAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtNonIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtEtcDtl & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtCarFut & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtOutNum & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.OhtAccDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OhtFodAmt & Chr(5)
    
    '20040101..HTS..add
    sPrmValue = sPrmValue & tPrmData.OhtNinAmt & Chr(5)
    
End Sub

    
Public Sub OicInfLoad(sPrmValue As String, tPrmData As OicInfRec)

    On Error GoTo OicInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.OicOcmNum = vVal(i)
    i = i + 1
    tPrmData.OicSeq = vVal(i)
    i = i + 1
    tPrmData.OicChtNum = vVal(i)
    i = i + 1
    tPrmData.OicIcdCod = vVal(i)
    i = i + 1
    tPrmData.OicIcdPri = vVal(i)
    i = i + 1
    tPrmData.OicEeeCod = vVal(i)
    i = i + 1
    tPrmData.OicVeeCod = vVal(i)
    i = i + 1
    tPrmData.OicOprYon = vVal(i)
    i = i + 1
    tPrmData.OicDenRgn = vVal(i)
    i = i + 1
    tPrmData.OicDgnDte = vVal(i)
    i = i + 1
    tPrmData.OicDepCod = vVal(i)
    i = i + 1
    tPrmData.OicCurRst = vVal(i)
    i = i + 1
    tPrmData.OicCurGrd = vVal(i)
    i = i + 1
    tPrmData.OicAddIcd = vVal(i)
    i = i + 1
    tPrmData.OicFinIcd = vVal(i)
    
    i = i + 1
    tPrmData.OicSpcCmt = vVal(i)
    
    i = i + 1   '20030228 lek add for 상병명 저장
    tPrmData.OicIdcNam = vVal(i)
    
    Exit Sub

OicInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OicInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OicInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.OicOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.OicSeq), "@@") & Chr(5)
    
    sPrmValue = Format((tPrmData.OicChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicIcdCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicIcdPri & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicEeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicVeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicOprYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicDenRgn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicDgnDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicCurRst & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicCurGrd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicAddIcd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicFinIcd & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OicSpcCmt & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.OicIdcNam & Chr(5) '20030228 lek add for 상병명 저장
    
    
End Sub

    
    
Public Sub OprInfLoad(sPrmValue As String, tPrmData As OprInfRec)

    On Error GoTo OprInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    
    i = i + 1
    tPrmData.OprNum = vVal(i)
    i = i + 1
    tPrmData.OprChtNum = vVal(i)
    i = i + 1
    tPrmData.OprOcmNum = vVal(i)
    i = i + 1
    tPrmData.OprCod = vVal(i)
    i = i + 1
    tPrmData.OprGbnCod = vVal(i)
    i = i + 1
    tPrmData.OprDte = vVal(i)
    i = i + 1
    tPrmData.OprTms = vVal(i)
    i = i + 1
    tPrmData.OprCfmYon = vVal(i)
    i = i + 1
    tPrmData.OprActYon = vVal(i)
    i = i + 1
    tPrmData.OprNarCod = vVal(i)
    i = i + 1
    tPrmData.OprIcdCod = vVal(i)
    i = i + 1
    tPrmData.OprDepCod = vVal(i)
    i = i + 1
    tPrmData.OprNssRom = vVal(i)
    i = i + 1
    tPrmData.OprEmgYon = vVal(i)
    i = i + 1
    tPrmData.OprManDtr = vVal(i)
    i = i + 1
    tPrmData.OprDtrCod = vVal(i)
    i = i + 1
    tPrmData.OprRqtEqp = vVal(i)
    i = i + 1
    tPrmData.OprSplCmt = vVal(i)
    i = i + 1
    tPrmData.OprRomCod = vVal(i)
    i = i + 1
    tPrmData.OprNarDtr = vVal(i)
    i = i + 1
    tPrmData.OprNrsCod = vVal(i)
    i = i + 1
    tPrmData.OprUpdDtm = vVal(i)
    i = i + 1
    tPrmData.OprUidCod = vVal(i)
    i = i + 1
    tPrmData.OprOldCod = vVal(i)
    i = i + 1
    tPrmData.OprUseTms = vVal(i)
    i = i + 1
    tPrmData.OprPbsQty = vVal(i)
    i = i + 1
    tPrmData.OprGazYon = vVal(i)
    i = i + 1
    tPrmData.OprPanCtr = vVal(i)
    i = i + 1
    tPrmData.OprIcdNam = vVal(i)
    i = i + 1
    tPrmData.OprStrTim = vVal(i)
    i = i + 1
    tPrmData.OprEndTim = vVal(i)
    i = i + 1
    tPrmData.OprPrtYon = vVal(i)
    i = i + 1
    tPrmData.OprBlood = vVal(i)    '혈액준비
    i = i + 1
    tPrmData.OprRstCxr = vVal(i) 'Chest x-ray
    i = i + 1
    tPrmData.OprRstEkg = vVal(i) 'EKG   판독결과
    i = i + 1
    tPrmData.OprNpoTms = vVal(i) '금식
    i = i + 1
    tPrmData.OprCodNam = vVal(i) '수술명칭
    
    Exit Sub

OprInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OprInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OprInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.OprNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(CDouble(tPrmData.OprChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OprOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprGbnCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprTms & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprCfmYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprActYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprNarCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprIcdCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprNssRom & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprEmgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprManDtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprRqtEqp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprRomCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprNarDtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprNrsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprOldCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprUseTms & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprPbsQty & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprGazYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprPanCtr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprIcdNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprStrTim & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprEndTim & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprPrtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprBlood & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OprRstCxr & Chr(5) 'Chest x-ray
    sPrmValue = sPrmValue & tPrmData.OprRstEkg & Chr(5) 'EKG   판독결과
    sPrmValue = sPrmValue & tPrmData.OprNpoTms & Chr(5) '금식
    sPrmValue = sPrmValue & tPrmData.OprCodNam & Chr(5) '수술명칭
    
End Sub

    
Public Sub OprTypLoad(sPrmValue As String, tPrmData As OprTypRec)

    On Error GoTo OprTypLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.OprNum = vVal(i)
    i = i + 1
    tPrmData.OprCodTyp = vVal(i)
    i = i + 1
    tPrmData.OprTypCod = vVal(i)
    i = i + 1
    tPrmData.OprTypNam = vVal(i)
    
    Exit Sub

OprTypLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OprTypStore(sPrmKey As String, sPrmValue As String, tPrmData As OprTypRec)

    
    sPrmKey = Format(CDouble(tPrmData.OprNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.OprCodTyp & Chr(5)
    sPrmKey = sPrmKey & tPrmData.OprTypCod & Chr(5)
    
    sPrmValue = tPrmData.OprTypNam & Chr(5)
    
End Sub

Public Sub TpmInfStore(sPrmKey As String, sPrmValue As String, tPrmTpmData As TpmInfRec)
    
    sPrmKey = tPrmTpmData.TpmCytoNum & Chr(5)
    sPrmKey = sPrmKey & Format(Trim(tPrmTpmData.TpmCodSeq), "@@") & Chr(5)
        
    sPrmValue = sPrmValue & tPrmTpmData.TpmCodDat & Chr(5)

End Sub
    
Public Sub TpmInfLoad(sPrmValue As String, tPrmTpmData As TpmInfRec)

    On Error GoTo TpmInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmTpmData.TpmCytoNum = vVal(i)
    i = i + 1
    tPrmTpmData.TpmCodSeq = vVal(i)
    i = i + 1
    tPrmTpmData.TpmCodDat = vVal(i)
        
    Exit Sub

TpmInfLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub OrgInfLoad(sPrmValue As String, tPrmData As OrgInfRec)

    On Error GoTo OrgInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.OrgChtNum = vVal(i)
    i = i + 1
    tPrmData.OrgInsCod = vVal(i)
    i = i + 1
    tPrmData.OrgRcuNum = vVal(i)
    i = i + 1
    tPrmData.OrgRcuTyp = vVal(i)
    i = i + 1
    tPrmData.OrgAdpDte = vVal(i)
    i = i + 1
    tPrmData.OrgExpDte = vVal(i)
    i = i + 1
    tPrmData.OrgDepCod = vVal(i)
    i = i + 1
    tPrmData.OrgIcdCod = vVal(i)
    
    Exit Sub

OrgInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OrgInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OrgInfRec)

    
    sPrmKey = tPrmData.OrgChtNum & Chr(5)
    sPrmKey = sPrmKey & tPrmData.OrgInsCod & Chr(5)
    
    sPrmValue = tPrmData.OrgRcuNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrgRcuTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrgAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrgExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrgDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrgIcdCod & Chr(5)
    
End Sub
   
Public Sub OrpInfLoad(sPrmValue As String, tPrmData As OrpInfRec)

    On Error GoTo OrpInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.OrpOcmNum = vVal(i)
    i = i + 1
    tPrmData.OrpRvnTyp = vVal(i)
    i = i + 1
    tPrmData.OrpChtNum = vVal(i)
    i = i + 1
    tPrmData.OrpDepCod = vVal(i)
    i = i + 1
    tPrmData.OrpDtrCod = vVal(i)
    i = i + 1
    tPrmData.OrpInsCod = vVal(i)
    i = i + 1
    tPrmData.OrpInsSeq = vVal(i)
    i = i + 1
    tPrmData.OrpRcpStt = vVal(i)
    i = i + 1
    tPrmData.OrpTotAmt = vVal(i)
    i = i + 1
    tPrmData.OrpInsAmt = vVal(i)
    i = i + 1
    tPrmData.OrpNonAmt = vVal(i)
    i = i + 1
    tPrmData.OrpCorAmt = vVal(i)
    i = i + 1
    tPrmData.OrpOwnAmt = vVal(i)
    i = i + 1
    tPrmData.OrpTotOwn = vVal(i)
    i = i + 1
    tPrmData.OrpSpcAmt = vVal(i)
    i = i + 1
    tPrmData.OrpDisAmt = vVal(i)
    i = i + 1
    tPrmData.OrpFutAmt = vVal(i)
    i = i + 1
    tPrmData.OrpAskAmt = vVal(i)
    i = i + 1
    tPrmData.OrpOldAmt = vVal(i)
    i = i + 1
    tPrmData.OrpNewAmt = vVal(i)
    i = i + 1
    tPrmData.OrpRcpYon = vVal(i)
    i = i + 1
    tPrmData.OrpRetRsn = vVal(i)
    i = i + 1
    tPrmData.OrpPubYon = vVal(i)
    i = i + 1
    tPrmData.OrpRcpNum = vVal(i)
    i = i + 1
    tPrmData.OrpOldNum = vVal(i)
    i = i + 1
    tPrmData.OrpManNum = vVal(i)
    i = i + 1
    tPrmData.OrpMdcNum = vVal(i)
    i = i + 1
    tPrmData.OrpBknDtm = vVal(i)
    i = i + 1
    tPrmData.OrpUpdDtm = vVal(i)
    i = i + 1
    tPrmData.OrpUidCod = vVal(i)
    i = i + 1
    tPrmData.OrpPrcFun = vVal(i)
    i = i + 1
    tPrmData.OrpMdcTyp = vVal(i)
    i = i + 1
    tPrmData.OrpDimAmt = vVal(i)
    i = i + 1
    tPrmData.OrpNonIns = vVal(i)
    i = i + 1
    tPrmData.OrpEtcDtl = vVal(i)
    i = i + 1
    tPrmData.OrpCarFut = vVal(i)
    i = i + 1
    tPrmData.OrpOutNum = vVal(i)
    i = i + 1
    tPrmData.OrpAccDte = vVal(i)
    i = i + 1
    tPrmData.OrpFodAmt = vVal(i)
    i = i + 1
    tPrmData.OrpNinAmt = vVal(i)
    
    Exit Sub

OrpInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OrpInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OrpInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.OrpOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.OrpRvnTyp & Chr(5)
    
    sPrmValue = Format(CDouble(tPrmData.OrpChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OrpInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpRcpStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpTotAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpNonAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpCorAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpOwnAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpTotOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpSpcAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpDisAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpFutAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpAskAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpOldAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpNewAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpRcpYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpRetRsn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpPubYon & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OrpRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OrpOldNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.OrpManNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpMdcNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpBknDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpPrcFun & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpMdcTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpDimAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpNonIns & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpEtcDtl & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpCarFut & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpOutNum & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.OrpAccDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OrpFodAmt & Chr(5)
        
    '20040101..HTS..add
    sPrmValue = sPrmValue & tPrmData.OrpNinAmt & Chr(5)
    
End Sub

    
Public Sub OspInfLoad(sPrmValue As String, tPrmData As OspInfRec)

    On Error GoTo OspInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.OspOcmNum = vVal(i)
    i = i + 1
    tPrmData.OspOdrNum = vVal(i)
    i = i + 1
    tPrmData.OspOdrSeq = vVal(i)
    i = i + 1
    tPrmData.OspOdrCod = vVal(i)
    i = i + 1
    tPrmData.OspOdrTyp = vVal(i)
    i = i + 1
    tPrmData.OspOdrStt = vVal(i)
    i = i + 1
    tPrmData.OspStkStt = vVal(i)
    i = i + 1
    tPrmData.OspFeeCod = vVal(i)
    i = i + 1
    tPrmData.OspAddCod = vVal(i)
    i = i + 1
    tPrmData.OspDepCod = vVal(i)
    i = i + 1
    tPrmData.OspSlpDep = vVal(i)
    i = i + 1
    tPrmData.OspSlpCod = vVal(i)
    i = i + 1
    tPrmData.OspItmCod = vVal(i)
    i = i + 1
    tPrmData.OspOdrDtm = vVal(i)
    i = i + 1
    tPrmData.OspOdrPrc = vVal(i)
    i = i + 1
    tPrmData.OspOdrSib = vVal(i)
    i = i + 1
    tPrmData.OspOdrQty = vVal(i)
    i = i + 1
    tPrmData.OspOdrTms = vVal(i)
    i = i + 1
    tPrmData.OspOdrDay = vVal(i)
    i = i + 1
    tPrmData.OspUsgCod = vVal(i)
    i = i + 1
    tPrmData.OspMthCod = vVal(i)
    i = i + 1
    tPrmData.OspSpmcod = vVal(i)
    i = i + 1
    tPrmData.OspInsYon = vVal(i)
    i = i + 1
    tPrmData.OspInsCod = vVal(i)
    i = i + 1
    tPrmData.OspInsSeq = vVal(i)
    i = i + 1
    tPrmData.OspDgsEtc = vVal(i)
    i = i + 1
    tPrmData.OspDgsRol = vVal(i)
    i = i + 1
    tPrmData.OspOprDnh = vVal(i)
    i = i + 1
    tPrmData.OspOprDtm = vVal(i)
    i = i + 1
    tPrmData.OspPrePay = vVal(i)
    i = i + 1
    tPrmData.OspEmgYon = vVal(i)
    i = i + 1
    tPrmData.OspSpcYon = vVal(i)
    i = i + 1
    tPrmData.OspSlpAmt = vVal(i)
    i = i + 1
    tPrmData.OspIncCod = vVal(i)
    i = i + 1
    tPrmData.OspSotCod = vVal(i)
    i = i + 1
    tPrmData.OspEntDtm = vVal(i)
    i = i + 1
    tPrmData.OspDtrCod = vVal(i)
    i = i + 1
    tPrmData.OspCasYon = vVal(i)
    i = i + 1
    tPrmData.OspCasDtm = vVal(i)
    i = i + 1
    tPrmData.OspUidCod = vVal(i)
    i = i + 1
    tPrmData.OspUpdDte = vVal(i)
    i = i + 1
    tPrmData.OspPreDtm = vVal(i)
    i = i + 1
    tPrmData.OspSplYon = vVal(i)
    i = i + 1
    tPrmData.OspSplCmt = vVal(i)
    i = i + 1
    tPrmData.OspChkStt = vVal(i)
    i = i + 1
    tPrmData.OspMdcNum = vVal(i)
    i = i + 1
    tPrmData.OspCanMdc = vVal(i)
    i = i + 1
    tPrmData.OspStgCod = vVal(i)
    i = i + 1
    tPrmData.OspBasUnt = vVal(i)
    i = i + 1
    tPrmData.OspMntUsg = vVal(i)
    i = i + 1
    tPrmData.OspXryPtb = vVal(i)
    i = i + 1
    tPrmData.OspDtrPrt = vVal(i)
    i = i + 1
    tPrmData.OspUpdPrt = vVal(i)
    i = i + 1
    tPrmData.OspImgYon = vVal(i)
    i = i + 1
    tPrmData.OspCanNum = vVal(i)
    i = i + 1
    tPrmData.OspCanSeq = vVal(i)
    i = i + 1
    tPrmData.OspDenRgn = vVal(i)            '코드별 치식 02.03.21 sebal
    i = i + 1
    tPrmData.OspOdrNam = vVal(i)            '지시형 오더 명칭
    i = i + 1
    tPrmData.OspOdrNo = vVal(i)            '지시형 오더 명칭
    i = i + 1
    tPrmData.OspQtyGbn = vVal(i)
    Exit Sub

OspInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OspInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OspInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.OspOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.OspOdrNum), "@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.OspOdrSeq), "@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.OspOdrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspStkStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspFeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspAddCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspSlpDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspSlpCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspItmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrPrc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrSib & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrQty & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrTms & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrDay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspUsgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspMthCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspSpmcod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspInsYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspInsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspInsSeq & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspDgsEtc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspDgsRol & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOprDnh & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOprDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspPrePay & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspEmgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspSlpAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspIncCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspSotCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspEntDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspCasYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspCasDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspUpdDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspPreDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspSplYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspChkStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspMdcNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspCanMdc & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspStgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspBasUnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspMntUsg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspXryPtb & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspDtrPrt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspUpdPrt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspImgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspCanNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspCanSeq & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspDenRgn & Chr(5)         '코드별 치식. 02.03.21 sebal
    sPrmValue = sPrmValue & tPrmData.OspOdrNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspOdrNo & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OspQtyGbn & Chr(5)
    
End Sub

    
Public Sub OutInfLoad(sPrmValue As String, tPrmOutData As OutInfRec)

    On Error GoTo OutInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmOutData.OutOdrDte = vVal(i)
    i = i + 1
    tPrmOutData.OutNum = vVal(i)
    i = i + 1
    tPrmOutData.OutOcmNum = vVal(i)
    i = i + 1
    tPrmOutData.OutOdrNum = vVal(i)
    i = i + 1
    tPrmOutData.OutOdrSeq = vVal(i)
    i = i + 1
    tPrmOutData.OutChtNum = vVal(i)
    i = i + 1
    tPrmOutData.OutOdrStt = vVal(i)
    i = i + 1
    tPrmOutData.OutUpdTms = vVal(i)
    i = i + 1
    tPrmOutData.OutPatNam = vVal(i)
    i = i + 1
    tPrmOutData.OutDepCod = vVal(i)
    i = i + 1
    tPrmOutData.OutResNum = vVal(i)
    i = i + 1                               '         <=추가
    'i = i + 1
    tPrmOutData.OutCanNum = vVal(i)
    
    Exit Sub

OutInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OutInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OutInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.OutOdrDte), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.OutNum), "@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmData.OutOcmNum, "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmData.OutOdrNum, "@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmData.OutOdrSeq, "@@@@@") & Chr(5)
    
    sPrmValue = Format(tPrmData.OutChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OutOdrStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OutUpdTms & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OutPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OutDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OutResNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.OutCanNum & Chr(5)
    
End Sub
Public Sub OscInfLoad(sPrmValue As String, tPrmOscData As OscInfRec)

    On Error GoTo OscInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    'Key
    i = i + 1
    tPrmOscData.OscChtNum = vVal(i)
    'value - 원무과 참조사항
    i = i + 1
    tPrmOscData.OscSplCmt = vVal(i)
    Exit Sub

OscInfLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub OscInfStore(sPrmKey As String, sPrmValue As String, tPrmData As OscInfRec)

    'Key
    sPrmKey = Format(CDouble(tPrmData.OscChtNum), "@@@@@@@@") & Chr(5)
    'value - 원무과 참조사항
    sPrmValue = tPrmData.OscSplCmt & Chr(5)
    
End Sub

    
    '*******************************************************************
    ' PbsInf(PbsInfRec) Data Load : 기본 인적사항                      *
    '*******************************************************************
Public Sub PbsInfLoad(sPrmValue As String, tPrmPbsData As PbsInfRec)

    On Error GoTo PbsInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmPbsData.PbsChtNum = vVal(i)
    i = i + 1
    tPrmPbsData.PbsPatNam = vVal(i)
    i = i + 1
    tPrmPbsData.PbsResNum = vVal(i)
    i = i + 1
    tPrmPbsData.PbsZipCod = vVal(i)
    i = i + 1
    tPrmPbsData.PbsDtlAdr = vVal(i)
    i = i + 1
    tPrmPbsData.PbsPhnNum = vVal(i)
    i = i + 1
    tPrmPbsData.PbsNewDte = vVal(i)
    i = i + 1
    tPrmPbsData.PbsMdcTyp = vVal(i)
    i = i + 1
    tPrmPbsData.PbsSexTyp = vVal(i)
    i = i + 1
    tPrmPbsData.PbsArtYon = vVal(i)
    i = i + 1
    tPrmPbsData.PbsSpcFlg = vVal(i)
    i = i + 1
    tPrmPbsData.PbsRefCmd = vVal(i)
    i = i + 1
    tPrmPbsData.PbsUpdTim = vVal(i)
    i = i + 1
    tPrmPbsData.PbsUidCod = vVal(i)
    i = i + 1
    tPrmPbsData.PbsOsuYon = vVal(i)
    i = i + 1
    tPrmPbsData.PbsOldNum = vVal(i)
    i = i + 1
    tPrmPbsData.PbsIcmYon = vVal(i)
    i = i + 1
    tPrmPbsData.PbsCruYon = vVal(i)
    i = i + 1
    tPrmPbsData.PbsHndPhn = vVal(i) '          19 핸드폰 번호
    i = i + 1
    tPrmPbsData.PbsE_Mail = vVal(i) '          20 E-Mail
    i = i + 1
    tPrmPbsData.PbsMomCht = vVal(i)
    i = i + 1
    tPrmPbsData.PbsRecUid = vVal(i) '          22 추천인 아이디
    i = i + 1
    tPrmPbsData.PbsRecNam = vVal(i) '           23 추천인 성명
    i = i + 1
    tPrmPbsData.PbsPatDte = vVal(i)
    
    
    Exit Sub

PbsInfLoad_ErrorTraping:
    Resume Next

End Sub
    
    '*******************************************************************
    ' PbsInf(PbsInfRec) Data Store : 기본 인적사항                     *
    '*******************************************************************
Public Sub PbsInfStore(sPrmKey As String, sPrmValue As String, tPrmPbsData As PbsInfRec)

    sPrmKey = Format(CDouble(tPrmPbsData.PbsChtNum), "@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmPbsData.PbsPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsResNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsZipCod & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsDtlAdr & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsPhnNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsNewDte & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsMdcTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsSexTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsArtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsSpcFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsRefCmd & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsUpdTim & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsOsuYon & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsOldNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsIcmYon & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsCruYon & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsHndPhn & Chr(5)      '핸드폰 번호
    sPrmValue = sPrmValue & tPrmPbsData.PbsE_Mail & Chr(5)      'E-Mail
    sPrmValue = sPrmValue & tPrmPbsData.PbsMomCht & Chr(5)
    sPrmValue = sPrmValue & tPrmPbsData.PbsRecUid & Chr(5)      '추천인 아이디
    sPrmValue = sPrmValue & tPrmPbsData.PbsRecNam & Chr(5)      '추천인 성명
    sPrmValue = sPrmValue & tPrmPbsData.PbsPatDte & Chr(5)
    
End Sub

Public Sub GrnInfLoad(sPrmValue As String, tPrmGrnData As GrnInfRec)

    On Error GoTo GrnInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
   
    i = i + 1
    tPrmGrnData.GrnOcmNum = vVal(i)
    i = i + 1
    tPrmGrnData.GrnChtNum = vVal(i)
    i = i + 1
    tPrmGrnData.GrnPatNam = vVal(i)
    i = i + 1
    tPrmGrnData.GrnResNum = vVal(i)
    i = i + 1
    tPrmGrnData.GrnDtlAdr = vVal(i)
    i = i + 1
    tPrmGrnData.GrnPhnNum = vVal(i)
    i = i + 1
    tPrmGrnData.GrnSexTyp = vVal(i)
    i = i + 1
    tPrmGrnData.GrnRelTyp = vVal(i)
    i = i + 1
    tPrmGrnData.GrnComNam = vVal(i)
    i = i + 1
    tPrmGrnData.GrnPrtNam = vVal(i)
    i = i + 1
    tPrmGrnData.GrnComTel = vVal(i)
    i = i + 1
    tPrmGrnData.GrnEtc = vVal(i)

    Exit Sub

GrnInfLoad_ErrorTraping:
    Resume Next

End Sub
'*******************************************************************
' GrnInf(GrnInfRec) Data Store : 연대 보증인 인적사항                     *
'*******************************************************************
Public Sub GrnInfStore(sPrmKey As String, sPrmValue As String, tPrmGrnData As GrnInfRec)
    
    sPrmKey = Format(CDouble(tPrmGrnData.GrnOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format(CDouble(tPrmGrnData.GrnChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnResNum & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnDtlAdr & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnPhnNum & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnSexTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnRelTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnComNam & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnPrtNam & Chr(5)
    sPrmValue = sPrmValue & tPrmGrnData.GrnComTel & Chr(5)
    'EverSky 지우지 말것
    sPrmValue = sPrmValue & tPrmGrnData.GrnEtc & Chr(5)
End Sub

    '*******************************************************************
    ' PcrInf(PcrInfRec) Data Load : 자보 보험 인적사항                 *
    '*******************************************************************
Public Sub PcrInfLoad(sPrmValue As String, tPrmPcrData As PcrInfRec)

    On Error GoTo PcrInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmPcrData.PcrChtNum = vVal(i)
    i = i + 1
    tPrmPcrData.PcrInsCod = vVal(i)
    i = i + 1
    tPrmPcrData.PcrInsSeq = vVal(i)
    i = i + 1
    tPrmPcrData.PcrInjDte = vVal(i)
    i = i + 1
    tPrmPcrData.PcrFstDte = vVal(i)
    i = i + 1
    tPrmPcrData.PcrAssCod = vVal(i)
    i = i + 1
    tPrmPcrData.PcrVeiNum = vVal(i)
    i = i + 1
    tPrmPcrData.PcrVeiOwn = vVal(i)
    i = i + 1
    tPrmPcrData.PcrAcpNum = vVal(i)
    i = i + 1
    tPrmPcrData.PcrAcpDte = vVal(i)
    i = i + 1
    tPrmPcrData.PcrCarUid = vVal(i)
    i = i + 1
    tPrmPcrData.PcrAdpDte = vVal(i)
    i = i + 1
    tPrmPcrData.PcrExpDte = vVal(i)
    i = i + 1
    tPrmPcrData.PcrUpdDtm = vVal(i)
    i = i + 1
    tPrmPcrData.PcrUidCod = vVal(i)
    i = i + 1
    tPrmPcrData.PcrCarRem = vVal(i)
    i = i + 1
    tPrmPcrData.PcrLmtAmt = vVal(i)
    
    Exit Sub

PcrInfLoad_ErrorTraping:
    Resume Next

End Sub
    
    '*******************************************************************
    ' PcrInf(PcrInfRec) Data Store : 자보 보험 인적사항                *
    '*******************************************************************
Public Sub PcrInfStore(sPrmKey As String, sPrmValue As String, tPrmPcrData As PcrInfRec)

    
    sPrmKey = Format(CDouble(tPrmPcrData.PcrChtNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmPcrData.PcrInsCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmPcrData.PcrInsSeq), "@@") & Chr(5)
    
    sPrmValue = tPrmPcrData.PcrInjDte & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrFstDte & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrAssCod & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrVeiNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrVeiOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrAcpNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrAcpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrCarUid & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrCarRem & Chr(5)
    sPrmValue = sPrmValue & tPrmPcrData.PcrLmtAmt & Chr(5)
    
End Sub

    
    '*******************************************************************
    ' PmdInf(PmdInfRec) Data Load : 보험 인적사항                      *
    '*******************************************************************
Public Sub PmdInfLoad(sPrmValue As String, tPrmPmdData As PmdInfRec)

    On Error GoTo PmdInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmPmdData.PmdChtNum = vVal(i)
    i = i + 1
    tPrmPmdData.PmdInsCod = vVal(i)
    i = i + 1
    tPrmPmdData.PmdInsSeq = vVal(i)
    i = i + 1
    tPrmPmdData.PmdAssCod = vVal(i)
    i = i + 1
    tPrmPmdData.PmdInsNum = vVal(i)
    i = i + 1
    tPrmPmdData.PmdPasNam = vVal(i)
    i = i + 1
    tPrmPmdData.PmdResNum = vVal(i)
    i = i + 1
    tPrmPmdData.PmdRelTyp = vVal(i)
    i = i + 1
    tPrmPmdData.PmdDsoNum = vVal(i)
    i = i + 1
    tPrmPmdData.PmdXplNum = vVal(i)
    i = i + 1
    tPrmPmdData.PmdRcuNum = vVal(i)
    i = i + 1
    tPrmPmdData.PmdRgnYon = vVal(i)
    i = i + 1
    tPrmPmdData.PmdAdpDte = vVal(i)
    i = i + 1
    tPrmPmdData.PmdExpDte = vVal(i)
    i = i + 1
    tPrmPmdData.PmdEntNam = vVal(i)
    i = i + 1
    tPrmPmdData.PmdUpdDtm = vVal(i)
    i = i + 1
    tPrmPmdData.PmdUidCod = vVal(i)
    i = i + 1
    tPrmPmdData.PmdXplAss = vVal(i)
    
    Exit Sub

PmdInfLoad_ErrorTraping:
    Resume Next

End Sub
    
    '*******************************************************************
    ' PmdInf(PmdInfRec) Data Store : 보험 인적사항                     *
    '*******************************************************************
Public Sub PmdInfStore(sPrmKey As String, sPrmValue As String, tPrmPmdData As PmdInfRec)

    
    sPrmKey = Format(CDouble(tPrmPmdData.PmdChtNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmPmdData.PmdInsCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmPmdData.PmdInsSeq), "@@") & Chr(5)
    
    sPrmValue = tPrmPmdData.PmdAssCod & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdInsNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdPasNam & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdResNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdRelTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdDsoNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdXplNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdRcuNum & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdRgnYon & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdEntNam & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmPmdData.PmdXplAss & Chr(5)
    
End Sub

    
Public Sub PspInfLoad(sPrmValue As String, tPrmData As PspInfRec)
    On Error GoTo PspInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    '
    i = i + 1
    tPrmData.PspChtNum = vVal(i)
    i = i + 1
    tPrmData.PspParNam = vVal(i)
    i = i + 1
    tPrmData.PspRelTyp = vVal(i)
    i = i + 1
    tPrmData.PspPhnNum = vVal(i)
    i = i + 1
    tPrmData.PspPatEdu = vVal(i)
    i = i + 1
    tPrmData.PspMryYon = vVal(i)
    i = i + 1
    tPrmData.PspParYon = vVal(i)
    i = i + 1
    tPrmData.PspPatJob = vVal(i)
    i = i + 1
    tPrmData.PspPatRlg = vVal(i)
    i = i + 1
    tPrmData.PspUpdTim = vVal(i)
    i = i + 1
    tPrmData.PspUidCod = vVal(i)

    i = i + 1
    tPrmData.PspZipCod = vVal(i)
    i = i + 1
    tPrmData.PspDtlAdr = vVal(i)
    i = i + 1
    tPrmData.PspSchCod = vVal(i)
    i = i + 1
    tPrmData.PspDtlJob = vVal(i)
    i = i + 1
    tPrmData.PspMrgCod = vVal(i)

    i = i + 1
    tPrmData.PspComPhn = vVal(i)
    i = i + 1
    tPrmData.PspPcsPhn = vVal(i)
    i = i + 1
    tPrmData.PspCslUid = vVal(i)
    i = i + 1
    tPrmData.PspCstUid = vVal(i)

    Exit Sub

PspInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub PspInfStore(sPrmKey As String, sPrmValue As String, tPrmData As PspInfRec)
    
    sPrmKey = Format((tPrmData.PspChtNum), "@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.PspParNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspRelTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspPhnNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspPatEdu & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspMryYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspParYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspPatJob & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspPatRlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspUpdTim & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspUidCod & Chr(5)

    sPrmValue = sPrmValue & tPrmData.PspZipCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspDtlAdr & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspSchCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspDtlJob & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspMrgCod & Chr(5)

    sPrmValue = sPrmValue & tPrmData.PspComPhn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspPcsPhn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspCslUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.PspCstUid & Chr(5)
    
End Sub

    
    '*******************************************************************
    ' PwkInf(PwkInfRec) Data Load : 산재 인적사항                      *
    '*******************************************************************
Public Sub PwkInfLoad(sPrmValue As String, tprmPwkData As PwkInfRec)

    On Error GoTo PwkInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tprmPwkData.PwkChtNum = vVal(i)
    i = i + 1
    tprmPwkData.PwkInsCod = vVal(i)
    i = i + 1
    tprmPwkData.PwkInsSeq = vVal(i)
    i = i + 1
    tprmPwkData.PwkAssCod = vVal(i)
    i = i + 1
    tprmPwkData.PwkEntNam = vVal(i)
    i = i + 1
    tprmPwkData.PwkSetCod = vVal(i)
    i = i + 1
    tprmPwkData.PwkRcuNum = vVal(i)
    i = i + 1
    tprmPwkData.PwkRcuDte = vVal(i)
    i = i + 1
    tprmPwkData.PwkDsaDte = vVal(i)
    i = i + 1
    tprmPwkData.PwkMcrDte = vVal(i)
    i = i + 1
    tprmPwkData.PwkReqRcu = vVal(i)
    i = i + 1
    tprmPwkData.PwkInjRgn = vVal(i)
    i = i + 1
    tprmPwkData.PwkCurRst = vVal(i)
    i = i + 1
    tprmPwkData.PwkAdpDte = vVal(i)
    i = i + 1
    tprmPwkData.PwkExpDte = vVal(i)
    i = i + 1
    tprmPwkData.PwkUpdDtm = vVal(i)
    i = i + 1
    tprmPwkData.PwkUidCod = vVal(i)
    
    Exit Sub

PwkInfLoad_ErrorTraping:
    Resume Next

End Sub
    
    '*******************************************************************
    ' PwkInf(PwkInfRec) Data Store : 산재 인적사항                     *
    '*******************************************************************
Public Sub PwkInfStore(sPrmKey As String, sPrmValue As String, tprmPwkData As PwkInfRec)

    
    sPrmKey = Format(CDouble(tprmPwkData.PwkChtNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tprmPwkData.PwkInsCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tprmPwkData.PwkInsSeq), "@@") & Chr(5)
    
    sPrmValue = tprmPwkData.PwkAssCod & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkEntNam & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkSetCod & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkRcuNum & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkRcuDte & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkDsaDte & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkMcrDte & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkReqRcu & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkInjRgn & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkCurRst & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkAdpDte & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkExpDte & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tprmPwkData.PwkUidCod & Chr(5)
    
End Sub

    
Public Sub RcsInfLoad(sPrmValue As String, tPrmData As RcsInfRec)

    On Error GoTo RcsInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.RcsDte = vVal(i)
    i = i + 1
    tPrmData.RcsOcmNum = vVal(i)
    i = i + 1
    tPrmData.RcsCod = vVal(i)
    i = i + 1
    tPrmData.RcsStt = vVal(i)
    i = i + 1
    tPrmData.RcsSlpDep = vVal(i)
    
    Exit Sub

RcsInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub RcsInfStore(sPrmKey As String, sPrmValue As String, tPrmData As RcsInfRec)

    
    sPrmKey = tPrmData.RcsDte & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmData.RcsOcmNum, "@@@@@@@@@@") & Chr(5)
    sPrmValue = tPrmData.RcsCod & Chr(5)
    
    sPrmValue = sPrmValue & tPrmData.RcsStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RcsSlpDep & Chr(5)
    
End Sub

    
Public Sub RctInfLoad(sPrmValue As String, tPrmData As RctInfRec)

    On Error GoTo RctInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.RctCod = vVal(i)
    i = i + 1
    tPrmData.RctDte = vVal(i)
    i = i + 1
    tPrmData.RctTotCnt = vVal(i)
    i = i + 1
    tPrmData.RctCurCnt = vVal(i)
    
    Exit Sub

RctInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub RctInfStore(sPrmKey As String, sPrmValue As String, tPrmData As RctInfRec)

    
    sPrmKey = tPrmData.RctCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.RctDte & Chr(5)
    
    sPrmValue = tPrmData.RctTotCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RctCurCnt & Chr(5)
    
End Sub

    
Public Sub RefInfLoad(sPrmValue As String, tPrmData As RefInfRec)

    On Error GoTo RefInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.RefOcmNum = vVal(i)
    i = i + 1
    tPrmData.RefSplCmt = vVal(i)
    
    Exit Sub

RefInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub RefInfStore(sPrmKey As String, sPrmValue As String, tPrmData As RefInfRec)

    
    sPrmKey = Format((tPrmData.RefOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.RefSplCmt & Chr(5)
    
End Sub

    
Public Sub ResInfLoad(sRetVal As String, tData As ResInfRec)

    On Error GoTo ResInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sRetVal = "" Then Exit Sub

    vVal = Split(sRetVal & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tData.ResAcpDte = vVal(i)
    i = i + 1
    tData.ResAcpCod = vVal(i)
    i = i + 1
    tData.ResAcpNum = vVal(i)
    i = i + 1
    tData.ResSpmCod = vVal(i)
    i = i + 1
    tData.ResLabCod = vVal(i)
    i = i + 1
    tData.ResOcmNum = vVal(i)
    i = i + 1
    tData.ResChtNum = vVal(i)
    i = i + 1
    tData.ResSotCod1 = vVal(i)
    i = i + 1
    tData.ResSotCod2 = vVal(i)
    i = i + 1
    tData.ResJbsSeq = vVal(i)
    i = i + 1
    tData.ResRltSeq = vVal(i)
    i = i + 1
    tData.ResMzhMin = vVal(i)
    i = i + 1
    tData.ResMzhMax = vVal(i)
    i = i + 1
    tData.ResMzhRef = vVal(i)
    i = i + 1
    tData.ResMzhUnt = vVal(i)
    i = i + 1
    tData.ResMzhMnt = vVal(i)
    i = i + 1
    tData.ResSplCmt = vVal(i)
    i = i + 1
    tData.ResOdrDtm = vVal(i)
    i = i + 1
    tData.ResAcpDtm = vVal(i)
    i = i + 1
    tData.ResTstDtm = vVal(i)
    i = i + 1
    tData.ResUpdDtm = vVal(i)
    i = i + 1
    tData.ResOdrUid = vVal(i)
    i = i + 1
    tData.ResAcpUid = vVal(i)
    i = i + 1
    tData.ResTstUid = vVal(i)
    i = i + 1
    tData.ResUpdUid = vVal(i)
    i = i + 1
    tData.ResSeeYon = vVal(i)
    i = i + 1
    tData.ResConLvl = vVal(i)
    i = i + 1
    tData.ResMzhTyp = vVal(i)
    i = i + 1
    tData.ResMzhLin = vVal(i)
    i = i + 1
    tData.ResShtNam = vVal(i)
    i = i + 1
    tData.ResSclCod = vVal(i)
    i = i + 1
    tData.ResPrtYon = vVal(i)
    i = i + 1
    tData.ResOspIsp = vVal(i)
    i = i + 1
    tData.ResMadYon = vVal(i)
    i = i + 1
    tData.ResJbsMth = vVal(i)
    i = i + 1
    tData.ResJbsQty = vVal(i)
    i = i + 1
    tData.ResCasYon = vVal(i)
    i = i + 1
    tData.ResEmgYon = vVal(i)
    i = i + 1
    tData.ResOdrNum = vVal(i)
    i = i + 1
    tData.ResOdrSeq = vVal(i)
    i = i + 1
    tData.ResOdrDep = vVal(i)
    i = i + 1
    tData.ResWrdCod = vVal(i)
    i = i + 1
    tData.ResStaYon = vVal(i)
    i = i + 1
    tData.ResOkYon = vVal(i)
    i = i + 1
    tData.ResOkUid = vVal(i)
    i = i + 1
    tData.ResOkDtm = vVal(i)
    i = i + 1
    tData.ResRedYon = vVal(i)
    i = i + 1
    tData.ResRedUid = vVal(i)
    i = i + 1
    tData.ResRedDtm = vVal(i)
    i = i + 1
    tData.ResPanMax = vVal(i)
    i = i + 1
    tData.ResPanMin = vVal(i)
    i = i + 1
    tData.ResRepTyp = vVal(i)
    i = i + 1
    tData.ResWrdPrn = vVal(i)
    i = i + 1
    tData.ResMchCod = vVal(i)
'    i = i + 1
'    tData.ResBtlCod = vVal(i)
'    i = i + 1
'    tData.ResSpmNum = vVal(i)
'    i = i + 1
'    tData.ResTryNum = vVal(i)
'    i = i + 1
'    tData.ResMicTyp = vVal(i)
'    i = i + 1
'    tData.ResLabNum = vVal(i)
'    i = i + 1
'    tData.ResGroYon = vVal(i)
'    i = i + 1
'    tData.ResGroDte = vVal(i)
    
    Exit Sub

ResInfLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub BacInfLoad(sRetVal As String, tData As BacInfRec)

    On Error GoTo BacInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sRetVal = "" Then Exit Sub

    vVal = Split(sRetVal & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tData.BacAcpDte = vVal(i)
    i = i + 1
    tData.BacAcpCod = vVal(i)
    i = i + 1
    tData.BacAcpNum = vVal(i)
    i = i + 1
    tData.BacSpmCod = vVal(i)
    i = i + 1
    tData.BacLabCod = vVal(i)
    i = i + 1
    tData.BacBacCod = vVal(i)
    i = i + 1
    tData.BacLabNum = vVal(i)
    i = i + 1
    tData.BacMicTyp = vVal(i)

    
    Exit Sub

BacInfLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub GroInfLoad(sRetVal As String, tData As GroInfRec)

    On Error GoTo GroInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sRetVal = "" Then Exit Sub

    vVal = Split(sRetVal & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tData.GroAcpDte = vVal(i)
    i = i + 1
    tData.GroAcpCod = vVal(i)
    i = i + 1
    tData.GroAcpNum = vVal(i)
    i = i + 1
    tData.GroSpmCod = vVal(i)
    i = i + 1
    tData.GroLabCod = vVal(i)
    i = i + 1
    tData.GroMicTyp = vVal(i)
    i = i + 1
    tData.GroGroYon = vVal(i)
    i = i + 1
    tData.GroRecCod = vVal(i)
    i = i + 1
    tData.GroLabNum = vVal(i)
    
    Exit Sub

GroInfLoad_ErrorTraping:
    Resume Next

End Sub


Public Sub StnInfLoad(sRetVal As String, tData As StnInfRec)

    On Error GoTo StnInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sRetVal = "" Then Exit Sub

    vVal = Split(sRetVal & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tData.StnAcpDte = vVal(i)
    i = i + 1
    tData.StnAcpCod = vVal(i)
    i = i + 1
    tData.StnAcpNum = vVal(i)
    i = i + 1
    tData.StnSpmCod = vVal(i)
    i = i + 1
    tData.StnLabCod = vVal(i)
    i = i + 1
    tData.StnStnCod = vVal(i)
    i = i + 1
    tData.StnLabNum = vVal(i)
    i = i + 1
    tData.StnMicTyp = vVal(i)

    
    Exit Sub

StnInfLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub MorInfLoad(sRetVal As String, tData As MorInfRec)

    On Error GoTo MorInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sRetVal = "" Then Exit Sub

    vVal = Split(sRetVal & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tData.MorAcpDte = vVal(i)
    i = i + 1
    tData.MorAcpCod = vVal(i)
    i = i + 1
    tData.MorAcpNum = vVal(i)
    i = i + 1
    tData.MorSpmCod = vVal(i)
    i = i + 1
    tData.MorLabCod = vVal(i)
    i = i + 1
    tData.MorColTyp = vVal(i)
    i = i + 1
    tData.MorSurTyp = vVal(i)
    i = i + 1
    tData.MorEdgTyp = vVal(i)
    i = i + 1
    tData.MorHemTyp = vVal(i)
    i = i + 1
    tData.MorExiTyp = vVal(i)
    i = i + 1
    tData.MorThiTyp = vVal(i)
    i = i + 1
    tData.MorLabNum = vVal(i)
    i = i + 1
    tData.MorMicTyp = vVal(i)
    Exit Sub

MorInfLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub AntInfLoad(sRetVal As String, tData As AntInfRec)

    On Error GoTo AntInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sRetVal = "" Then Exit Sub

    vVal = Split(sRetVal & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tData.AntAcpDte = vVal(i)
    i = i + 1
    tData.AntAcpCod = vVal(i)
    i = i + 1
    tData.AntAcpNum = vVal(i)
    i = i + 1
    tData.AntSpmCod = vVal(i)
    i = i + 1
    tData.AntLabCod = vVal(i)
    i = i + 1
    tData.AntBacCod = vVal(i)
    i = i + 1
    tData.AntBioTyp = vVal(i)
    i = i + 1
    tData.AntLabNum = vVal(i)
    i = i + 1
    tData.AntMicTyp = vVal(i)
    i = i + 1
    tData.AntMicRes = vVal(i)
    i = i + 1
    tData.AntMicDan = vVal(i)
    i = i + 1
    tData.AntRisRes = vVal(i)
    
    Exit Sub

AntInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ResInfStore(sKey As String, sRetVal As String, tData As ResInfRec)

    sKey = tData.ResAcpDte & Chr(5)
    sKey = sKey & tData.ResAcpCod & Chr(5)
    sKey = sKey & tData.ResAcpNum & Chr(5)
    sKey = sKey & tData.ResSpmCod & Chr(5)
    sKey = sKey & tData.ResLabCod & Chr(5)
    
    sRetVal = tData.ResOcmNum & Chr(5)
    sRetVal = sRetVal & tData.ResChtNum & Chr(5)
    sRetVal = sRetVal & tData.ResSotCod1 & Chr(5)
    sRetVal = sRetVal & tData.ResSotCod2 & Chr(5)
    sRetVal = sRetVal & tData.ResJbsSeq & Chr(5)
    sRetVal = sRetVal & tData.ResRltSeq & Chr(5)
    sRetVal = sRetVal & tData.ResMzhMin & Chr(5)
    sRetVal = sRetVal & tData.ResMzhMax & Chr(5)
    sRetVal = sRetVal & tData.ResMzhRef & Chr(5)
    sRetVal = sRetVal & tData.ResMzhUnt & Chr(5)
    sRetVal = sRetVal & tData.ResMzhMnt & Chr(5)
    sRetVal = sRetVal & tData.ResSplCmt & Chr(5)
    sRetVal = sRetVal & tData.ResOdrDtm & Chr(5)
    sRetVal = sRetVal & tData.ResAcpDtm & Chr(5)
    sRetVal = sRetVal & tData.ResTstDtm & Chr(5)
    sRetVal = sRetVal & tData.ResUpdDtm & Chr(5)
    sRetVal = sRetVal & tData.ResOdrUid & Chr(5)
    sRetVal = sRetVal & tData.ResAcpUid & Chr(5)
    sRetVal = sRetVal & tData.ResTstUid & Chr(5)
    sRetVal = sRetVal & tData.ResUpdUid & Chr(5)
    sRetVal = sRetVal & tData.ResSeeYon & Chr(5)
    sRetVal = sRetVal & tData.ResConLvl & Chr(5)
    sRetVal = sRetVal & tData.ResMzhTyp & Chr(5)
    sRetVal = sRetVal & tData.ResMzhLin & Chr(5)
    sRetVal = sRetVal & tData.ResShtNam & Chr(5)
    sRetVal = sRetVal & tData.ResSclCod & Chr(5)
    sRetVal = sRetVal & tData.ResPrtYon & Chr(5)
    sRetVal = sRetVal & tData.ResOspIsp & Chr(5)
    sRetVal = sRetVal & tData.ResMadYon & Chr(5)
    sRetVal = sRetVal & tData.ResJbsMth & Chr(5)
    sRetVal = sRetVal & tData.ResJbsQty & Chr(5)
    sRetVal = sRetVal & tData.ResCasYon & Chr(5)
    sRetVal = sRetVal & tData.ResEmgYon & Chr(5)
    sRetVal = sRetVal & tData.ResOdrNum & Chr(5)
    sRetVal = sRetVal & tData.ResOdrSeq & Chr(5)
    sRetVal = sRetVal & tData.ResOdrDep & Chr(5)
    sRetVal = sRetVal & tData.ResWrdCod & Chr(5)
    sRetVal = sRetVal & tData.ResStaYon & Chr(5)
    sRetVal = sRetVal & tData.ResOkYon & Chr(5)
    sRetVal = sRetVal & tData.ResOkUid & Chr(5)
    sRetVal = sRetVal & tData.ResOkDtm & Chr(5)
    sRetVal = sRetVal & tData.ResRedYon & Chr(5)
    sRetVal = sRetVal & tData.ResRedUid & Chr(5)
    sRetVal = sRetVal & tData.ResRedDtm & Chr(5)
    sRetVal = sRetVal & tData.ResPanMax & Chr(5)
    sRetVal = sRetVal & tData.ResPanMin & Chr(5)
    sRetVal = sRetVal & tData.ResRepTyp & Chr(5)
    sRetVal = sRetVal & tData.ResWrdPrn & Chr(5)
    sRetVal = sRetVal & tData.ResMchCod & Chr(5)
'    sRetVal = sRetVal & tData.ResBtlCod & Chr(5)
'    sRetVal = sRetVal & tData.ResSpmNum & Chr(5)
'    sRetVal = sRetVal & tData.ResTryNum & Chr(5)
'    sRetVal = sRetVal & tData.ResMicTyp & Chr(5)
'    sRetVal = sRetVal & tData.ResLabNum & Chr(5)
'    sRetVal = sRetVal & tData.ResGroYon & Chr(5)
'    sRetVal = sRetVal & tData.ResGroDte & Chr(5)
    
End Sub

Public Sub BacInfStore(sKey As String, sRetVal As String, tData As BacInfRec)

    sKey = tData.BacAcpDte & Chr(5)
    sKey = sKey & tData.BacAcpCod & Chr(5)
    sKey = sKey & tData.BacAcpNum & Chr(5)
    sKey = sKey & tData.BacSpmCod & Chr(5)
    sKey = sKey & tData.BacLabCod & Chr(5)
    sKey = sKey & tData.BacBacCod & Chr(5)
    
    sRetVal = tData.BacLabNum & Chr(5)
    sRetVal = sRetVal & tData.BacMicTyp & Chr(5)
    
End Sub


Public Sub GroInfStore(sKey As String, sRetVal As String, tData As GroInfRec)

    sKey = tData.GroAcpDte & Chr(5)
    sKey = sKey & tData.GroAcpCod & Chr(5)
    sKey = sKey & tData.GroAcpNum & Chr(5)
    sKey = sKey & tData.GroSpmCod & Chr(5)
    sKey = sKey & tData.GroLabCod & Chr(5)
        
    sRetVal = tData.GroMicTyp & Chr(5)
    sRetVal = sRetVal & tData.GroGroYon & Chr(5)
    sRetVal = sRetVal & tData.GroRecCod & Chr(5)
    sRetVal = sRetVal & tData.GroLabNum & Chr(5)
    
End Sub


Public Sub StnInfStore(sKey As String, sRetVal As String, tData As StnInfRec)

    sKey = tData.StnAcpDte & Chr(5)
    sKey = sKey & tData.StnAcpCod & Chr(5)
    sKey = sKey & tData.StnAcpNum & Chr(5)
    sKey = sKey & tData.StnSpmCod & Chr(5)
    sKey = sKey & tData.StnLabCod & Chr(5)
    sKey = sKey & tData.StnStnCod & Chr(5)
    
    sRetVal = tData.StnLabNum & Chr(5)
    sRetVal = sRetVal & tData.StnMicTyp & Chr(5)
    
End Sub

Public Sub MorInfStore(sKey As String, sRetVal As String, tData As MorInfRec)

    sKey = tData.MorAcpDte & Chr(5)
    sKey = sKey & tData.MorAcpCod & Chr(5)
    sKey = sKey & tData.MorAcpNum & Chr(5)
    sKey = sKey & tData.MorSpmCod & Chr(5)
    sKey = sKey & tData.MorLabCod & Chr(5)
    
    sRetVal = tData.MorColTyp & Chr(5)
    sRetVal = sRetVal & tData.MorSurTyp & Chr(5)
    sRetVal = sRetVal & tData.MorEdgTyp & Chr(5)
    sRetVal = sRetVal & tData.MorHemTyp & Chr(5)
    sRetVal = sRetVal & tData.MorExiTyp & Chr(5)
    sRetVal = sRetVal & tData.MorThiTyp & Chr(5)
    sRetVal = sRetVal & tData.MorLabNum & Chr(5)
    sRetVal = sRetVal & tData.MorMicTyp & Chr(5)
    
End Sub

Public Sub AntInfStore(sKey As String, sRetVal As String, tData As AntInfRec)

    sKey = tData.AntAcpDte & Chr(5)
    sKey = sKey & tData.AntAcpCod & Chr(5)
    sKey = sKey & tData.AntAcpNum & Chr(5)
    sKey = sKey & tData.AntSpmCod & Chr(5)
    sKey = sKey & tData.AntLabCod & Chr(5)
    sKey = sKey & tData.AntBacCod & Chr(5)
    sKey = sKey & tData.AntBioTyp & Chr(5)
    
    sRetVal = tData.AntLabNum & Chr(5)
    sRetVal = sRetVal & tData.AntMicTyp & Chr(5)
    sRetVal = sRetVal & tData.AntMicRes & Chr(5)
    sRetVal = sRetVal & tData.AntMicDan & Chr(5)
    sRetVal = sRetVal & tData.AntRisRes & Chr(5)
    
End Sub
    
    
    
Public Sub RsbInfLoad(sRetVal As String, tData As RsbInfRec)

    On Error GoTo RsbInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sRetVal = "" Then Exit Sub

    vVal = Split(sRetVal & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tData.RsbAcpDte = vVal(i)
    i = i + 1
    tData.RsbAcpCod = vVal(i)
    i = i + 1
    tData.RsbAcpNum = vVal(i)
    i = i + 1
    tData.RsbSpmCod = vVal(i)
    i = i + 1
    tData.RsbOcmNum = vVal(i)
    i = i + 1
    tData.RsbChtNum = vVal(i)
    i = i + 1
    tData.RsbItfYon = vVal(i)
    i = i + 1
    tData.RsbPrnYon = vVal(i)
    i = i + 1
    tData.RsbPrnUid = vVal(i)
    i = i + 1
    tData.RsbPrnDtm = vVal(i)
    i = i + 1
    tData.RsbAcpTim = vVal(i)
    i = i + 1
    tData.RsbSpcCmt = vVal(i)
    i = i + 1
    tData.RsbOkSw = vVal(i)
    i = i + 1
    tData.RsbOspIsp = vVal(i)
    i = i + 1
    tData.RsbTryNum = vVal(i)
    i = i + 1
    tData.RsbWrdCod = vVal(i)
    i = i + 1
    tData.RsbParTms = vVal(i)
    i = i + 1
    tData.RsbParDte = vVal(i)
    i = i + 1
    tData.RsbParTim = vVal(i)
    i = i + 1
    tData.RsbSpmNum = vVal(i)
    i = i + 1
    tData.RsbSpmDte = vVal(i)
    i = i + 1
    tData.RsbSpmTim = vVal(i)
    i = i + 1
    tData.RsbSpmUid = vVal(i)
    
    Exit Sub
    

RsbInfLoad_ErrorTraping:
    Resume Next



End Sub
    
Sub RsbInfRead(sPrmKey As String, RsbData As RsbInfRec)
    
    Dim sCurKey As String
    Dim sRetVal As String
    
    sCurKey = sPrmKey
    sCurKey = mSetReadEqual("RsbInf", sCurKey, sRetVal)
    Call RsbInfLoad(sRetVal, RsbData)
    
    Exit Sub

RsbInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub RsbInfStore(sKey As String, sRetVal As String, tData As RsbInfRec)

    
    sKey = tData.RsbAcpDte & Chr(5)
    sKey = sKey & tData.RsbAcpCod & Chr(5)
    sKey = sKey & tData.RsbAcpNum & Chr(5)
    sKey = sKey & tData.RsbSpmCod & Chr(5)
    
    sRetVal = tData.RsbOcmNum & Chr(5)
    sRetVal = sRetVal & tData.RsbChtNum & Chr(5)
    sRetVal = sRetVal & tData.RsbItfYon & Chr(5)
    sRetVal = sRetVal & tData.RsbPrnYon & Chr(5)
    sRetVal = sRetVal & tData.RsbPrnUid & Chr(5)
    sRetVal = sRetVal & tData.RsbPrnDtm & Chr(5)
    sRetVal = sRetVal & tData.RsbAcpTim & Chr(5)
    sRetVal = sRetVal & tData.RsbSpcCmt & Chr(5)
    sRetVal = sRetVal & tData.RsbOkSw & Chr(5)
    sRetVal = sRetVal & tData.RsbOspIsp & Chr(5)
    sRetVal = sRetVal & tData.RsbTryNum & Chr(5)
    sRetVal = sRetVal & tData.RsbWrdCod & Chr(5)
    sRetVal = sRetVal & tData.RsbParTms & Chr(5)
    sRetVal = sRetVal & tData.RsbParDte & Chr(5)
    sRetVal = sRetVal & tData.RsbParTim & Chr(5)
    sRetVal = sRetVal & tData.RsbSpmNum & Chr(5)
    sRetVal = sRetVal & tData.RsbSpmDte & Chr(5)
    sRetVal = sRetVal & tData.RsbSpmTim & Chr(5)
    sRetVal = sRetVal & tData.RsbSpmUid & Chr(5)
    
End Sub

    
Public Sub RstInfLoad(sPrmValue As String, tPrmData As RstInfRec)

    On Error GoTo RstInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.RstSotCod = vVal(i)
    i = i + 1
    tPrmData.RstSpmNum = vVal(i)
    i = i + 1
    tPrmData.RstSpmSeq = vVal(i)
    i = i + 1
    tPrmData.RstLabCod = vVal(i)
    i = i + 1
    tPrmData.RstSeq = vVal(i)
    i = i + 1
    tPrmData.RstSlpCod = vVal(i)
    i = i + 1
    tPrmData.RstSpmNam = vVal(i)
    i = i + 1
    tPrmData.RstOcmNum = vVal(i)
    i = i + 1
    tPrmData.RstAcpDte = vVal(i)
    i = i + 1
    tPrmData.RstSplCmt = vVal(i)
    i = i + 1
    tPrmData.RstMzhMax = vVal(i)
    i = i + 1
    tPrmData.RstMzhLow = vVal(i)
    i = i + 1
    tPrmData.RstMzhMnt = vVal(i)
    i = i + 1
    tPrmData.RstMzhUnt = vVal(i)
    i = i + 1
    tPrmData.RstJugCod = vVal(i)
    i = i + 1
    tPrmData.RstSlpDep = vVal(i)
    i = i + 1
    tPrmData.RstUidCod = vVal(i)
    i = i + 1
    tPrmData.RstUpdDtm = vVal(i)
    i = i + 1
    tPrmData.RstOdrCod = vVal(i)
    i = i + 1
    tPrmData.RstOdrNum = vVal(i)
    
    Exit Sub

RstInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub RstInfStore(sPrmKey As String, sPrmValue As String, tPrmData As RstInfRec)

    
    sPrmKey = tPrmData.RstSotCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.RstSpmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.RstSpmSeq), "@@") & Chr(5)
    
    sPrmValue = tPrmData.RstLabCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.RstSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstSlpCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstSpmNam & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.RstOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstAcpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstMzhMax & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstMzhLow & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstMzhMnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstMzhUnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstJugCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstSlpDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstOdrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RstOdrNum & Chr(5)
    
End Sub

    
Public Sub RsvInfLoad(sPrmValue As String, tPrmData As RsvInfRec)

    On Error GoTo RsvInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.RsvOcmNum = vVal(i)
    i = i + 1
    tPrmData.RsvDtm = vVal(i)
    i = i + 1
    tPrmData.RsvSts = vVal(i)
    i = i + 1
    tPrmData.RsvChtNum = vVal(i)
    i = i + 1
    tPrmData.RsvDepCod = vVal(i)
    i = i + 1
    tPrmData.RsvUidCod = vVal(i)
    i = i + 1
    tPrmData.RsvChkYon = vVal(i)
    i = i + 1
    tPrmData.RsvDtrCod = vVal(i)
    
    Exit Sub

RsvInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub RsvInfStore(sPrmKey As String, sPrmValue As String, tPrmData As RsvInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.RsvOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.RsvDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RsvSts & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RsvChtNum & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RsvDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RsvUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RsvChkYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RsvDtrCod & Chr(5)
    
End Sub

    
Public Sub SdlInfLoad(sPrmValue As String, tPrmData As SdlInfRec)

    On Error GoTo SdlInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.SdlEndDte = vVal(i)
    i = i + 1
    tPrmData.SdlOcmNum = vVal(i)
    i = i + 1
    tPrmData.SdlOcmSeq = vVal(i)
    i = i + 1
    tPrmData.SdlDupSeq = vVal(i)
    i = i + 1
    tPrmData.SdlDepCod = vVal(i)
    i = i + 1
    tPrmData.SdlIncCod = vVal(i)
    
    i = i + 1
    tPrmData.SdlChtNum = vVal(i)
    i = i + 1
    tPrmData.SdlInsCod = vVal(i)
    i = i + 1
    tPrmData.SdlInsSeq = vVal(i)
    i = i + 1
    tPrmData.SdlDtrCod = vVal(i)
    i = i + 1
    tPrmData.SdlInsAct = vVal(i)
    i = i + 1
    tPrmData.SdlInsMat = vVal(i)
    i = i + 1
    tPrmData.SdlNonAct = vVal(i)
    i = i + 1
    tPrmData.SdlNonMat = vVal(i)
    i = i + 1
    tPrmData.SdlInsAmt = vVal(i)
    i = i + 1
    tPrmData.SdlNonAmt = vVal(i)
    i = i + 1
    tPrmData.SdlInsOwn = vVal(i)
    i = i + 1
    tPrmData.SdlTotOwn = vVal(i)
    i = i + 1
    tPrmData.SdlSpcAmt = vVal(i)
    
    Exit Sub

SdlInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub SdlInfStore(sPrmKey As String, sPrmValue As String, tPrmData As SdlInfRec)

    
    sPrmKey = tPrmData.SdlEndDte & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.SdlOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.SdlOcmSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.SdlDupSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.SdlDepCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.SdlIncCod), "@@") & Chr(5)
    
    sPrmValue = Format(tPrmData.SdlChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.SdlInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlInsAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlInsMat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlNonAct & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlNonMat & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlNonAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlInsOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlTotOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SdlSpcAmt & Chr(5)
    
End Sub

    
Public Sub SrpInfLoad(sPrmValue As String, tPrmData As SrpInfRec)

    On Error GoTo SrpInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.SrpEndDte = vVal(i)
    i = i + 1
    tPrmData.SrpOcmNum = vVal(i)
    i = i + 1
    tPrmData.SrpOcmSeq = vVal(i)
    i = i + 1
    tPrmData.SrpDupSeq = vVal(i)
    i = i + 1
    tPrmData.SrpDepCod = vVal(i)
    i = i + 1
    tPrmData.SrpChtNum = vVal(i)
    i = i + 1
    tPrmData.SrpDtrCod = vVal(i)
    i = i + 1
    tPrmData.SrpInsCod = vVal(i)
    i = i + 1
    tPrmData.SrpInsSeq = vVal(i)
    i = i + 1
    tPrmData.SrpTotAmt = vVal(i)
    i = i + 1
    tPrmData.SrpCorAmt = vVal(i)
    i = i + 1
    tPrmData.SrpNonAmt = vVal(i)
    i = i + 1
    tPrmData.SrpOwnAmt = vVal(i)
    i = i + 1
    tPrmData.SrpTotOwn = vVal(i)
    i = i + 1
    tPrmData.SrpInsAmt = vVal(i)
    i = i + 1
    tPrmData.SrpSpcAmt = vVal(i)
    i = i + 1
    tPrmData.SrpAskAmt = vVal(i)
    i = i + 1
    tPrmData.SrpDisAmt = vVal(i)
    i = i + 1
    tPrmData.SrpFutAmt = vVal(i)
    i = i + 1
    tPrmData.SrpOldAmt = vVal(i)
    i = i + 1
    tPrmData.SrpNewAmt = vVal(i)
    i = i + 1
    tPrmData.SrpGrnAmt = vVal(i)
    i = i + 1
    tPrmData.SrpRcpNum = vVal(i)
    i = i + 1
    tPrmData.SrpOldNum = vVal(i)
    i = i + 1
    tPrmData.SrpCalDte = vVal(i)
    i = i + 1
    tPrmData.SrpUpdDtm = vVal(i)
    i = i + 1
    tPrmData.SrpUidCod = vVal(i)
    i = i + 1
    tPrmData.SrpDimAmt = vVal(i)
    i = i + 1
    tPrmData.SrpChgYon = vVal(i)
                
    Exit Sub

SrpInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub SrpInfStore(sPrmKey As String, sPrmValue As String, tPrmData As SrpInfRec)

    
    sPrmKey = tPrmData.SrpEndDte & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.SrpOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.SrpOcmSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.SrpDupSeq), "@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.SrpDepCod & Chr(5)
    
    sPrmValue = Format(tPrmData.SrpChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpInsCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.SrpInsSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpTotAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpCorAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpNonAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpOwnAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpTotOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpSpcAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpAskAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpDisAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpFutAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpOldAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpNewAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpGrnAmt & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.SrpRcpNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmData.SrpOldNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpCalDte & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpDimAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.SrpChgYon & Chr(5)

End Sub

    
Public Sub StbInfLoad(sPrmValue As String, tPrmData As StbInfRec)

    On Error GoTo StbInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.StbDepCod = vVal(i)
    i = i + 1
    tPrmData.StbAcpDtm = vVal(i)
    i = i + 1
    tPrmData.StbOcmNum = vVal(i)
    i = i + 1
    tPrmData.StbChtNum = vVal(i)
    i = i + 1
    tPrmData.StbPatNam = vVal(i)
    i = i + 1
    tPrmData.StbAcpStt = vVal(i)
    i = i + 1
    tPrmData.StbFlgStt = vVal(i)
    i = i + 1
    tPrmData.StbEmgYon = vVal(i)
    i = i + 1
    tPrmData.StbSplCmt = vVal(i)
    i = i + 1
    tPrmData.StbCstDep = vVal(i)
    i = i + 1
    tPrmData.StbDtrCod = vVal(i)
    i = i + 1
    tPrmData.StbXryFlg = vVal(i)
    i = i + 1
    tPrmData.StbOrgDtm = vVal(i)
    i = i + 1
    tPrmData.StbCfmStt = vVal(i)
    i = i + 1
    tPrmData.StbAtoBar = vVal(i)
    i = i + 1
    tPrmData.StbLabStt = vVal(i)    '13
    i = i + 1
    tPrmData.StbXryStt = vVal(i)    '14
    
    Exit Sub

StbInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub StbInfStore(sPrmKey As String, sPrmValue As String, tPrmData As StbInfRec)

    
    sPrmKey = tPrmData.StbDepCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.StbAcpDtm & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.StbOcmNum), "@@@@@@@@@@") & Chr(5)
    
    sPrmValue = Format((tPrmData.StbChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbAcpStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbFlgStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbEmgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbSplCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbCstDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbXryFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbOrgDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbCfmStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbAtoBar & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbLabStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.StbXryStt & Chr(5)
    
End Sub
    
Public Sub TcpInfLoad(sPrmValue As String, tPrmData As TcpInfRec)

On Error GoTo TcpInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
        
    i = i + 1
    tPrmData.TcpIp = vVal(i)
    
    i = i + 1
    tPrmData.TcpExeNam = vVal(i)
    
    i = i + 1
    tPrmData.TcpPath = vVal(i)
    
    i = i + 1
    tPrmData.TcpExeVer = vVal(i)
    
    i = i + 1
    tPrmData.TcpComNam = vVal(i)
    
    i = i + 1
    tPrmData.TcpAcpDtm = vVal(i)
    
'    i = i + 1
'    tPrmData.TcpUidCod = vVal(i)
    
    i = i + 1
    tPrmData.TcpPortNum = vVal(i)
    
    Exit Sub

TcpInfLoad_ErrorTraping:
    Resume Next

End Sub
        
Public Sub TcpInfStore(sPrmKey As String, sPrmValue As String, tPrmData As TcpInfRec)
   
    sPrmKey = tPrmData.TcpIp & Chr(5)
    sPrmKey = sPrmKey & tPrmData.TcpExeNam & Chr(5)
    
    sPrmValue = tPrmData.TcpPath & Chr(5)
    sPrmValue = sPrmValue & tPrmData.TcpExeVer & Chr(5)
    sPrmValue = sPrmValue & tPrmData.TcpComNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.TcpAcpDtm & Chr(5)
    'sPrmValue = sPrmValue & tPrmData.TcpUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.TcpPortNum & Chr(5)
    
End Sub

Public Sub ZfmInfLoad(sPrmValue As String, tPrmData As ZfmInfRec)

    On Error GoTo ZfmInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
        
    i = i + 1
    tPrmData.ZfmChtNum = vVal(i)
    i = 1 + 1
    i = i + 1
    tPrmData.ZfmAdpDte = vVal(i)
    i = i + 1
    tPrmData.ZfmInsAmt = vVal(i)
    i = i + 1
    tPrmData.ZfmCorAmt = vVal(i)
    i = i + 1
    tPrmData.ZfmOwnAmt = vVal(i)
    i = i + 1
    tPrmData.ZfmAskNew = vVal(i)
    
    Exit Sub

ZfmInfLoad_ErrorTraping:
    Resume Next

End Sub
    
    '-----------------------------------------
    '   정액,정률 Check를 위한 ZfmInf 읽기
    '-----------------------------------------
Public Sub ZfmInfRead(sChtNum As String, sAdpDte As String, ZfmData As ZfmInfRec)
    
    Dim sZfmInfCurKey As String, sZfmInfRetVal As String
    
    sZfmInfCurKey = sChtNum & Chr(5) & sAdpDte & Chr(5)
    
    sZfmInfCurKey = mSetReadEqual("ZfmInf", sZfmInfCurKey, sZfmInfRetVal)
    ZfmInfLoad sZfmInfRetVal, ZfmData
    
    Exit Sub

ZfmInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ZfmInfStore(sPrmKey As String, sPrmValue As String, tPrmData As ZfmInfRec)

    
    sPrmKey = Format(CDouble(tPrmData.ZfmChtNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.ZfmAdpDte & Chr(5)
    
    sPrmValue = tPrmData.ZfmInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ZfmCorAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ZfmOwnAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.ZfmAskNew & Chr(5)
    
End Sub


Public Sub BarInfLoad(sPrmValue As String, tPrmBarData As BarInfRec)

'    BarOcmNum As String     'Key 내원번호
'    BarOdrDte As String     'Key 적용일자
'    BarLabTst As String     'Key 검사학부
'
'    BarPrnDtm As String     '    출력일시
'    BarPrnCnt As String     '    출력매수
'    BarPrnUid As String     '    출력자 ID
    
    On Error GoTo BarInfLoad

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmBarData.BarOcmNum = vVal(i)
    i = i + 1
    tPrmBarData.BarOdrDte = vVal(i)
    i = i + 1
    tPrmBarData.BarLabTst = vVal(i)
    i = i + 1
    tPrmBarData.BarPrnDtm = vVal(i)
    i = i + 1
    tPrmBarData.BarPrnCnt = vVal(i)
    i = i + 1
    tPrmBarData.BarPrnUid = vVal(i)
    
    Exit Sub

BarInfLoad:
    Resume Next

End Sub
    
Public Sub BarInfStore(sPrmKey As String, sPrmValue As String, tPrmBarData As BarInfRec)
    
    sPrmKey = Format(CDouble(tPrmBarData.BarOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmBarData.BarOdrDte & Chr(5)
    sPrmKey = sPrmKey & tPrmBarData.BarLabTst & Chr(5)
    
    sPrmValue = tPrmBarData.BarPrnDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmBarData.BarPrnCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmBarData.BarPrnUid & Chr(5)
    
End Sub

Public Function BarInfRead(sPrmOcmNum As String, sPrmOdrDte As String, Optional sPrmLabTst As String) As BarInfRec

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String
    
    If sPrmLabTst = "" Then
        sCmpKey = sPrmOcmNum & Chr(5) & sPrmOdrDte & Chr(5)
    Else
        sCmpKey = sPrmOcmNum & Chr(5) & sPrmOdrDte & Chr(5) & sPrmLabTst & Chr(5)
    End If
    
    sCurKey = sCmpKey
    sCurKey = mSetNext("BarInf", sCurKey)
    sCurKey = mReadNext("BarInf", sCurKey, sCmpKey, sRetVal)
    
    Call BarInfLoad(sRetVal, BarInfRead)
    
End Function

Public Sub CdxInfLoad(sPrmValue As String, tPrmData As CdxInfRec)

    On Error Resume Next

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.CdxChtNum = vVal(i)
    i = i + 1
    tPrmData.CdxOdrDte = vVal(i)
    i = i + 1
    tPrmData.CdxFreNot = vVal(i)
    i = i + 1
    tPrmData.CdxCmtNot = vVal(i)
    i = i + 1
    tPrmData.CdxFreNt2 = vVal(i)

End Sub

Public Sub CdxInfStore(sPrmKey As String, sPrmValue As String, tPrmData As CdxInfRec)

    sPrmKey = Format((tPrmData.CdxChtNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CdxOdrDte & Chr(5)
    
    sPrmValue = tPrmData.CdxFreNot & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdxCmtNot & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdxFreNt2 & Chr(5)

End Sub

'예약어 등록
Public Sub CrvInfLoad(sPrmValue As String, tPrmData As CrvInfRec)

    On Error Resume Next

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1

    i = i + 1
    tPrmData.CrvMotCod = vVal(i)
    i = i + 1
    tPrmData.CrvChrCod = vVal(i)
    i = i + 1
    tPrmData.CrvCodNam = vVal(i)

End Sub

Public Sub CrvInfStore(sPrmKey As String, sPrmValue As String, tPrmData As CrvInfRec)

    sPrmKey = tPrmData.CrvMotCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CrvChrCod & Chr(5)

    sPrmValue = tPrmData.CrvCodNam & Chr(5)

End Sub

Public Sub CrvCtpInfLoad(sPrmValue As String, tPrmData As CrvCtpInfRec)

    On Error Resume Next

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then Exit Sub

    vVal = Split(sPrmValue & Replicate(Chr(5), 10), Chr(5))

    i = -1
    i = i + 1
    tPrmData.CrvCtpWrdNam = vVal(i)
    i = i + 1
    tPrmData.CrvCtpMotCod = vVal(i)
    i = i + 1
    tPrmData.CrvCtpChdCod = vVal(i)
    i = i + 1
    tPrmData.CrvCtpMotCodNam = vVal(i)
    i = i + 1
    tPrmData.CrvCtpChdCodNam = vVal(i)
    i = i + 1
    tPrmData.CrvCtpCodCon = vVal(i)

End Sub

Public Sub CrvCtpInfStore(sPrmKey As String, sPrmValue As String, tPrmData As CrvCtpInfRec)
    
    sPrmKey = tPrmData.CrvCtpWrdNam & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CrvCtpMotCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CrvCtpChdCod & Chr(5)

    
    sPrmValue = tPrmData.CrvCtpMotCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrvCtpChdCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CrvCtpCodCon & Chr(5)

End Sub
  
Public Sub CdeInfLoad(sPrmValue As String, tPrmData As CdeInfRec)

    On Error GoTo CdeInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.CdeChtNum = vVal(i)
    i = i + 1
    tPrmData.CdeOutDtm = vVal(i)
    i = i + 1
    tPrmData.CdeChtStt = vVal(i)
    i = i + 1
    tPrmData.CdeOutDep = vVal(i)
    i = i + 1
    tPrmData.CdeAcpDtm = vVal(i)
    i = i + 1
    tPrmData.CdeAskUid = vVal(i)
    i = i + 1
    tPrmData.CdeOutUid = vVal(i)
    i = i + 1
    tPrmData.CdeResDtm = vVal(i)
    i = i + 1
    tPrmData.CdeResUid = vVal(i)
    i = i + 1
    tPrmData.CdeDgsNfs = vVal(i)
    i = i + 1
    tPrmData.CdeFlgStt = vVal(i)
    i = i + 1
    tPrmData.CdePrnYon = vVal(i)
    i = i + 1
    tPrmData.CdeMemo = vVal(i)
    
    Exit Sub

CdeInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CdeInfStore(sPrmKey As String, sPrmValue As String, tPrmData As CdeInfRec)

    sPrmKey = Format(tPrmData.CdeChtNum, "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CdeOutDtm & Chr(5)
    
    sPrmValue = tPrmData.CdeChtStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeOutDep & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeAcpDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeAskUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeOutUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeResDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeResUid & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeDgsNfs & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeFlgStt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdePrnYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CdeMemo & Chr(5)
    
End Sub
    

Public Sub LbqInfRead(sPrmSotCod As String, sPrmAcpDte As String, sPrmOcmNum As String, LbqData As LbqInfRec)

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String
    
    sCurKey = sPrmSotCod & Chr(5) & sPrmAcpDte & Chr(5) & sPrmOcmNum & Chr(5)
    sCurKey = mSetReadEqual("LbqInf", sCurKey, sRetVal)
    Call LbqInfLoad(sRetVal, LbqData)

End Sub

Public Function IspInfRead(ByVal psOcmNum As String, ByVal psOdrNum As String, ByVal psOdrSeq As String, ptIspData As IspInfRec)
    
'    Dim scurKey As String
'    Dim sRetVal As String
    
    
'    scurKey = psOcmNum & Chr(5) & psOdrNum & Chr(5) & psOdrSeq & Chr(5)
'    scurKey = mSetReadEqual("IspInf", scurKey, sRetVal)
'    Call IspInfLoad(sRetVal, ptIspData)
    
    Dim sCurKey As String
    Dim sValue As String
    
    sCurKey = psOcmNum & Chr(5) & psOdrNum & Chr(5) & psOdrSeq & Chr(5)
    sCurKey = mSetReadEqual("IspInf", sCurKey, sValue)
    
    If (sCurKey <> "") Then
        IspInfRead = True
    Else
        IspInfRead = False
    End If
    
    Call IspInfLoad(sValue, ptIspData)
    
End Function

Public Sub PmdInfRead(ByVal psChtNum As String, ByVal psInsCod As String, ByVal psInsSeq As String, PmdData As PmdInfRec)

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String

    sCurKey = psChtNum & Chr(5) & psInsCod & Chr(5) & psInsSeq & Chr(5)
    sCurKey = mSetReadEqual("PmdInf", sCurKey, sRetVal)
    Call PmdInfLoad(sRetVal, PmdData)
    
End Sub

Public Sub PwkInfRead(ByVal psChtNum As String, ByVal psInsCod As String, ByVal psInsSeq As String, PwkData As PwkInfRec)

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String

    sCurKey = psChtNum & Chr(5) & psInsCod & Chr(5) & psInsSeq & Chr(5)
    sCurKey = mSetReadEqual("PwkInf", sCurKey, sRetVal)
    Call PwkInfLoad(sRetVal, PwkData)
    
End Sub

Public Sub PcrInfRead(ByVal psChtNum As String, ByVal psInsCod As String, ByVal psInsSeq As String, PcrData As PcrInfRec)

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String

    sCurKey = psChtNum & Chr(5) & psInsCod & Chr(5) & psInsSeq & Chr(5)
    sCurKey = mSetReadEqual("PcrInf", sCurKey, sRetVal)
    Call PcrInfLoad(sRetVal, PcrData)
    
End Sub


Public Sub NarInfLoad(sPrmValue As String, tPrmData As NarInfRec)

    On Error GoTo NarInfLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.NarOdrDte = vVal(i)
    i = i + 1
    tPrmData.NarMdcNum = vVal(i)
    i = i + 1
    tPrmData.NarOcmNum = vVal(i)
    i = i + 1
    tPrmData.NarOdrNum = vVal(i)
    i = i + 1
    tPrmData.NarOdrSeq = vVal(i)
    i = i + 1
    tPrmData.NarIOsw = vVal(i)
    i = i + 1
    tPrmData.NarInpCan = vVal(i)
    i = i + 1
    tPrmData.NarPrtDtm = vVal(i)
    i = i + 1
    tPrmData.NarPrtID = vVal(i)
    i = i + 1
    tPrmData.NarOutDtm = vVal(i)
    i = i + 1
    tPrmData.NarOutID = vVal(i)
    i = i + 1
    tPrmData.NarRcvID = vVal(i)
    i = i + 1
    tPrmData.NarInpPrt = vVal(i)
    i = i + 1
    tPrmData.NarEmgYon = vVal(i)
    i = i + 1
    tPrmData.NarInpDtm = vVal(i)
    Exit Sub

NarInfLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub NarInfStore(sPrmKey As String, sPrmValue As String, tPrmData As NarInfRec)
    
    sPrmKey = tPrmData.NarOdrDte & Chr(5)
    sPrmKey = sPrmKey & Format(Trim(tPrmData.NarMdcNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(Trim(tPrmData.NarOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(Trim(tPrmData.NarOdrNum), "@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(Trim(tPrmData.NarOdrSeq), "@@@@@") & Chr(5)
    
    sPrmValue = tPrmData.NarIOsw & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarInpCan & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarPrtDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarPrtID & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarOutDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarOutID & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarRcvID & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarInpPrt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarEmgYon & Chr(5)
    sPrmValue = sPrmValue & tPrmData.NarInpDtm & Chr(5)
    
End Sub

Public Function ReadNarInf(ByVal psKey As String, ptNarData As NarInfRec) As Boolean

    Dim sCurKey As String
    Dim sRetVal As String
    
    sCurKey = mSetReadEqual("NarInf", psKey, sRetVal)
    If sCurKey = "" Then
        ReadNarInf = False
    Else
        ReadNarInf = True
        Call NarInfLoad(sRetVal, ptNarData)
    End If
    
End Function

Public Sub OrvInfLoad(sPrmValue As String, ptOrvData As OrvInfRec)
    
    On Error GoTo OrvInfLoad

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 10)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 5), Chr(5))

    i = -1
    
    i = i + 1
    ptOrvData.OrvIcmNum = vVal(i)
    i = i + 1
    ptOrvData.OrvSeq = vVal(i)
    i = i + 1
    ptOrvData.OrvDepCod = vVal(i)
    i = i + 1
    ptOrvData.OrvDtrCod = vVal(i)
    i = i + 1
    ptOrvData.OrvRsvDte = vVal(i)
    i = i + 1
    ptOrvData.OrvRsvTim = vVal(i)
    i = i + 1
    ptOrvData.OrvRsvOcm = vVal(i)

    Exit Sub

OrvInfLoad:
    Resume Next
    
End Sub

Public Sub OrvInfStore(sPrmKey As String, sPrmValue As String, ptOrvData As OrvInfRec)

    sPrmKey = Format((ptOrvData.OrvIcmNum), "@@@@@@@@") & Chr(5)
    sPrmKey = sPrmKey & Format((ptOrvData.OrvSeq), "@@") & Chr(5)

    sPrmValue = ptOrvData.OrvDepCod & Chr(5)
    sPrmValue = sPrmValue & ptOrvData.OrvDtrCod & Chr(5)
    sPrmValue = sPrmValue & ptOrvData.OrvRsvDte & Chr(5)
    sPrmValue = sPrmValue & ptOrvData.OrvRsvTim & Chr(5)
    sPrmValue = sPrmValue & ptOrvData.OrvRsvOcm & Chr(5)

End Sub
