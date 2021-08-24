Attribute VB_Name = "MstDbRef"
 Option Explicit
    '******************************************************
    ' 마스터관련 Data Base Referance Field
    '******************************************************
    '------------------------------------------------------
    '29) 사용자 정의 코드 SeeMst              96/02/14
    '------------------------------------------------------
Type SeeMstRec
    SeeOdrCod  As String    'SeeMstKey 수가코드
    SeeSotCod  As String    '1        처방종목        96/03/04    Phr,Rad,Lab,Etc,Icd,Com
    SeeSlpDep  As String    '2        처방항목                    Lab,Rad,등등
    SeeSlpCod  As String    '3        처방종류        96/03/04    PHA,INJ,CHE,HEM,PCH,SCH,MRI
    SeeCodTyp  As String    '4        코드Type
    SeeEngNam  As String    '5        영문명칭
    SeeKorNam  As String    '6        한글명칭
    SeeElcCod  As String    '7        전산코드
    SeeItmCod  As String    '8        진료항목코드
    SeeAstCod  As String    '9        항목보조코드
    SeePhrTyp  As String    '10       약품종류        96/03/04    1'내복,2'외용,3'주사, 마약, 향정...
    SeeSlpTyp  As String    '11       처방전구분      96/03/04    1'일반,2'향정신,3'마약,4'독극
    SeeSclRat  As String    '12       수탁가산률
    SeeDivYon  As String    '13       분할사용여부
    SeeDrgCod  As String    '14       약품분류코드
    SeeUsgCod  As String    '15       용법/투여방법
    SeeMthCod  As String    '16       행위코드        96/03/04    행위코드,속도,간격,횟수
    SeeRepYon  As String    '17       함량처방 <--- 대체여부    99/03/07    함량단위 사용여부(Y)
    SeeAddCod  As String    '18       급여가산코드    95/10/25 명칭 수정, 순서수정
    SeeCalTyp  As String    '19       계산방법        95/10/25 신규 (문서에 없음)----->확인요함
    SeeUntQty  As String    '20       단위량          96/03/04
    SeeUntCod  As String    '21       단위            95/10/25 신규 (문서에 없음)----->확인요함
    SeeSpmCod  As String    '22       검체코드        96/03/04
    SeeMakCmp  As String    '23       제조회사        95/10/25 신규
    SeeSpcAmt  As String    '24       특진률/액
    SeeLftCnt  As String    '25       검사횟수
    SeeAdpDte  As String    '26       적용일
    SeeExpDte  As String    '27       종료일
    SeeUidCod  As String    '28       담당자코드
    SeeUpdDtm  As String    '29       수정일시
    SeeComNam  As String    '30       성분명칭
    SeeAddNon  As String    '31       비급여 가산 코드  97/04/17
    SeeCodDiv  As String    '32       코드구분
    SeeRelCod  As String    '33       관계코드
    SeeAdmCod  As String    '34       고가약관리코드
    SeeSotTyp  As String    '35       ocs 중분류
    SeeTotQty  As String    '36       1일 총투여량
    SeeTotTms  As String    '37       1일 총회수
    SeeEffect  As String    '38       효능분류  상세정보의 MEDEFF와 연결
    SeeInsAmt  As String    '39       보험단가
    SeeCarAmt  As String    '40       자보단가
    SeeWrkAmt  As String    '41       산재단가
    SeeCodYon  As String    '42       단독입력가능여부(Y이면 병동,외래OCS에서 직접 입력이 불가능함)
End Type
    
    '------------------------------------------------------
    '1-1) 시마터  History SeeHst
    '------------------------------------------------------
    
Type SeeHstRec
    SeeOdrCod  As String    'SeeMstKey 수가코드
    SeeAdpKey  As String    'SeeMstKey 적용일
    SeeSotCod  As String    '1        처방종목        96/03/04    Phr,Rad,Lab,Etc,Icd,Com
    SeeSlpDep  As String    '2        처방항목                    Lab,Rad,등등
    SeeSlpCod  As String    '3        처방종류        96/03/04    PHA,INJ,CHE,HEM,PCH,SCH,MRI
    SeeCodTyp  As String    '4        코드Type
    SeeEngNam  As String    '5        영문명칭
    SeeKorNam  As String    '6        한글명칭
    SeeElcCod  As String    '7        전산코드
    SeeItmCod  As String    '8        진료항목코드
    SeeAstCod  As String    '9        항목보조코드
    SeePhrTyp  As String    '10        약품종류        96/03/04    1'내복,2'외용,3'주사
    SeeSlpTyp  As String    '11       처방전구분      96/03/04    1'일반,2'향정신,3'마약,4'독극
    SeeSclRat  As String    '12       수탁가산률
    SeeDivYon  As String    '13       분할사용여부
    SeeDrgCod  As String    '14       약품분류코드
    SeeMthCod  As String    '15       용법/투여방법
    SeeUsgCod  As String    '16       행위코드        96/03/04    행위코드,속도,간격,횟수
    SeeRepYon  As String    '17       대체여부        95/10/25 순서 수정
    SeeAddCod  As String    '18       가산코드        95/10/25 명칭 수정, 순서수정
    SeeCalTyp  As String    '19       계산방법        95/10/25 신규 (문서에 없음)----->확인요함
    SeeUntQty  As String    '20       단위량          96/03/04
    SeeUntCod  As String    '21       단위            95/10/25 신규 (문서에 없음)----->확인요함
    SeeSpmCod  As String    '22       검체코드        96/03/04
    SeeMakCmp  As String    '23       제조회사        95/10/25 신규
    SeeSpcAmt  As String    '24       특진률/액
    SeeLftCnt  As String    '25       검사횟수
    SeeAdpDte  As String    '26       적용일
    SeeExpDte  As String    '27       종료일
    SeeUidCod  As String    '28       담당자코드
    SeeUpdDtm  As String    '29       수정일시
    SeeComNam  As String    '30       성분명칭
    SeeAddNon  As String    '31       비급여 가산 코드  97/04/17
    SeeCodDiv  As String    '32       코드구분
    SeeRelCod  As String    '33       관계코드
    SeeAdmCod  As String    '34       고가약관리코드
    SeeSotTyp  As String    '35       ocs 중분류
    SeeTotQty  As String    '36       1일 총투여량
    SeeTotTms  As String    '37       1일 총회수
    SeeEffect  As String    '38       효능분류  상세정보의 MEDEFF와 연결
    SeeInsAmt  As String    '39       보험단가
    SeeCarAmt  As String    '40       자보단가
    SeeWrkAmt  As String    '41       산재단가
    SeeCodYon  As String    '42       단독입력가능여부(Y이면 병동,외래OCS에서 직접 입력이 불가능함)
End Type
    
    '------------------------------------------------------
    '1-1) 수가코드 FeeHst
    '------------------------------------------------------
Type FeeMstRec
    FeeElcCod  As String    'FeeMstKey 전산코드
    FeeEngNam  As String    ' 1        영문명칭
    FeeKorNam  As String    ' 2        한글명칭
    FeeInsAmt  As String    ' 3        보험가
    FeeGenAmt  As String    ' 4        일반가
    FeeCarAmt  As String    ' 5        자보가
    FeeAdpDte  As String    ' 6        적용일
    FeeExpDte  As String    ' 7        종료일
    FeeUidCod  As String    ' 8        담당자코드
    FeeUpdDtm  As String    ' 9        수정일시
    FeeWrkAmt  As String    ' 10       산재가
    FeeGudAmt  As String    ' 11       보호가
    FeeLftAmt  As String    ' 12       신검가
    FeeInsAdp  As String    ' 13       보험적용금액
    FeeMakNam  As String    ' 14       제조회사
    FeeDrgCod  As String    ' 15       약품번호
    FeeUntCod  As String    ' 16       단위
    FeeCodDiv  As String    ' 17       수가구분
    FeeExtAmt  As String    ' 18       퇴장방지가
End Type
    
    '------------------------------------------------------
    '1-1) 수가코드 History FeeHst
    '------------------------------------------------------
Type FeeHstRec
    FeeElcCod  As String    'FeeMstKey 전산코드
    FeeAdpKey  As String    'FeeMstKey 적용일
    FeeEngNam  As String    ' 1        영문명칭
    FeeKorNam  As String    ' 2        한글명칭
    FeeInsAmt  As String    ' 3        보험가
    FeeGenAmt  As String    ' 4        일반가
    FeeCarAmt  As String    ' 5        자보가
    FeeAdpDte  As String    ' 6        적용일
    FeeExpDte  As String    ' 7        종료일
    FeeUidCod  As String    ' 8        담당자코드
    FeeUpdDtm  As String    ' 9        수정일시
    FeeWrkAmt  As String    ' 10       산재가
    FeeGudAmt  As String    ' 11       보호가
    FeeLftAmt  As String    ' 12       신검가
    FeeInsAdp  As String    ' 13       보험적용금액
    FeeMakNam  As String    ' 14       제조회사
    FeeDrgCod  As String    ' 15       약품번호
    FeeUntCod  As String    ' 16       단위
    FeeCodDiv  As String    ' 17       수가구분
    FeeExtAmt  As String    ' 18       퇴장방지가
End Type
    
    '------------------------------------------------------
    '2) 약속코드 마스터 GrpMst
    '------------------------------------------------------
Type GrpMstRec
    GrpCod      As String       '그룹코드 Key
    GrpOdrSeq   As String * 2   'Seq      Key
    GrpOdrCod   As String       '처방코드
    GrpOdrNam   As String       '처방명칭
    GrpAdpTyp   As String       '적용구분
    GrpOdrQty   As String       '투여량
    GrpOdrTms   As String       '횟수
    GrpOdrDay   As String       '일수
    GrpUsgCod   As String       '용법
    GrpMthCod   As String       '행위코드
    GrpSpmCod   As String       '검체코드
    GrpDgsYon   As String       '상병코드여부
    GrpInsYon   As String       '급여구분
    GrpSpcYon   As String       '특기여부
    GrpSpcCmt   As String       '특기사항
    GrpDgsRol   As String       '방사선촬영부위
    GrpItmCod   As String       '항목코드 (Group으로 묶여있는 코드는 자신의 항목코드보다
    GrpAstCod   As String       '항목보조코드     Group코드의 항목코드를 우선 적용한다.
    GrpSlpDep   As String       '처방전달 부서
    GrpDgsEtc   As String       '특기사항
    GrpAdpDte   As String       '적용일자
    GrpExpDte   As String       '종료일자
End Type
    
    '------------------------------------------------------
    '3) 항목코드 ItmMst
    '------------------------------------------------------
''Type ItmMstRec
''    ItmCod     As String        'ItmMstKey 항목코드
''    ItmAstCod  As String        'ItmMstKey 보조코드
''    ItmCodNam  As String        '          항목명칭
''    ItmWrkCod  As String        '          산재코드
''    ItmWrkYon  As String        '          산재급여구분
''    ItmCarCod  As String        '          자보코드
''    ItmCarYon  As String        '          자보급여구분
''    ItmIncCod  As String * 2    '          코드
''    ItmGudCod  As String        '          보호코드
''    ItmGudYon  As String        '          보호급여구분
''End Type
Type ItmMstRec
    ItmCod     As String        'ItmMstKey 항목코드
    ItmAstCod  As String        'ItmMstKey 보조코드
    ItmCodNam  As String        '          항목명칭
    ItmWrkCod  As String        '          산재코드
    ItmWrkYon  As String        '          산재급여구분
    ItmCarCod  As String        '          자보코드
    ItmCarYon  As String        '          자보급여구분
    ItmIncCod  As String * 2    '          수입원코드
    ItmGudCod  As String        '          보호코드
    ItmGudYon  As String        '          보호급여구분
    ItmAdpDte  As String        '          적용개시일
    ItmExpDte  As String        '          적용종료일
End Type



    '------------------------------------------------------
    '3) 항목코드 ItmMst
    '------------------------------------------------------
Type ItmHstRec
    ItmCod     As String        'ItmHstKey 항목코드
    ItmAstCod  As String        'ItmHstKey 보조코드
    ItmAdpKey  As String        'ItmHstKey 적용개시일
    ItmCodNam  As String        '          항목명칭
    ItmWrkCod  As String        '          산재코드
    ItmWrkYon  As String        '          산재급여구분
    ItmCarCod  As String        '          자보코드
    ItmCarYon  As String        '          자보급여구분
    ItmIncCod  As String * 2    '          수입원코드
    ItmGudCod  As String        '          보호코드
    ItmGudYon  As String        '          보호급여구분
    ItmAdpDte  As String        '          적용개시일
    ItmExpDte  As String        '          적용종료일
End Type

    
    '------------------------------------------------------
    '4) 상병코드 마스터 IcdMst
    '------------------------------------------------------
Type IcdMstRec
    IcdCod     As String    'IcdMstKey 상병코드
    IcdEngNam  As String    '1          영문상병
    IcdKorNam  As String    '2          한글상병
    IcdDepAra  As String    '3          진료분야
    IcdUpdDtm  As String    '4          수정일시
    IcdUidCod  As String    '5          담당자코드
    IcdLagCod  As String    '6          병명대분류
    IcdMidCod  As String    '7          병명중분류
    IcdHanIcd  As String    '8          한방/양방 연결코드
    IcdCanYon  As String    '9           암진단명        'yk : 퇴원분석때문에 추가...원래 병명대분류로 쓰이나...기존에 있던 값들의 용도를 몰라 새로추가한다.
'****************************************************> 추가
    IcdVeeCod  As String    '          V_코드   '20040115..HTS..
'****************************************************> 추가
End Type
    
    '------------------------------------------------------
    '5) 공휴코드 마스터 HolMst
    '------------------------------------------------------
Type HolMstRec
    HolDte     As String    'HolMstKey 일자
    HolDteNam  As String    '          명칭
End Type
    
    '------------------------------------------------------
    '6) 주소코드 마스터 ZipMst
    '------------------------------------------------------
Type ZipMstRec
    ZipCod     As String    'ZipMstKey 우편번호
    ZipLrgNam  As String    '          시,도명칭
    ZipMdlNam  As String    '          구,군명칭
    ZipSmlNam  As String    '          동,면명칭
    ZipLclAra  As String    '          지역권코드
End Type
    
    '------------------------------------------------------
    '7) 과목코드 마스터 DepMst
    '------------------------------------------------------
Type DepMstRec
    DepCod     As String        'DepMstKey 과목코드
    DepAdpDte  As String        'DepMstKey 적용일자
    DepKorNam  As String        '          한글명칭
    DepEngNam  As String        '          영문명칭
    DepGrpCod  As String        '          그룹총괄과
    DepBilCod  As String        '          청구코드
    DepBilAra  As String        '          청구분야
    DepBilSeq  As String * 2    '          명세서출력순서
    DepSndYon  As String        '          재진접수기 사용여부
    DepHspTyp  As String        '          병원구분(보험정보 1.의원, 2.병원, 3.종합병원, 4.대학병원)
    DepMdcTyp  As String        '          진료구분(보험정보 1.의과, 2.치과, 3.정신과, 4.한방과)
    DepMisPos  As String        '          재고부서
    DepIncTyp  As String        '          수입구분
    DepDgsCod As String         '          접수과목계
End Type
    
    
    '------------------------------------------------------
    '8) 소속 마스터 AssMst
    '------------------------------------------------------
Type AssMstRec
    AssCod     As String    'AssMstKey 소속코드
    AssInsCod  As String    '          보험코드
    AssCodNam  As String    '          소속명칭
    AssCtyTyp  As String    '          6대도시구분
    AssUpdDtm  As String    '          수정일시
    AssUidCod  As String    '          담당일시
    AssAddDtl  As String
    AssTelNum  As String
    AssFaxNum  As String
    AssEmlAdr  As String
End Type
    
    '------------------------------------------------------
    '9) 보험유형 InsMst
    '------------------------------------------------------
Type InsMstRec
    InsCod     As String    'InsMstKey 보험유형
    InsHspTyp  As String    'InsMstKey 병원구분 (1.의원, 2.병원, 3.종합병원, 4.대학병원)
    InsMdcTyp  As String    'InsMstKey 진료구분 (1.의과, 2.치과, 3.정신과, 4.한방과)
    InsAdpDte  As String    'InsMstKey 적용일자
    InsCodNam  As String    '          보험명칭
    InsConYon  As String    '          진찰료급비구분
    InsFeeYon  As String    '          진료비급비구분
    InsFeeLvl  As String    '          수가기준( 1:보험가 ,2:일반가 ,3:자보가 )
    InsOpoRat  As String    '          외래본인부담율
    InsOpbRat  As String    '          외래청구율
    InsIpoRat  As String    '          입원본인부담율
    InsIpbRat  As String    '          입원청구율
    InsHadRat  As String    '          병원가산율
    InsLmtHig  As String    '          정액 적용 상한액
    InsLmtOwn  As String    '          정액 본인일부 부담액
    InsLmt70   As String    '          정액 본인일부 부담액(70세이상)
    InsCasAmt  As String    '          포괄수가
    InsCodTyp  As String    '          내부코드( 11:일반 21:자보 31,32,33:보험 41:산재 51,52:보호)
    InsCutCod  As String    '          계약 유형(거래처코드 "G")
    InsNonYon  As String    '          비급여조합부담여부(Default="N", 조합청구면="Y"
    InsConCor  As String    '          진찰료조합부담여부(Default="N", 조합청구면="Y"
    InsDgsOpo  As String    '          진찰료 본인부담률
    InsDgsOpb  As String    '          진찰료 청구율
    InsLmtOut  As String    '          원외처방 포괄      <=추가
    InsLmtDig  As String    '          원외처방 정액      <=추가
    InsCasOut  As String    '          정액 치과 원외 발행 상한액   <=추가
    InsReqCod  As String    '          진료과와 청구과가 다를때 사용
End Type
    
    '------------------------------------------------------
    '10) 병동코드 WrdMst
    '------------------------------------------------------
Type WrdMstRec
    WrdCod     As String    'WrdMstKey 병동코드
    WrdCodNam  As String    '          병동명
    WrdAsgBed  As String    '          할당병상수
    WrdAprBed  As String    '          인가병상수
    WrdOcpBed  As String    '          점유병상수
    WrdMonDay  As String    '          당월재원일수
    WrdAnnDay  As String    '          금년재원일수
    WrdSnsDte  As String    '          최근통계보고일자
    WrdBasInf  As String    '          최근통계보고일자
End Type
    
    '------------------------------------------------------
    '11) 병실코드 RomMst
    '------------------------------------------------------
Type RomMstRec
    RomWrdCod  As String    'RomMstKey 병동코드     1
    RomCod     As String    'RomMstKey 병실코드     2
    RomCodNam  As String    ' 1        명칭         1
    RomDepCod  As String    ' 2        진료과       2
    RomBasBed  As String    ' 3        기준병상     3
    RomActBed  As String    ' 4        가동병상     4
    RomRemBed  As String    ' 5        잔여병상     5
    RomSexCod  As String    ' 6        성별         6
    RomTyp     As String    ' 7        구분         7
    RomGrdCod  As String    ' 8        병실등급     8
    RomStsCod  As String    ' 9        상태구분     9
    RomEqpInf  As String    '10        병실정보    10
End Type
    
    '------------------------------------------------------
    '12) 병상코드 BedMst
    '------------------------------------------------------
Type BedMstRec
    BedWrdCod  As String        'BedMstKey 병동코드
    BedRomCod  As String        'BedMstKey 병실코드
    BedCod     As String        'BedMstKey 병상코드
    BedSttCod  As String        ' 4        병상상태  O, V
    BedChtNum  As String * 8    ' 5        챠트번호
    BedPatNam  As String        ' 6        환자명
    BedPatSex  As String        ' 7        환자성별
    BedOcmNum  As String * 10   ' 8        내원변호
    BedIcdNam  As String        ' 9        진단병명
    BedDepCod  As String        ' 10       진료과목
    BedPatSts  As String        ' 11       환자상태
    BedTrnDtm  As String        ' 12       이송일시
    BedCsnDtm  As String        ' 13       승인일시
    BedCsnTyp  As String        ' 14       승인형태
    BedBirDay  As String        ' 15       생년월일
    BedLevTyp  As String        ' 16       퇴원코드
    BedDtrCod  As String        ' 17       담당의사
    BedAcuCod  As String        ' 18       응급코드
    BedIntTel  As String        ' 19       병실내선번호 및 전화번호
End Type
    
    '------------------------------------------------------
    '14) 메시지 정보 IcdMst
    '------------------------------------------------------
Type MsgMstRec
    MsgCod     As String    'IcdMstKey 메시지코드
    MsgCodNam  As String    '          메시지명칭
End Type
    
    '------------------------------------------------------
    '15) 수입원 코드 IncMst
    '------------------------------------------------------
Type IncMstRec
    IncCod     As String * 2    'IncMstKey 수입원코드
    IncCodNam  As String        '          수입원명칭
    IncIprSit  As String        '          입원영수증 순서
    IncOprSit  As String        '          외래영수증 순서
    IncEtcDep  As String        '          기타수입과목
    IncInoSit  As String        '          입원영수증 순서
    IncOnoSit  As String        '          외래영수증 순서
    IncTyp  As String           '          수입구분
    IncHOpSit  As String        '          입원영수증 순서
    IncHOnSit  As String        '          입원영수증 순서
    IncHIpSit  As String        '          입원영수증 순서
    IncHInSit  As String        '          입원영수증 순서
    
End Type
    
    '----------------------------------------------
    '16) 최종 정보 FnlMst
    '----------------------------------------------
Type FnlMstRec
    FnlCod     As String    'FnlMstKey 구분 코드
    FnlNum     As String    '          최종번호
    FnlDte     As String    '          최종일자
End Type
    
    '----------------------------------------------
    '17) 상세 도움말 정보 DtlMst
    '----------------------------------------------
Type DtlMstRec
    DtlTblCod  As String    'DtlMstKey 테이블 코드
    DtlCod     As String    'DtlMstKey 상세코드
    DtlCodNam  As String    '          상세코드명칭
End Type
    
    '----------------------------------------------
    '18) 상세 도움말 테이블 TabMst
    '----------------------------------------------
Type TabMstRec
    TabCod     As String    'TabMstKey 테이블 코드
    TabCodNam  As String    '          테이블 코드명칭
    TabUpdYon  As String    '          테이블 코드 수정 여부
End Type
    
    '------------------------------------------------------
    'XXXXX 과목코드 마스터 DepMst
    '------------------------------------------------------
Type DepMstGrpRec
    DepGrpCod  As String    'DepMstKey 그룹총괄과
    DepCod     As String    '          과목코드
End Type
    
    '------------------------------------------------------
    '20) 진료비계산 마스터 FtdMst
    '------------------------------------------------------
Type FtdMstRec
    FtdDgsCod  As String    'FtdMstKey  전문과목별 계
    FtdDgsNfs  As String    'FtdMstKey  초재진 구분
    FtdDgsDnh  As String    'FtdMstKey  주야공휴 구분
    FtdAgeDiv  As String    'FtdMstKey  나이 여부
    FtdAdpDte  As String    'FtdMstKey  적용기준일
    FtdCodNam  As String    '           진찰료코드 명칭
    FtdFeeCod  As String    '           수가코드
    FtdRsuAmt  As String    '           진찰권 재발행료
    FtdSpcAmt  As String    '           특진료
End Type
    
    '------------------------------------------------------
    '21) 가산코드 마스터 AddMst         97/08/26 신규...
    '------------------------------------------------------
Type AddMstRec
    AddCod     As String    'AddMstKey  가산 대분류 코드
    AddAdpDte  As String    'AddMstKey  적용일자
    AddCodNam  As String    '           명칭
    AddRatOne  As String    '           가산 I
    AddRatTwo  As String    '           가산 II
    AddRatThr  As String    '           가산 III
    AddRatTot  As String    '           총 가산
End Type
    
    '------------------------------------------------------
    '22) 가산코드 마스터 AddMst         97/08/26 신규...
    '------------------------------------------------------
Type CalMstRec
    CalAddCod  As String    'CalMstKey  가산 대분류 코드
    CalAdpDte  As String    'CalMstKey  적용일자
    CalCodNam  As String    '           명칭
    CalAddRat  As String    '           가산률
    CalAddPrc  As String    '           가산액
    CalAddAmt  As String    '           가산정액
    CalFeeCod  As String    '           적용코드
    CalRatOne  As String    '           가산 I
    CalFeeOne  As String    '           적용코드 I
    CalRatTwo  As String    '           가산 II
    CalFeeTwo  As String    '           적용코드 II
    CalRatThr  As String    '           가산 III
    CalFeeThr  As String    '           적용코드 III
End Type
    
    '------------------------------------------------------
    'XXXXX 처리구분 EtcMst                      95/10/27 신규
    '------------------------------------------------------
Type EtcMstRec
    EtcItmCod  As String    'EtcMstKey 처리구분 분류 코드
    EtcCod     As String    'EtcMstKey 구분코드
    EtcCodNam  As String    '          명칭
End Type
    
    '------------------------------------------------------
    '23) 계정 마스터 AccMst
    '------------------------------------------------------
Type AccMstRec
    AccCod     As String    'AccMstKey 계정 코드
    AccCodNam  As String    ' 1        계정 명칭
    AccClsCod  As String    ' 2        계정 구분
    AccEmpYon  As String    ' 3        직원 관련 여부
    AccFncYon  As String    ' 4        적용 업무 구분
    AccAmtYon  As String    ' 5        적용 금액 구분
    AccFdgRat  As String    ' 6        접수 초진 계정률
    AccSdgRat  As String    ' 7        접수 재진 계정률
    AccCalRat  As String    ' 8        수납 계정률
    AccIncDiv  As String    ' 9        적용 수입원 코드
    AccInsYon  As String    ' 10       조합 청구 여부
    AccAssTyp  As String    ' 11       청구유형
    AccGbnTyp  As String    ' 12       청구구분
    AccConYon  As String    ' 13       계속 여부
    '-------------------'
    '- 상세할인 작업중 -'              각 수입원 코드는 ^ 를 구분자로 저장한다.
    AccInsMat  As String    ' 15       보험재료 금액을 할인에 적용하는 수입원 코드
    AccInsAct  As String    ' 16       보험행위 금액을 할인에 적용하는 수입원 코드
    AccNonMat  As String    ' 17       비급여재료 금액을 할인에 적용하는 수입원 코드
    AccNonAct  As String    ' 18       비급여행위 금액을 할인에 적용하는 수입원 코드
    AccSpcAmt  As String    ' 19       특진료 금액을 할인에 적용하는 수입원 코드
    '-------------------'
    AccShwYon  As String    ' 20       조회시 Display여부 (대구성삼추가)
End Type
    
    '------------------------------------------------------
    '24) 병원 기본정보 HspMst
    '------------------------------------------------------
Type HspMstRec
    HspCod     As String    'HspMstKey 병원코드
    HspNam     As String    ' 2        병원명
    HspInsNum  As String    ' 3        요양기관지정번호
    HspInsNam  As String    ' 4        기관명칭
    HspWrkNum  As String    ' 5        산재지정번호
    HspRgnNum  As String    ' 6        사업자등록번호
    HspHspAdr  As String    ' 7        사업장소재지
    HspRcpNam  As String    ' 8        상호
    HspMdcYon  As String    ' 9        진료기관구분
    HspLmtYon  As String    ' 10       정액정률구분
    HspCloTim  As String    ' 11       통계마감시간
    HspOwnNam  As String    ' 12       대표자성명
    HspZipCod  As String    ' 13       지역우편번호
    HspManNam  As String    ' 14       청구서 작성자 성명
    HspManRes  As String    ' 15       청구서 작성자 주민번호
    HspTelNum  As String    ' 16       전화번호
    HspBilSlp  As String    ' 17       처방전 매수
    HspBilDte  As String    ' 18       청구 년/월/일
    HspBilMan  As String    ' 19       청구인
    HspBilCnt  As String    ' 20       명세서 매수
    HspAdmMth  As String    ' 21       접수 선,후 수납
    HspAdmRcp  As String    ' 22       외래접수 영수증
    HspOcmRcp  As String    ' 23       외래수납 영수증
    HspMidRcp  As String    ' 24       중간계산서 매수      '0 이면 출력 안함
    HspMrpRcp  As String    ' 25       중간수납영수증 매수
    HspDisRcp  As String    ' 26       퇴원계산서 매수
    HspIcmRcp  As String    ' 27       퇴원수납 영수증  매수
    HspPreRcp  As String    ' 28       선수납영수증
    HspGrnRcp  As String    ' 29       보증금영수증
    HspStaDgs  As String    ' 30       마감구분 (DtlMst의 "STATBL" 1:24시마감,2:회계일자마감)
    HspLgoPth  As String    ' 31       로고패스 (ex:\\Asp\Hnt.Cnv\Icon\로고.Bmp)
    HspRsvFee  As String    ' 32       예약진료비 선수납(Y) /후수납(N)
End Type
    
    '------------------------------------------------------
    '25) 등급 코드 GrdMst                     96/01/20 신규
    '------------------------------------------------------
Type GrdMstRec
    grdCod     As String    'GrdMstKey 병실등급
    GrdAdpDte  As String    'GrdMstKey 적용기준일
    GrdNam     As String    '          명칭
    GrdFeeCod  As String    '          병실료코드
    GrdMedCod  As String    '          내,정,소 코드
    GrdIsoCod  As String    '          격리병실
    GrdAmtCod  As String    '          병실차액코드
    GrdIcuCod  As String    '          중환자실 가산코드
    GrdRomGrd  As String    '          병실유형(G:일반병실, I:ICU, K:격리병실..)
End Type
    
    '------------------------------------------------------
    '26) 행위 코드 MthMst                     96/01/23 신규
    '------------------------------------------------------
Type MthMstRec
    MthCod     As String    'MthMstKey 행위코드
    MthAdpDte  As String    'MthMstKey 적용기준일
    MthNam     As String    '          명칭
    MthFeeCod  As String    '          수가코드
    MthInpLmt  As String    '          입원한계치
    MthOutLmt  As String    '          외래한계치
    MthInpOvr  As String    '          입원한계치초과분처리 flag("":무시 "N":비급여처리)
    MthOutOvr  As String    '          외래한계치초과분처리 flag("":무시 "N":비급여처리)
    '수기료수정 @:^)
    MthCinLmt  As String    '          자보-입원한계치
    MthCouLmt  As String    '          자보-외래한계치
    MthCinOvr  As String    '          자보-입원한계치초과분처리 flag("":무시 "N":비급여처리)
    MthCouOvr  As String    '          자보-외래한계치초과분처리 flag("":무시 "N":비급여처리)
    MthSinLmt  As String    '          산재-입원한계치
    MthSouLmt  As String    '          산재-외래한계치
    MthSinOvr  As String    '          산재-입원한계치초과분처리 flag("":무시 "N":비급여처리)
    MthSouOvr  As String    '          산재-외래한계치초과분처리 flag("":무시 "N":비급여처리)
    MthBinLmt  As String    '          보호-입원한계치
    MthBouLmt  As String    '          보호-외래한계치
    MthBinOvr  As String    '          보호-입원한계치초과분처리 flag("":무시 "N":비급여처리)
    MthBouOvr  As String    '          보호-외래한계치초과분처리 flag("":무시 "N":비급여처리)
    '/수기료수정 @:^)
End Type
    
    '------------------------------------------------------
    '27) 리포트 프로그램 RptMst               96/01/24 신규
    '------------------------------------------------------
Type RptMstRec
    RptJobTyp  As String        'RptMstKey 리포트 구분 (외래, 입원, 발생)
    RptSeqNum  As String * 3    'RptMstKey 순서
    RptTypFlg  As String        '          Report, Statatistics
    RptNam     As String        '          명칭
    RptExeNam  As String        '          실행화일명
    RptSgnCnt  As String        '          결제란수
    RptSgnNam  As String        '          결제란 ','로 구분
End Type
    
    '------------------------------------------------------
    '28) 용법 코드 UsgMst                     96/02/12 신규
    '------------------------------------------------------
Type UsgMstRec
    UsgCod     As String        'UsgMstKey  용법 Code
    UsgFulDsc  As String        '           영문 명칭
    UsgCodNam  As String        '           용법 명칭
    UsgOdrTms  As String        '           횟수
    UsgMthCod  As String        '           행위코드
    UsgDspSeq  As String * 4    '           화면 Display 순서
    UsgDspGrp  As String * 2    '           화면 Display 그룹
    UsgActTim  As String        '           Default Acting Time
    UsgMainYon As String       '횟수에 대한 디폴트 용법을 표시함. Y/N 20030214 lek edit
End Type
    

'---------------------------------------------------------------------------------------
' new 슬립코드 마스터   DrsMst(Doctor's Routine Slip Master              2003/02/12
'---------------------------------------------------------------------------------------
Type DrsMstRec

    DrsSotTyp    As String      'DrsMstKey  슬립구분
    DrsSotCod    As String      'DrsMstKey  슬립종목(항목)
    DrsSitCod    As String      'DrsMstKey  작성부서
    DrsDtrCod    As String      'DrsMstKey  의사코드
    DrsSlpCod    As String * 5  'DrsMstKey  슬립종류(범주)
    DrsOdrSeq    As String * 5    'DrsMstKey  처방 Seq
    
    DrsOdrCod    As String      '1          처방 코드
    DrsCodNam    As String      '2          처방명
    DrsOdrQty    As String      '3          투여량
    DrsOdrTms    As String      '4          횟수
    DrsOdrDay    As String      '5          일수
    DrsUsgCod    As String      '6          용법
    DrsSpmCod    As String      '7          검체코드
    DrsSpcYon    As String      '8          특기여부
    DrsSpcCmt    As String      '9          특기사항
    DrsDgsRol    As String      '10         방사선촬영부위(Left, Right)
    DrsAdpTyp    As String      '11         적용구분        ----> 96/05/07 추가 (GrpMst에 있는 Field)
    DrsMthCod    As String      '12         행위구분        ----> 96/05/07 추가 (GrpMst에 있는 Field)
    DrsDgsYon    As String      '13         상병코드여부    ----> 96/05/07 추가 (GrpMst에 있는 Field)
    DrsInsYon    As String      '14         급여구분        ----> 96/05/07 추가 (GrpMst에 있는 Field)
    DrsSlpDep    As String      '15         가야할곳
    DrsDgsEtc    As String      '16         추가사항
    DrsSlpSeq    As String * 3  'DrsMstKey  슬립순서
    DrsRepYon    As String      '           의사사용량
    DrsItmCod    As String      '19         항목정보 ====> 2001/11/30 james 추가 (SeeMst의 SeeItmCod)
    
End Type

    '------------------------------------------------------
    '33) 특기사항 정보 CmtMst
    '------------------------------------------------------
Type CmtMstRec
    CmtUidCod  As String    'CmtMstKey 담당자 코드
    CmtUsgPgm  As String    'CmtMstKey 사용 program ( "SPC" : Special Comment, "CHT" : Chart, "MAL" : Mail )
    CmtCod     As String    'CmtMstKey 특기사항코드
    CmtCodNam  As String    '          특기사항명칭
End Type

    '------------------------------------------------------
    '34) 검사정보 LabMst
    '------------------------------------------------------
Type LabMstRec
    LabCod     As String        'LabMstKey  검사코드   * 8                  1
    LabSeq     As String * 2    'LabMstKey  순서 0...Single 1...n           2
    LabSubCod  As String        '  1        = LabCod                        3
    LabCodTyp  As String        '  2        I : Indivisual, S : SubGroup    4
    LabCodNam  As String        '  3        명칭                            5
    LabTubCod  As String        '  4        용기코드   * 3                  6
    LabSpmCod  As String        '  5        검체코드   * 3                  7
    LabSpmReq  As String        '  6        검체필요량 * 3                  8
    LabComMax  As String        '  7        공통상한치 * 5.2                9
    LabComLow  As String        '  8        공통하한치 * 5.2               10
    LabMalMax  As String        '  9        남성상한치 * 5.2               11
    LabMalLow  As String        ' 10        남성하한치 * 5.2               12
    LabFmlMax  As String        ' 11        여성상한치 * 5.2               13
    LabFmlLow  As String        ' 12        여성하한치 * 5.2               14
    LabMzhUnt  As String        ' 13        결과단위   * 5                 15
    LabSclYon  As String        ' 14        수탁여부   * 1                 16
End Type
    
    
    '------------------------------------------------------
    '35) 식대등급정보   MgdMst              96/09/17 신규
    '------------------------------------------------------
Type MgdMstRec
    MgdCod      As String       'MgdMstKey  식대코드
    MgdAdpDte   As String       'MgdMstKey  적용일시
    MgdNam      As String       '1           코드명
    MgdInsCod   As String       '2           보험코드
    MgdNatCod   As String       '3           보호코드
    MgdCarCod   As String       '4           자보코드
    MgdWrkCod   As String       '5           산재코드
    MgdGenCod   As String       '6           일반코드
    MgdDifCod   As String       '7           차액코드
    MgdMgdSeq   As String * 3     '8           식대순서
    MgdMgdPic   As String       '9             도형모양
    MgdCruCod   As String       '10            직원식대코드
    '3월1일보호1종식이변경처리
    MgdSecCod   As String       '11          순수보호1종 전용코드
    
End Type
    
    '------------------------------------------------------
    '35) 병동관리정보 WmnMst
    '------------------------------------------------------
Type wmnMstRec
    WmnWrdCod As String         'WmnMstKey  병동코드
    WmnSexTyp As String         '           남녀구분
    WmnManTot As String         '           남자환자수
    WmnDepTyp As String         '           과별구분
    WmnWrdTyp As String         '           병동구분
End Type
    
    '------------------------------------------------------
    '36) 계약처 정보   CutMst              97/02/24 신규
    '------------------------------------------------------
Type CutMstRec
    CutGub      As String       'CutMstKey  거래구분코드(계약공상 "G", 검사수탁 "S")
    CutCod      As String       'CutMstKey  거래처코드
    CutAdpDte   As String       'CutMstKey  거래 시작일
    CutExpDte   As String       '           거래 종료일
    CutNam      As String       '           거래처명
    CutInsMat   As String       '           보험적용 재료가산률
    CutInsAct   As String       '           보험적용 행위가산률
    CutNonMat   As String       '           비급여 재료가산률
    CutNonAct   As String       '           비급여 행위가산률
    CutStgGen   As String       '           일반 수탁 가산률
    CutStgCar   As String       '           자보 수탁 가산률
    CutStgIns   As String       '           보험 수탁 가산률
    CutStgWrk   As String       '           산재 수탁 가산률
    CutStgBoh   As String       '           보호 수탁 가산률
    CutUpdDtm   As String       '           수정일시
    CutUidCod   As String       '           수정담당
    CutNum      As String       '           거래정보
    
End Type
    
    '-------------------------------------------------------
    ' 공급실 관리 품목 분류
    '-------------------------------------------------------
Type CsrMstRec
    CsrDepTyp   As String   'Key    입력부서(공급실:CSR, 병동:WRD, ...)
    CsrCod      As String   'Key    분류코드(Key)
    CsrCsmYon   As String   'Key    소모품여부
    CsrCodNam   As String   '       코드명칭
    CsrOmsNam   As String   '       약자
    CsrUntCod   As String   '       단위코드
    CsrUntQty   As String   '       단위용량
    CsrSeeCod   As String   '       수가코드(970522추가)
End Type
    
    '-------------------------------------------------------
    ' 공급실 기초 목록 순서
    '-------------------------------------------------------
Type CspMstRec
    CspUsgPrt   As String      'Key    사용부서(공급실:CSR, 병동:WRD, ...)
    CspSeq      As String * 3  'Key    Display 순서
    CspCsrCod   As String      '       공급실 품목
    CspDepTyp   As String      '       Csr의 Key 입력부서(공급실:CSR, 병동:WRD, ...)
End Type
    
    '------------------------------------------------------
    '40) 수술 코드 정보       OprMst      97.1.27
    'Index                    OprMstOpr   K-2
    '------------------------------------------------------
Type OprMstRec
    OprDepCod       As String   'OprMstKey  과목코드
    OprCod          As String   'OprMstKey  수술코드
    OprCodDsc       As String   '           코드설명
    OprSeeCod       As String   '           수가코드
    OprUseTms       As String   '           수술시간
    OprIcdYon       As String   '           진단명 여부("Y" or "N")
End Type
    
    '------------------------------------------------------
    '40-1) 수술상병코드 정보       OpiMst     971023
    '------------------------------------------------------
Type OpiMstRec
    OpiDepCod       As String   'OpiMstKey  과목코드
    OpiCod          As String   'OpiMstKey  수술코드
    OpiCodDsc       As String   '           코드설명
End Type
    
    '----------------------------------------------------------------
    '41) 실행파일에 대한 권한부여 정보          ExeMst      97.5.14
    '-----------------------------------------------------------------
Type ExeMstRec
    ExeCod          As String               'ExeMstKey  개인 ID
    ExeMainNam      As String               'ExeMstKey 주 메뉴이름
    'ExeSubIdx       As String               'ExeMstKey  부메뉴의 Index
    ExeExeNam       As String               'ExeMstKey  실행화일 명
    ExeFlg          As String               '"Y" & "N"
End Type
    
    '----------------------------------------------------------------
    '42)물품 Group묶음정보    SCGMST     970515
    '-----------------------------------------------------------------
Type CsgMstRec
    CsgCod As String
    CsgCsrCod As String
    
    CsgNam As String
    CsgCsrUnt As String
    CsgCsrQty As String
    CsgUseYon As String
End Type
    
    '----------------------------------------------------------------
    '43)물품 Set묶음정보    SCSMST       970515
    '-----------------------------------------------------------------
Type CssMstRec
    CssSetCod As String
    CssCsrCod As String
    CssSetNam As String
    CssCsrCnt As String
    CssUseYon As String
    CssCsrTyp As String
End Type
    
    '----------------------------------------------------------------
    '44) Drg정보
    '-----------------------------------------------------------------
Type DrgMstRec
    DrgCod    As String                 'Key DRG Code
    DrgOdrDay As String * 2             'Key 적용일수
    DrgAdpDte As String                 'Key 적용개시일
    DrgCorAmt As String                 '조합부담
    DrgAskAmt As String                 '본인부담
End Type
    
    '----------------------------------------------------------------
    '45) 방사선 특수 촬영저보(Special Xray)
    '-----------------------------------------------------------------
Type SxyMstRec
    SxyElcCod As String                 'Key 방사선 코드
    SxyOdrSeq As String * 2             'Key 일련번호
    SxyOdrCod As String                 '적용코드
    SxyOdrQty As String                 '적용수량
End Type
    
    '----------------------------------------------------------------
    '46) 수입마감정보
    '-----------------------------------------------------------------
Type ImgMstRec
    ImgElcCod As String                 'P-K    대상코드 ++
    ImgAdpDte As String                 'P-K    적용시작일자 ++
    ImgExpDte As String                 'D-1    작용종료일자 ++
    ImgAdpCod As String                 'D-2    적용코드
    ImgAdpPrc As String                 'D-3    적용금액 ++
    ImgPatTyp As String                 'D-4    환자유형
    ImgDepCod As String                 'D-5    과목정보
    ImgInsCod As String                 'D-6    보험유형
End Type
    
    '----------------------------------------------------------------
    '47) 심사지침정보
    '-----------------------------------------------------------------
Type SimMstRec
    SimOdrCod  As String    'SimMstKey  수가코드
    SimRepCod  As String    '1          대체가능코드
    SimLowQty  As String    '2          1일 최소허용량
    SimHigQty  As String    '2          1일 최대허용량
    SimAvgQty  As String    '2          1일 표준용량
    SimRefCmd  As String    '3          심사지침
    SimIcdCod  As String
End Type
    
    '------------------------------------------------------
    '13) 담당코드 UidMst
    '------------------------------------------------------
'Type UidMstRec
'    UidCod     As String    'UidMstKey 담당자코드
'    UidNam     As String    '          담당자 성명
'    UidPwd     As String    '          Password
'    UidDepCod  As String    '          소속과목
'    UidSecLev  As String    '          보안수준
'    UidEmpNum  As String    '          회사번호
'    UidPrtCod  As String    '          부서명
'    UiddtrYon  As String    '          의사여부
'    UidSpcYon  As String    '          특진여부
'    UidAssLev  As String    '          보조구분(직급)
'    UidPosCod  As String    '          재고부서
'    UidSgnDir  As String    '          Sign Image
'    UidSgnFle  As String    '          Sign Image
'    UidLicNum  As String    '          의사면허번호
'    UidTelNum  As String
'    UidMalAdd  As String
'    UidAdpDte  As String    '          적용개시일
'    UidExpDte  As String    '          적용종료일
'
'End Type

Type UidMstRec
    UidCod     As String    'UidMstKey 담당자코드
    UidNam     As String    '          담당자 성명
    UidPwd     As String    '          Password
    UidDepCod  As String    '          소속과목
    UidSecLev  As String    '          보안수준
    UidEmpNum  As String    '          회사번호
    UidPrtCod  As String    '          부서명
    UidDtrYon  As String    '          의사여부
    UidSpcYon  As String    '          특진여부
    UidAssLev  As String    '          보조구분(직급)
    UidPosCod  As String    '          재고부서
    UidSgnDir  As String    '          Sign Image
    UidSgnFle  As String    '          Sign Image
    UidLicNum  As String    '          의사면허번호
    UidTelNum  As String    '          의사전화번호
    UidMalAdd  As String    '          E-Mail Address
    UidAdpDte  As String    '          적용개시일
    UidEndDte  As String
    UidSpcNum  As String    '          전문의(Specialist) 면허번호
End Type

Type SecMstRec
    SecUidCod As String     '사용자 코드
    SecPrgCod As String     '프로그램명
    SecAllPwr As String     '모든권한
    SecRedOny As String     '읽기만 허용
End Type
    
    '------------------------------------------------------
    '14) 담당코드 History UidHst
    '------------------------------------------------------
'Type UidHstRec
'    UidCod     As String    'UidMstKey 담당자코드
'    UidAdpKey  As String    '          적용개시일
'    UidNam     As String    '          담당자 성명
'    UidPwd     As String    '          Password
'    UidDepCod  As String    '          소속과목
'    UidSecLev  As String    '          보안수준
'    UidEmpNum  As String    '          회사번호
'    UidPrtCod  As String    '          부서명
'    UiddtrYon  As String    '          의사여부
'    UidSpcYon  As String    '          특진여부
'    UidAssLev  As String    '          보조구분(직급)
'    UidPosCod  As String    '          재고부서
'    UidSgnDir  As String    '          Sign Image
'    UidSgnFle  As String    '          Sign Image
'    UidLicNum  As String    '          의사면허번호
'    UidTelNum  As String
'    UidMalAdd  As String
'    UidAdpDte  As String    '          적용개시일
'    UidExpDte  As String    '          적용종료일
'
'End Type

Type UidHstRec
    UidCod     As String    'UidMstKey 담당자코드
    UidAdpKey  As String    '          적용개시일
    UidNam     As String    '          담당자 성명
    UidPwd     As String    '          Password
    UidDepCod  As String    '          소속과목
    UidSecLev  As String    '          보안수준
    UidEmpNum  As String    '          회사번호
    UidPrtCod  As String    '          부서명
    UidDtrYon  As String    '          의사여부
    UidSpcYon  As String    '          특진여부
    UidAssLev  As String    '          보조구분(직급)
    UidPosCod  As String    '          재고부서
    UidSgnDir  As String    '          Sign Image
    UidSgnFle  As String    '          Sign Image
    UidLicNum  As String    '          의사면허번호
    UidTelNum  As String
    UidMalAdd  As String
    UidAdpDte  As String    '          적용개시일
    UidEndDte  As String
    UidSpcNum  As String    '          전문의(Specialist) 면허번호
End Type

'Index      ChtManMstRomRakStt      K-1,K-2,D-4
'           ChtManMstRomStt         K-1,D-4
'           ChtManMstRomRakCabStt   K-1,K-2,K-3,D-4
'           ChtManMstCht            D-1
'           ChtManMstRes            D-3
'           ChtManMstNam            D-2
Type ChtManMstRec
    ChtManRomNum  As String         'Key-1 Room 번호
    ChtManRakNum  As String * 2     'Key-2 Rack 번호
    ChtManCabNum  As String * 3     'Key-3 Cabinet 번호
    ChtManDtlNum  As String * 4     'Key-4 Detail 번호
    ChtManChtNum  As String * 8     'D-1 차트번호
    ChtManPatNam  As String         'D-2 환자명
    ChtManResNum  As String         'D-3 주민번호
    ChtManCurStt  As String         'D-4 상태구분
End Type

'--------------
'''지시형 오더
'--------------
Type IdcMstRec
    IdcDepCod      As String           'K-1 과목
    IdcDtrCod      As String           'K-2 의사
    IdcIdcCod      As String           'K-3 오더 범주
    IdcOdrCod      As String           'K-3 오더 코드
    IdcMsgNam      As String           'D-1 오더 메세지
    IdcOdrSeq       As String           'D-2 오더 순서
End Type
    
    
Type KgoMstRec
    KgoSeeCod  As String        '수가코드
    KgoUntCod  As String        'Kg 단위 1,2,...
    KgoOdrQty  As String        'KgoOdrQty
    KgoDtrQty  As String        'KgoDtrQty
    KgoSpcRem  As String        'KgoSpcRem
    KgoUpdDtm  As String        'KgoUpdDtm
    KgoUidCod  As String        'KgoUidCod

End Type
'--------------
'''지시형 오더
'--------------
Type OutMstRec
    
    OutOdrDte As String     '처방일자
    OutNum    As String     '최종교부번호
    OutUpdDtm As String     '수정일시
    
End Type


'------------------------------------------------------
'''TPM 코드 TpmMst
'------------------------------------------------------
Type TpmMstRec
    TpmCod  As String       'TpmMstKey TPM 코드
    TpmCodNam  As String    'TpmMstKey TPM 코드 명칭
End Type


Type DssDtlRec

    DssDtlTblCod  As String    'DssDtlKey 테이블 코드
    DssDtlCod     As String    'DssDtlKey 상세코드
    DssDtlCodNam  As String    '          상세코드명칭
    
End Type



Public Sub TpmMstStore(sPrmKey As String, sPrmValue As String, tPrmTpmData As TpmMstRec)
    
    sPrmKey = tPrmTpmData.TpmCod & Chr(5)
    sPrmKey = sPrmKey & tPrmTpmData.TpmCodNam & Chr(5)
    
End Sub
    
Public Sub TpmMstLoad(sPrmValue As String, tPrmTpmData As TpmMstRec)

    On Error GoTo TpmMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmTpmData.TpmCod = vVal(i)
    i = i + 1
    tPrmTpmData.TpmCodNam = vVal(i)
    
    Exit Sub

TpmMstLoad_ErrorTraping:
    Resume Next

End Sub

    
    

Public Sub IdcMstLoad(sPrmValue As String, tPrmIdcData As IdcMstRec)
    
    On Error GoTo IdcMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmIdcData.IdcDepCod = vVal(i) '0
    i = i + 1
    tPrmIdcData.IdcDtrCod = vVal(i) '1
    i = i + 1
    tPrmIdcData.IdcIdcCod = vVal(i) '2
    i = i + 1
    tPrmIdcData.IdcOdrCod = vVal(i) '3
    i = i + 1
    tPrmIdcData.IdcMsgNam = vVal(i) '4
    i = i + 1
    tPrmIdcData.IdcOdrSeq = vVal(i) '5
    
    Exit Sub

IdcMstLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub IdcMstStore(sPrmKey As String, sPrmValue As String, tPrmIdcData As IdcMstRec)

    sPrmKey = tPrmIdcData.IdcDepCod & Chr(5)
    sPrmKey = sPrmKey & tPrmIdcData.IdcDtrCod & Chr(5)
    sPrmKey = sPrmKey & tPrmIdcData.IdcIdcCod & Chr(5)
    sPrmKey = sPrmKey & tPrmIdcData.IdcOdrCod & Chr(5)
    
    sPrmValue = tPrmIdcData.IdcMsgNam & Chr(5)
    sPrmValue = sPrmValue & tPrmIdcData.IdcOdrSeq & Chr(5)

End Sub
    
Public Sub AccMstLoad(sPrmValue As String, tPrmAccData As AccMstRec)

    On Error GoTo AccMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmAccData.AccCod = vVal(i)
    i = i + 1
    tPrmAccData.AccCodNam = vVal(i)
    i = i + 1
    tPrmAccData.AccClsCod = vVal(i)
    i = i + 1
    tPrmAccData.AccEmpYon = vVal(i)
    i = i + 1
    tPrmAccData.AccFncYon = vVal(i)
    i = i + 1
    tPrmAccData.AccAmtYon = vVal(i)
    i = i + 1
    tPrmAccData.AccFdgRat = vVal(i)
    i = i + 1
    tPrmAccData.AccSdgRat = vVal(i)
    i = i + 1
    tPrmAccData.AccCalRat = vVal(i)
    i = i + 1
    tPrmAccData.AccIncDiv = vVal(i)
    i = i + 1
    tPrmAccData.AccInsYon = vVal(i)
    i = i + 1
    tPrmAccData.AccAssTyp = vVal(i)
    i = i + 1
    tPrmAccData.AccGbnTyp = vVal(i)
    i = i + 1
    tPrmAccData.AccConYon = vVal(i)
    '-------------------'
    '- 상세할인 작업중 -'
    i = i + 1
    tPrmAccData.AccInsMat = vVal(i)
    i = i + 1
    tPrmAccData.AccInsAct = vVal(i)
    i = i + 1
    tPrmAccData.AccNonMat = vVal(i)
    i = i + 1
    tPrmAccData.AccNonAct = vVal(i)
    i = i + 1
    tPrmAccData.AccSpcAmt = vVal(i)
    '-------------------'
    i = i + 1
    tPrmAccData.AccShwYon = vVal(i)
    
    Exit Sub

AccMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub AccMstRead(sAccCod As String, AccData As AccMstRec)
    
    Dim sAccMstCurKey As String
    Dim sAccMstCmpKey As String
    Dim sAccMstRetVal As String
    
    sAccMstCurKey = sAccCod & Chr(5)
    sAccMstCurKey = mSetReadEqual("AccMst", sAccMstCurKey, sAccMstRetVal)
    Call AccMstLoad(sAccMstRetVal, AccData)
    
    Exit Sub

AccMstLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub SecMstLoad(sPrmValue As String, Secdata As SecMstRec)

  On Error GoTo SecMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 10)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    Secdata.SecUidCod = vVal(i) '사용자 코드
    i = i + 1
    Secdata.SecPrgCod = vVal(i) '프로그램명
    i = i + 1
    Secdata.SecAllPwr = vVal(i) '모든권한
    i = i + 1
    Secdata.SecRedOny = vVal(i) '읽기만 허용
    
    Exit Sub

SecMstLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub SecMstStore(sCurKey As String, sRetVal As String, Secdata As SecMstRec)

    sCurKey = Secdata.SecUidCod & Chr(5) & Secdata.SecPrgCod & Chr(5)
    
    sRetVal = Secdata.SecAllPwr & Chr(5)
    sRetVal = sRetVal & Secdata.SecRedOny & Chr(5)

End Sub

    
Public Sub AccMstStore(sPrmKey As String, sPrmValue As String, tPrmAccData As AccMstRec)

    
    sPrmKey = tPrmAccData.AccCod & Chr(5)
    
    sPrmValue = tPrmAccData.AccCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccClsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccEmpYon & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccFncYon & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccAmtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccFdgRat & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccSdgRat & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccCalRat & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmAccData.AccIncDiv), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccInsYon & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccAssTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccGbnTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccConYon & Chr(5)
    '-------------------'
    '- 상세할인 작업중 -'
    sPrmValue = sPrmValue & tPrmAccData.AccInsMat & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccInsAct & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccNonMat & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccNonAct & Chr(5)
    sPrmValue = sPrmValue & tPrmAccData.AccSpcAmt & Chr(5)
    '-------------------'
    sPrmValue = sPrmValue & tPrmAccData.AccShwYon & Chr(5)
    
End Sub

    
Public Sub AddMstLoad(sPrmValue As String, tPrmAddData As AddMstRec)

    On Error GoTo AddMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmAddData.AddCod = vVal(i)
    i = i + 1
    tPrmAddData.AddAdpDte = vVal(i)
    i = i + 1
    tPrmAddData.AddCodNam = vVal(i)
    i = i + 1
    tPrmAddData.AddRatOne = vVal(i)
    i = i + 1
    tPrmAddData.AddRatTwo = vVal(i)
    i = i + 1
    tPrmAddData.AddRatThr = vVal(i)
    i = i + 1
    tPrmAddData.AddRatTot = vVal(i)
    
    Exit Sub

AddMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub AddMstStore(sPrmKey As String, sPrmValue As String, tPrmAddData As AddMstRec)

    
    sPrmKey = tPrmAddData.AddCod & Chr(5)
    sPrmKey = sPrmKey & tPrmAddData.AddAdpDte & Chr(5)
    
    sPrmValue = tPrmAddData.AddCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmAddData.AddRatOne & Chr(5) & tPrmAddData.AddRatTwo & Chr(5)
    sPrmValue = sPrmValue & tPrmAddData.AddRatThr & Chr(5)
    sPrmValue = sPrmValue & tPrmAddData.AddRatTot & Chr(5)
    
End Sub

    
Public Sub AssMstLoad(sPrmValue As String, tPrmAssData As AssMstRec)

    On Error GoTo AssMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmAssData.AssCod = vVal(i)
    i = i + 1
    tPrmAssData.AssInsCod = vVal(i)
    i = i + 1
    tPrmAssData.AssCodNam = vVal(i)
    i = i + 1
    tPrmAssData.AssCtyTyp = vVal(i)
    i = i + 1
    tPrmAssData.AssUpdDtm = vVal(i)
    i = i + 1
    tPrmAssData.AssUidCod = vVal(i)
    i = i + 1
    tPrmAssData.AssAddDtl = vVal(i)
    i = i + 1
    tPrmAssData.AssTelNum = vVal(i)
    i = i + 1
    tPrmAssData.AssFaxNum = vVal(i)
    i = i + 1
    tPrmAssData.AssEmlAdr = vVal(i)
    
    Exit Sub

AssMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub AssMstStore(sPrmKey As String, sPrmValue As String, tPrmAssData As AssMstRec)

    
    sPrmKey = tPrmAssData.AssCod & Chr(5)
    
    sPrmValue = tPrmAssData.AssInsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmAssData.AssCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmAssData.AssCtyTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmAssData.AssUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmAssData.AssUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmAssData.AssAddDtl & Chr(5)
    sPrmValue = sPrmValue & tPrmAssData.AssTelNum & Chr(5)
    sPrmValue = sPrmValue & tPrmAssData.AssFaxNum & Chr(5)
    sPrmValue = sPrmValue & tPrmAssData.AssEmlAdr & Chr(5)
    
End Sub

    
Public Sub BedMstLoad(sPrmValue As String, tPrmBedData As BedMstRec)

    On Error GoTo BedMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmBedData.BedWrdCod = vVal(i)
    i = i + 1
    tPrmBedData.BedRomCod = vVal(i)
    i = i + 1
    tPrmBedData.BedCod = vVal(i)
    i = i + 1
    tPrmBedData.BedSttCod = vVal(i)
    i = i + 1
    tPrmBedData.BedChtNum = vVal(i)
    i = i + 1
    tPrmBedData.BedPatNam = vVal(i)
    i = i + 1
    tPrmBedData.BedPatSex = vVal(i)
    i = i + 1
    tPrmBedData.BedOcmNum = vVal(i)
    i = i + 1
    tPrmBedData.BedIcdNam = vVal(i)
    i = i + 1
    tPrmBedData.BedDepCod = vVal(i)
    i = i + 1
    tPrmBedData.BedPatSts = vVal(i)
    i = i + 1
    tPrmBedData.BedTrnDtm = vVal(i)
    i = i + 1
    tPrmBedData.BedCsnDtm = vVal(i)
    i = i + 1
    tPrmBedData.BedCsnTyp = vVal(i)
    i = i + 1
    tPrmBedData.BedBirDay = vVal(i)
    i = i + 1
    tPrmBedData.BedLevTyp = vVal(i)
    i = i + 1
    tPrmBedData.BedDtrCod = vVal(i)
    i = i + 1
    tPrmBedData.BedAcuCod = vVal(i)
    i = i + 1
    tPrmBedData.BedIntTel = vVal(i)
    
    Exit Sub

BedMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub BedMstStore(sPrmKey As String, sPrmValue As String, tPrmBedData As BedMstRec)

    
    sPrmKey = tPrmBedData.BedWrdCod & Chr(5)
    sPrmKey = sPrmKey & tPrmBedData.BedRomCod & Chr(5)
    sPrmKey = sPrmKey & tPrmBedData.BedCod & Chr(5)
    
    sPrmValue = tPrmBedData.BedSttCod & Chr(5)
    sPrmValue = sPrmValue & Format(tPrmBedData.BedChtNum, "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedPatSex & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmBedData.BedOcmNum), "@@@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedIcdNam & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedPatSts & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedTrnDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedCsnDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedCsnTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedBirDay & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedLevTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedDtrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedAcuCod & Chr(5)
    sPrmValue = sPrmValue & tPrmBedData.BedIntTel & Chr(5)
    
End Sub

    
Public Sub CalMstLoad(sPrmValue As String, tPrmCalData As CalMstRec)

    On Error GoTo CalMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmCalData.CalAddCod = vVal(i)
    i = i + 1
    tPrmCalData.CalAdpDte = vVal(i)
    i = i + 1
    tPrmCalData.CalCodNam = vVal(i)
    i = i + 1
    tPrmCalData.CalAddRat = vVal(i)
    i = i + 1
    tPrmCalData.CalAddPrc = vVal(i)
    i = i + 1
    tPrmCalData.CalAddAmt = vVal(i)
    i = i + 1
    tPrmCalData.CalFeeCod = vVal(i)
    i = i + 1
    tPrmCalData.CalRatOne = vVal(i)
    i = i + 1
    tPrmCalData.CalFeeOne = vVal(i)
    i = i + 1
    tPrmCalData.CalRatTwo = vVal(i)
    i = i + 1
    tPrmCalData.CalFeeTwo = vVal(i)
    i = i + 1
    tPrmCalData.CalRatThr = vVal(i)
    i = i + 1
    tPrmCalData.CalFeeThr = vVal(i)
    
    Exit Sub

CalMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CalMstStore(sPrmKey As String, sPrmValue As String, tPrmCalData As CalMstRec)

    
    sPrmKey = tPrmCalData.CalAddCod & Chr(5)
    sPrmKey = sPrmKey & tPrmCalData.CalAdpDte & Chr(5)
    
    sPrmValue = tPrmCalData.CalCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalAddRat & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalAddPrc & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalAddAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalFeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalRatOne & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalFeeOne & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalRatTwo & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalFeeTwo & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalRatThr & Chr(5)
    sPrmValue = sPrmValue & tPrmCalData.CalFeeThr & Chr(5)
    
End Sub

    
Public Sub CmtMstLoad(sPrmValue As String, tPrmCmtData As CmtMstRec)

    On Error GoTo CmtMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmCmtData.CmtUidCod = vVal(i)
    i = i + 1
    tPrmCmtData.CmtUsgPgm = vVal(i)
    i = i + 1
    tPrmCmtData.CmtCod = vVal(i)
    i = i + 1
    tPrmCmtData.CmtCodNam = vVal(i)
    
    Exit Sub

CmtMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CmtMstStore(sPrmKey As String, sPrmValue As String, tPrmCmtData As CmtMstRec)

    
    sPrmKey = tPrmCmtData.CmtUidCod & Chr(5)
    sPrmKey = sPrmKey & tPrmCmtData.CmtUsgPgm & Chr(5)
    sPrmKey = sPrmKey & tPrmCmtData.CmtCod & Chr(5)
    
    sPrmValue = tPrmCmtData.CmtCodNam & Chr(5)
    
End Sub

    
Public Sub CspMstLoad(sPrmValue As String, tPrmData As CspMstRec)

    On Error GoTo CspMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.CspUsgPrt = vVal(i)
    i = i + 1
    tPrmData.CspSeq = vVal(i)
    i = i + 1
    tPrmData.CspCsrCod = vVal(i)
    i = i + 1
    tPrmData.CspDepTyp = vVal(i)
    
    Exit Sub

CspMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CspMstStore(sPrmKey As String, sPrmValue As String, tPrmData As CspMstRec)

    
    sPrmKey = tPrmData.CspUsgPrt & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmData.CspSeq), "@@@") & Chr(5)
    
    sPrmValue = tPrmData.CspCsrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CspDepTyp & Chr(5)
    
End Sub

    
Public Sub CsrMstLoad(sPrmValue As String, tPrmData As CsrMstRec)

    On Error GoTo CsrMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.CsrDepTyp = vVal(i)
    i = i + 1
    tPrmData.CsrCod = vVal(i)
    i = i + 1
    tPrmData.CsrCsmYon = vVal(i)
    i = i + 1
    tPrmData.CsrCodNam = vVal(i)
    i = i + 1
    tPrmData.CsrOmsNam = vVal(i)
    i = i + 1
    tPrmData.CsrUntCod = vVal(i)
    i = i + 1
    tPrmData.CsrUntQty = vVal(i)
    i = i + 1
    tPrmData.CsrSeeCod = vVal(i)
    
    Exit Sub

CsrMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CsrMstStore(sPrmKey As String, sPrmValue As String, tPrmData As CsrMstRec)

    
    sPrmKey = tPrmData.CsrDepTyp & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CsrCod & Chr(5)
    sPrmKey = sPrmKey & tPrmData.CsrCsmYon & Chr(5)
    
    sPrmValue = tPrmData.CsrCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsrOmsNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsrUntCod & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsrUntQty & Chr(5)
    sPrmValue = sPrmValue & tPrmData.CsrSeeCod & Chr(5)
    
End Sub

    
Public Sub CutMstLoad(sPrmValue As String, tPrmCutData As CutMstRec)

    On Error GoTo CutMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmCutData.CutGub = vVal(i)
    i = i + 1
    tPrmCutData.CutCod = vVal(i)
    i = i + 1
    tPrmCutData.CutAdpDte = vVal(i)
    i = i + 1
    tPrmCutData.CutExpDte = vVal(i)
    i = i + 1
    tPrmCutData.CutNam = vVal(i)
    i = i + 1
    tPrmCutData.CutInsMat = vVal(i)
    i = i + 1
    tPrmCutData.CutInsAct = vVal(i)
    i = i + 1
    tPrmCutData.CutNonMat = vVal(i)
    i = i + 1
    tPrmCutData.CutNonAct = vVal(i)
    i = i + 1
    tPrmCutData.CutStgGen = vVal(i)
    i = i + 1
    tPrmCutData.CutStgCar = vVal(i)
    i = i + 1
    tPrmCutData.CutStgIns = vVal(i)
    i = i + 1
    tPrmCutData.CutStgWrk = vVal(i)
    i = i + 1
    tPrmCutData.CutStgBoh = vVal(i)
    i = i + 1
    tPrmCutData.CutUpdDtm = vVal(i)
    i = i + 1
    tPrmCutData.CutUidCod = vVal(i)
    i = i + 1
    tPrmCutData.CutNum = vVal(i)
    Exit Sub

CutMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub CutMstStore(sPrmKey As String, sPrmValue As String, tPrmCutData As CutMstRec)

    
    sPrmKey = tPrmCutData.CutGub & Chr(5)
    sPrmKey = sPrmKey & tPrmCutData.CutCod & Chr(5)
    sPrmKey = sPrmKey & tPrmCutData.CutAdpDte & Chr(5)
    
    sPrmValue = tPrmCutData.CutExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutNam & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutInsMat & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutInsAct & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutNonMat & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutNonAct & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutStgGen & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutStgCar & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutStgIns & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutStgWrk & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutStgBoh & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmCutData.CutNum & Chr(5)
    
End Sub

    
Public Sub DepMstGrpLoad(sPrmValue As String, tPrmDepGrpData As DepMstGrpRec)

    On Error GoTo DepMstGrpLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmDepGrpData.DepGrpCod = vVal(i)
    i = i + 1
    tPrmDepGrpData.DepCod = vVal(i)
    
    Exit Sub

DepMstGrpLoad_ErrorTraping:
    Resume Next

End Sub
    
    
Public Sub DepMstGrpStore(sPrmKey As String, sPrmValue As String, tPrmDepGrpData As DepMstGrpRec)

    
    sPrmKey = tPrmDepGrpData.DepGrpCod & Chr(5)
    
    sPrmValue = tPrmDepGrpData.DepCod & Chr(5)
    
End Sub

    
Public Sub DepMstLoad(sPrmValue As String, tPrmDepData As DepMstRec)

    On Error GoTo DepMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmDepData.DepCod = vVal(i)
    i = i + 1
    tPrmDepData.DepAdpDte = vVal(i)
    i = i + 1
    tPrmDepData.DepKorNam = vVal(i)
    i = i + 1
    tPrmDepData.DepEngNam = vVal(i)
    i = i + 1
    tPrmDepData.DepGrpCod = vVal(i)
    i = i + 1
    tPrmDepData.DepBilCod = vVal(i)
    i = i + 1
    tPrmDepData.DepBilAra = vVal(i)
    i = i + 1
    tPrmDepData.DepBilSeq = vVal(i)
    i = i + 1
    tPrmDepData.DepSndYon = vVal(i)
    i = i + 1
    tPrmDepData.DepHspTyp = vVal(i)
    i = i + 1
    tPrmDepData.DepMdcTyp = vVal(i)
    i = i + 1
    tPrmDepData.DepMisPos = vVal(i)
    i = i + 1
    tPrmDepData.DepIncTyp = vVal(i)
    i = i + 1
    tPrmDepData.DepDgsCod = vVal(i)
    
    Exit Sub

DepMstLoad_ErrorTraping:
    Resume Next

End Sub
    
    '**************************************
    '       진료과 찾기
    '
    '   sDepCod : 진료과 코드
    '   sAdpDte : 적용일자 (ex. 19961231)
    '**************************************
Public Sub DepMstRead(sDepCod As String, sAdpDte As String, DepData As DepMstRec)
    
    Dim sDepMstCurKey As String, sDepMstCmpKey As String, sDepMstRetVal As String
    Dim sAdpDate As String
    
    sAdpDate = Left(sAdpDte, 8)
    
    sDepMstCmpKey = sDepCod & Chr(5)
    sDepMstCurKey = sDepMstCmpKey & sAdpDate & "99" & Chr(5)
    sDepMstCurKey = mSetPrev("DepMst", sDepMstCurKey)
    sDepMstCurKey = mReadPrev("DepMst", sDepMstCurKey, sDepMstCmpKey, sDepMstRetVal)
    
    DepMstLoad sDepMstRetVal, DepData
    
    Exit Sub

DepMstLoad_ErrorTraping:
    Resume Next

End Sub
    
    
Public Sub DepMstStore(sPrmKey As String, sPrmValue As String, tPrmDepData As DepMstRec)

    
    sPrmKey = tPrmDepData.DepCod & Chr(5)
    sPrmKey = sPrmKey & tPrmDepData.DepAdpDte & Chr(5)
    
    sPrmValue = tPrmDepData.DepKorNam & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepEngNam & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepGrpCod & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepBilCod & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepBilAra & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmDepData.DepBilSeq), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepSndYon & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepHspTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepMdcTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepMisPos & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepIncTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmDepData.DepDgsCod & Chr(5)
    
End Sub

    
Public Sub DrgMstLoad(sPrmValue As String, tPrmDrgData As DrgMstRec)

    On Error GoTo DrgMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmDrgData.DrgCod = vVal(i)
    i = i + 1
    tPrmDrgData.DrgOdrDay = vVal(i)
    i = i + 1
    tPrmDrgData.DrgAdpDte = vVal(i)
    i = i + 1
    tPrmDrgData.DrgCorAmt = vVal(i)
    i = i + 1
    tPrmDrgData.DrgAskAmt = vVal(i)
    
    Exit Sub

DrgMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub DrgMstStore(sPrmKey As String, sPrmValue As String, tPrmDrgData As DrgMstRec)

    
    sPrmKey = tPrmDrgData.DrgCod & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmDrgData.DrgOdrDay, "@@") & Chr(5)
    sPrmKey = sPrmKey & tPrmDrgData.DrgAdpDte & Chr(5)
    
    sPrmValue = tPrmDrgData.DrgCorAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmDrgData.DrgAskAmt & Chr(5)
    
End Sub

    
Public Sub DtlMstLoad(sPrmValue As String, tPrmDtlData As DtlMstRec)

    On Error GoTo DtlMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmDtlData.DtlTblCod = vVal(i)
    i = i + 1
    tPrmDtlData.DtlCod = vVal(i)
    i = i + 1
    tPrmDtlData.DtlCodNam = vVal(i)
    
    Exit Sub

DtlMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub DtlMstStore(sPrmKey As String, sPrmValue As String, tPrmDtlData As DtlMstRec)

    
    sPrmKey = tPrmDtlData.DtlTblCod & Chr(5)
    sPrmKey = sPrmKey & tPrmDtlData.DtlCod & Chr(5)
    
    sPrmValue = tPrmDtlData.DtlCodNam & Chr(5)
    
End Sub

   
Public Sub DssDtlStore(sPrmKey As String, sPrmValue As String, tPrmDssData As DssDtlRec)

    
    sPrmKey = tPrmDssData.DssDtlTblCod & Chr(5)
    sPrmKey = sPrmKey & tPrmDssData.DssDtlCod & Chr(5)
    
    sPrmValue = tPrmDssData.DssDtlCodNam & Chr(5)
    
End Sub
   
   
Public Sub EtcMstLoad(sPrmValue As String, tPrmEtcData As EtcMstRec)

    On Error GoTo EtcMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmEtcData.EtcItmCod = vVal(i)
    i = i + 1
    tPrmEtcData.EtcCod = vVal(i)
    i = i + 1
    tPrmEtcData.EtcCodNam = vVal(i)
    
    Exit Sub

EtcMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub EtcMstStore(sPrmKey As String, sPrmValue As String, tPrmEtcData As EtcMstRec)

    
    sPrmKey = tPrmEtcData.EtcItmCod & Chr(5)
    sPrmKey = sPrmKey & tPrmEtcData.EtcCod & Chr(5)
    
    sPrmValue = tPrmEtcData.EtcCodNam & Chr(5)
    
End Sub

    
Public Sub ExeMstLoad(sPrmValue As String, tPrmExeData As ExeMstRec)

    On Error GoTo ExeMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmExeData.ExeCod = vVal(i)
    i = i + 1
    tPrmExeData.ExeMainNam = vVal(i)
    'i = i + 1
    i = i + 1
    'tPrmExeData.ExeSubIdx = vVal(i)
    i = i + 1
    tPrmExeData.ExeExeNam = vVal(i)
    i = i + 1
    tPrmExeData.ExeFlg = vVal(i)
    
    Exit Sub

ExeMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ExeMstStore(sPrmKey As String, sPrmValue As String, tPrmExeMstData As ExeMstRec)

    
    sPrmKey = tPrmExeMstData.ExeCod & Chr(5)
    sPrmKey = sPrmKey & tPrmExeMstData.ExeMainNam & Chr(5)
    
    sPrmKey = sPrmKey & tPrmExeMstData.ExeExeNam & Chr(5)
    
    sPrmValue = tPrmExeMstData.ExeFlg & Chr(5)
    
    
End Sub

Public Sub FeeHstLoad(sPrmValue As String, tPrmFeeData As FeeHstRec)

    On Error GoTo FeeHstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmFeeData.FeeElcCod = vVal(i)
    i = i + 1
    tPrmFeeData.FeeAdpKey = vVal(i)
    i = i + 1
    tPrmFeeData.FeeEngNam = vVal(i)
    i = i + 1
    tPrmFeeData.FeeKorNam = vVal(i)
    i = i + 1
    tPrmFeeData.FeeInsAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeGenAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeCarAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeAdpDte = vVal(i)
    i = i + 1
    tPrmFeeData.FeeExpDte = vVal(i)
    i = i + 1
    tPrmFeeData.FeeUidCod = vVal(i)
    i = i + 1
    tPrmFeeData.FeeUpdDtm = vVal(i)
    i = i + 1
    tPrmFeeData.FeeWrkAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeGudAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeLftAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeInsAdp = vVal(i)
    i = i + 1
    tPrmFeeData.FeeMakNam = vVal(i)
    i = i + 1
    tPrmFeeData.FeeDrgCod = vVal(i)
    i = i + 1
    tPrmFeeData.FeeUntCod = vVal(i)
    i = i + 1
    tPrmFeeData.FeeCodDiv = vVal(i)
    i = i + 1
    tPrmFeeData.FeeExtAmt = vVal(i)
    
    Exit Sub

FeeHstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub FeeHstStore(sPrmKey As String, sPrmValue As String, tPrmFeeData As FeeHstRec)

    
    sPrmKey = tPrmFeeData.FeeElcCod & Chr(5)
    sPrmKey = sPrmKey & tPrmFeeData.FeeAdpKey & Chr(5)
    
    sPrmValue = tPrmFeeData.FeeEngNam & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeKorNam & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeGenAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeCarAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeWrkAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeGudAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeLftAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeInsAdp & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeMakNam & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeDrgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeUntCod & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeCodDiv & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeExtAmt & Chr(5)
    
End Sub

    
Public Sub FeeMstLoad(sPrmValue As String, tPrmFeeData As FeeMstRec)

    On Error GoTo FeeMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmFeeData.FeeElcCod = vVal(i)
    i = i + 1
    tPrmFeeData.FeeEngNam = vVal(i)
    i = i + 1
    tPrmFeeData.FeeKorNam = vVal(i)
    i = i + 1
    tPrmFeeData.FeeInsAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeGenAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeCarAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeAdpDte = vVal(i)
    i = i + 1
    tPrmFeeData.FeeExpDte = vVal(i)
    i = i + 1
    tPrmFeeData.FeeUidCod = vVal(i)
    i = i + 1
    tPrmFeeData.FeeUpdDtm = vVal(i)
    i = i + 1
    tPrmFeeData.FeeWrkAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeGudAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeLftAmt = vVal(i)
    i = i + 1
    tPrmFeeData.FeeInsAdp = vVal(i)
    i = i + 1
    tPrmFeeData.FeeMakNam = vVal(i)
    i = i + 1
    tPrmFeeData.FeeDrgCod = vVal(i)
    i = i + 1
    tPrmFeeData.FeeUntCod = vVal(i)
    i = i + 1
    tPrmFeeData.FeeCodDiv = vVal(i)
    i = i + 1
    tPrmFeeData.FeeExtAmt = vVal(i)
    
    Exit Sub

FeeMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub FeeMstStore(sPrmKey As String, sPrmValue As String, tPrmFeeData As FeeMstRec)

    
    sPrmKey = tPrmFeeData.FeeElcCod & Chr(5)
    
    sPrmValue = tPrmFeeData.FeeEngNam & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeKorNam & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeGenAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeCarAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeWrkAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeGudAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeLftAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeInsAdp & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeMakNam & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeDrgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeUntCod & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeCodDiv & Chr(5)
    sPrmValue = sPrmValue & tPrmFeeData.FeeExtAmt & Chr(5)
    
End Sub

    
Public Sub FnlMstLoad(sPrmValue As String, tPrmFnlData As FnlMstRec)

    On Error GoTo FnlMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmFnlData.FnlCod = vVal(i)
    i = i + 1
    tPrmFnlData.FnlNum = vVal(i)
    i = i + 1
    tPrmFnlData.FnlDte = vVal(i)
    
    Exit Sub

FnlMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub FnlMstStore(sPrmKey As String, sPrmValue As String, tPrmFnlData As FnlMstRec)

    
    sPrmKey = tPrmFnlData.FnlCod & Chr(5)
    
    sPrmValue = tPrmFnlData.FnlNum & Chr(5)
    
    sPrmValue = sPrmValue & tPrmFnlData.FnlDte & Chr(5)
    
End Sub

    
Public Sub FtdMstLoad(sPrmValue As String, tPrmFtdData As FtdMstRec)

    On Error GoTo FtdMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    '20010701
    i = i + 1
    tPrmFtdData.FtdDgsCod = vVal(i)
    '20010701
    i = i + 1
    tPrmFtdData.FtdDgsNfs = vVal(i)
    i = i + 1
    tPrmFtdData.FtdDgsDnh = vVal(i)
    i = i + 1
    tPrmFtdData.FtdAgeDiv = vVal(i)
    i = i + 1
    tPrmFtdData.FtdAdpDte = vVal(i)
    i = i + 1
    tPrmFtdData.FtdCodNam = vVal(i)
    i = i + 1
    tPrmFtdData.FtdFeeCod = vVal(i)
    i = i + 1
    tPrmFtdData.FtdRsuAmt = vVal(i)
    i = i + 1
    tPrmFtdData.FtdSpcAmt = vVal(i)
    
    Exit Sub

FtdMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub FtdMstStore(sPrmKey As String, sPrmValue As String, tPrmFtdData As FtdMstRec)

    '20010701
    sPrmKey = tPrmFtdData.FtdDgsCod & Chr(5)
    'sPrmKey = stPrmFtdData.FtdDgsNfs & Chr(5)
    sPrmKey = sPrmKey & tPrmFtdData.FtdDgsNfs & Chr(5)
    '20010701
    sPrmKey = sPrmKey & tPrmFtdData.FtdDgsDnh & Chr(5)
    sPrmKey = sPrmKey & tPrmFtdData.FtdAgeDiv & Chr(5)
    sPrmKey = sPrmKey & tPrmFtdData.FtdAdpDte & Chr(5)
    
    sPrmValue = tPrmFtdData.FtdCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmFtdData.FtdFeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmFtdData.FtdRsuAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmFtdData.FtdSpcAmt & Chr(5)
    
End Sub
Public Sub GrdMstRead(sGrdCod As String, sAdpDte As String, grdData As GrdMstRec)
    
    Dim sGrdMstCurKey As String, sGrdMstCmpKey As String, sGrdMstRetVal As String
    Dim sAdpDate As String
    
    sAdpDate = Left(sAdpDte, 8)
    
    sGrdMstCmpKey = sGrdCod & Chr(5)
    sGrdMstCurKey = sGrdMstCmpKey & sAdpDate & "99" & Chr(5)
    
    sGrdMstCurKey = mSetPrev("GrdMst", sGrdMstCurKey)
    sGrdMstCurKey = mReadPrev("GrdMst", sGrdMstCurKey, sGrdMstCmpKey, sGrdMstRetVal)
    
    Call GrdMstLoad(sGrdMstRetVal, grdData)
    
    Exit Sub

GrdMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub GrdMstLoad(sPrmValue As String, tPrmData As GrdMstRec)

    On Error GoTo GrdMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmData.grdCod = vVal(i)
    i = i + 1
    tPrmData.GrdAdpDte = vVal(i)
    i = i + 1
    tPrmData.GrdNam = vVal(i)
    i = i + 1
    tPrmData.GrdFeeCod = vVal(i)
    i = i + 1
    tPrmData.GrdMedCod = vVal(i)
    i = i + 1
    tPrmData.GrdIsoCod = vVal(i)
    i = i + 1
    tPrmData.GrdAmtCod = vVal(i)
    i = i + 1
    tPrmData.GrdIcuCod = vVal(i)
    i = i + 1
    tPrmData.GrdRomGrd = vVal(i)
    
    Exit Sub

GrdMstLoad_ErrorTraping:
    Resume Next

End Sub
    
    
Public Sub GrdMstStore(sPrmKey As String, sPrmValue As String, tPrmGrdData As GrdMstRec)

    
    sPrmKey = tPrmGrdData.grdCod & Chr(5)
    sPrmKey = sPrmKey & tPrmGrdData.GrdAdpDte & Chr(5)
    
    sPrmValue = tPrmGrdData.GrdNam & Chr(5)
    sPrmValue = sPrmValue & tPrmGrdData.GrdFeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrdData.GrdMedCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrdData.GrdIsoCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrdData.GrdAmtCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrdData.GrdIcuCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrdData.GrdRomGrd & Chr(5)
    
End Sub
    
Public Sub GrpMstLoad(sPrmValue As String, tPrmGrpData As GrpMstRec)

    On Error GoTo GrpMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmGrpData.GrpCod = vVal(i)
    i = i + 1
    tPrmGrpData.GrpOdrSeq = vVal(i)
    i = i + 1
    tPrmGrpData.GrpOdrCod = vVal(i)
    i = i + 1
    tPrmGrpData.GrpOdrNam = vVal(i)
    i = i + 1
    tPrmGrpData.GrpAdpTyp = vVal(i)
    i = i + 1
    tPrmGrpData.GrpOdrQty = vVal(i)
    i = i + 1
    tPrmGrpData.GrpOdrTms = vVal(i)
    i = i + 1
    tPrmGrpData.GrpOdrDay = vVal(i)
    i = i + 1
    tPrmGrpData.GrpUsgCod = vVal(i)
    i = i + 1
    tPrmGrpData.GrpMthCod = vVal(i)
    i = i + 1
    tPrmGrpData.GrpSpmCod = vVal(i)
    i = i + 1
    tPrmGrpData.GrpDgsYon = vVal(i)
    i = i + 1
    tPrmGrpData.GrpInsYon = vVal(i)
    i = i + 1
    tPrmGrpData.GrpSpcYon = vVal(i)
    i = i + 1
    tPrmGrpData.GrpSpcCmt = vVal(i)
    i = i + 1
    tPrmGrpData.GrpDgsRol = vVal(i)
    i = i + 1
    tPrmGrpData.GrpItmCod = vVal(i)
    i = i + 1
    tPrmGrpData.GrpAstCod = vVal(i)
    i = i + 1
    tPrmGrpData.GrpSlpDep = vVal(i)
    i = i + 1
    tPrmGrpData.GrpDgsEtc = vVal(i)
    i = i + 1
    tPrmGrpData.GrpAdpDte = vVal(i)
    i = i + 1
    tPrmGrpData.GrpExpDte = vVal(i)
    
    Exit Sub

GrpMstLoad_ErrorTraping:
    Resume Next

End Sub
    
    
Public Sub GrpMstStore(sPrmKey As String, sPrmValue As String, tPrmGrpData As GrpMstRec)

    
    sPrmKey = tPrmGrpData.GrpCod & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmGrpData.GrpOdrSeq, "@@") & Chr(5)
    
    sPrmValue = tPrmGrpData.GrpOdrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpOdrNam & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpAdpTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpOdrQty & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpOdrTms & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpOdrDay & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpUsgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpMthCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpSpmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpDgsYon & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpInsYon & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpSpcCmt & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpDgsRol & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpItmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpAstCod & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpSlpDep & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpDgsEtc & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmGrpData.GrpExpDte & Chr(5)
    
End Sub

    
    
Public Sub HolMstLoad(sPrmValue As String, tPrmHolData As HolMstRec)

    On Error GoTo HolMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmHolData.HolDte = vVal(i)
    i = i + 1
    tPrmHolData.HolDteNam = vVal(i)
    
    Exit Sub

HolMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub HolMstStore(sPrmKey As String, sPrmValue As String, tPrmHolData As HolMstRec)

    
    sPrmKey = tPrmHolData.HolDte & Chr(5)
    
    sPrmValue = tPrmHolData.HolDteNam & Chr(5)
    
End Sub

    
Public Sub HspMstLoad(sPrmValue As String, tPrmHspData As HspMstRec)

    On Error GoTo HspMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmHspData.HspCod = vVal(i)
    i = i + 1
    tPrmHspData.HspNam = vVal(i)
    i = i + 1
    tPrmHspData.HspInsNum = vVal(i)
    i = i + 1
    tPrmHspData.HspInsNam = vVal(i)
    i = i + 1
    tPrmHspData.HspWrkNum = vVal(i)
    i = i + 1
    tPrmHspData.HspRgnNum = vVal(i)
    i = i + 1
    tPrmHspData.HspHspAdr = vVal(i)
    i = i + 1
    tPrmHspData.HspRcpNam = vVal(i)
    i = i + 1
    tPrmHspData.HspMdcYon = vVal(i)
    i = i + 1
    tPrmHspData.HspLmtYon = vVal(i)
    i = i + 1
    tPrmHspData.HspCloTim = vVal(i)
    i = i + 1
    tPrmHspData.HspOwnNam = vVal(i)
    i = i + 1
    tPrmHspData.HspZipCod = vVal(i)
    i = i + 1
    tPrmHspData.HspManNam = vVal(i)
    i = i + 1
    tPrmHspData.HspManRes = vVal(i)
    i = i + 1
    tPrmHspData.HspTelNum = vVal(i)
    i = i + 1
    tPrmHspData.HspBilSlp = vVal(i)
    i = i + 1
    tPrmHspData.HspBilDte = vVal(i)
    i = i + 1
    tPrmHspData.HspBilMan = vVal(i)
    i = i + 1
    tPrmHspData.HspBilCnt = vVal(i)
    i = i + 1
    tPrmHspData.HspAdmMth = vVal(i)
    i = i + 1
    tPrmHspData.HspAdmRcp = vVal(i)
    i = i + 1
    tPrmHspData.HspOcmRcp = vVal(i)
    i = i + 1
    tPrmHspData.HspMidRcp = vVal(i)
    i = i + 1
    tPrmHspData.HspMrpRcp = vVal(i)
    i = i + 1
    tPrmHspData.HspDisRcp = vVal(i)
    i = i + 1
    tPrmHspData.HspIcmRcp = vVal(i)
    i = i + 1
    tPrmHspData.HspPreRcp = vVal(i)
    i = i + 1
    tPrmHspData.HspGrnRcp = vVal(i)
    
    i = i + 1
    tPrmHspData.HspStaDgs = vVal(i)     ' 마감구분 (DtlMst의 "STATBL" 1:24시마감,2:회계일자마감)
    i = i + 1
    tPrmHspData.HspLgoPth = vVal(i)     ' 로고패스 (ex:\\Asp\Hnt.Cnv\Icon\로고.Bmp)
    i = i + 1
    tPrmHspData.HspRsvFee = vVal(i)     ' 로고패스 (ex:\\Asp\Hnt.Cnv\Icon\로고.Bmp)
        
    Exit Sub

HspMstLoad_ErrorTraping:
    Resume Next

End Sub
    
    '
    '   병원 정보 읽기
    '
Public Sub HspMstRead(sHspCod As String, tHspData As HspMstRec)
    
    Dim sCurKey As String
    Dim sRetVal As String
    
    sCurKey = sHspCod & Chr(5)
    sCurKey = mSetReadEqual("HspMst", sCurKey, sRetVal)
    If sCurKey <> "" Then
    Call HspMstLoad(sRetVal, tHspData)
    End If
    
    Exit Sub

HspMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub HspMstStore(sPrmKey As String, sPrmValue As String, tPrmHspData As HspMstRec)

    
    sPrmKey = tPrmHspData.HspCod & Chr(5)
    
    sPrmValue = tPrmHspData.HspNam & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspInsNum & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspInsNam & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspWrkNum & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspRgnNum & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspHspAdr & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspRcpNam & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspMdcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspLmtYon & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspCloTim & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspOwnNam & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspZipCod & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspManNam & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspManRes & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspTelNum & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspBilSlp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspBilDte & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspBilMan & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspBilCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspAdmMth & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspAdmRcp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspOcmRcp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspMidRcp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspMrpRcp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspDisRcp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspIcmRcp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspPreRcp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspGrnRcp & Chr(5)
    sPrmValue = sPrmValue & tPrmHspData.HspStaDgs & Chr(5) ' 마감구분 (DtlMst의 "STATBL" 1:24시마감,2:회계일자마감)
    sPrmValue = sPrmValue & tPrmHspData.HspLgoPth & Chr(5) ' 로고패스 (ex:\\Asp\Hnt.Cnv\Icon\로고.Bmp)
    sPrmValue = sPrmValue & tPrmHspData.HspRsvFee & Chr(5) ' 로고패스 (ex:\\Asp\Hnt.Cnv\Icon\로고.Bmp)
        
End Sub

Public Sub IcdMstLoad(sPrmValue As String, tPrmIcdData As IcdMstRec)

    On Error GoTo IcdMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmIcdData.IcdCod = vVal(i)
    i = i + 1
    tPrmIcdData.IcdEngNam = vVal(i)
    i = i + 1
    tPrmIcdData.IcdKorNam = vVal(i)
    i = i + 1
    tPrmIcdData.IcdDepAra = vVal(i)
    i = i + 1
    tPrmIcdData.IcdUpdDtm = vVal(i)
    i = i + 1
    tPrmIcdData.IcdUidCod = vVal(i)
    i = i + 1
    tPrmIcdData.IcdLagCod = vVal(i)
    i = i + 1
    tPrmIcdData.IcdMidCod = vVal(i)
    i = i + 1
    tPrmIcdData.IcdHanIcd = vVal(i)
    i = i + 1
    tPrmIcdData.IcdCanYon = vVal(i)
'****************************************************> 추가
    '20040115..HTS..
    i = i + 1
    tPrmIcdData.IcdVeeCod = vVal(i)
'****************************************************> 추가
    Exit Sub

IcdMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Function IcdMstRead(sIcdCod As String, tIcdData As IcdMstRec, sAdpDte As String) As Integer
    
    Dim sCurKey As String
    Dim sRetVal As String
    Dim sDBName As String
    
    If sAdpDte >= "20030101" Then
        sDBName = "Icd2003"
    Else
        sDBName = "IcdMst"
    End If

    sCurKey = sIcdCod & Chr(5)
    
    '20030203 lek edit====================================
    'sCurKey = mSetReadEqual("IcdMst", sCurKey, sRetVal)
    sCurKey = mSetReadEqual(sDBName, sCurKey, sRetVal)
    '=====================================================
    
    Call IcdMstLoad(sRetVal, tIcdData)
    
    If sCurKey <> "" Then
        IcdMstRead = True
    Else
        IcdMstRead = False
    End If
    Exit Function

IcdMstLoad_ErrorTraping:
    Resume Next

End Function
    
Public Sub IcdMstStore(sPrmKey As String, sPrmValue As String, tPrmIcdData As IcdMstRec)

    
    sPrmKey = tPrmIcdData.IcdCod & Chr(5)
    
    sPrmValue = tPrmIcdData.IcdEngNam & Chr(5)
    sPrmValue = sPrmValue & tPrmIcdData.IcdKorNam & Chr(5)
    sPrmValue = sPrmValue & tPrmIcdData.IcdDepAra & Chr(5)
    sPrmValue = sPrmValue & tPrmIcdData.IcdUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmIcdData.IcdUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmIcdData.IcdLagCod & Chr(5)
    sPrmValue = sPrmValue & tPrmIcdData.IcdMidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmIcdData.IcdHanIcd & Chr(5)
    sPrmValue = sPrmValue & tPrmIcdData.IcdCanYon & Chr(5)
'****************************************************> 추가
    sPrmValue = sPrmValue & tPrmIcdData.IcdVeeCod & Chr(5)  '20040115..HTS..
'****************************************************> 추가

    
End Sub

    
Public Sub ImgMstLoad(sPrmValue As String, tPrmImgData As ImgMstRec)

    On Error GoTo ImgMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    
    i = i + 1
    tPrmImgData.ImgElcCod = vVal(i)
    i = i + 1
    tPrmImgData.ImgAdpDte = vVal(i)
    i = i + 1
    tPrmImgData.ImgExpDte = vVal(i)
    i = i + 1
    tPrmImgData.ImgAdpCod = vVal(i)
    i = i + 1
    tPrmImgData.ImgAdpPrc = vVal(i)
    i = i + 1
    tPrmImgData.ImgPatTyp = vVal(i)
    i = i + 1
    tPrmImgData.ImgDepCod = vVal(i)
    i = i + 1
    tPrmImgData.ImgInsCod = vVal(i)
    
    Exit Sub

ImgMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ImgMstStore(sPrmKey As String, sPrmValue As String, tPrmImgData As ImgMstRec)

    
    sPrmKey = tPrmImgData.ImgElcCod & Chr(5)
    sPrmKey = sPrmKey & tPrmImgData.ImgAdpDte & Chr(5)
    
    sPrmValue = tPrmImgData.ImgExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmImgData.ImgAdpCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImgData.ImgAdpPrc & Chr(5)
    sPrmValue = sPrmValue & tPrmImgData.ImgPatTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmImgData.ImgDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmImgData.ImgInsCod & Chr(5)
    
End Sub

    
Public Sub IncMstLoad(sPrmValue As String, tPrmIncData As IncMstRec)

    On Error GoTo IncMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmIncData.IncCod = vVal(i)
    i = i + 1
    tPrmIncData.IncCodNam = vVal(i)
    i = i + 1
    tPrmIncData.IncIprSit = vVal(i)
    i = i + 1
    tPrmIncData.IncOprSit = vVal(i)
    i = i + 1
    tPrmIncData.IncEtcDep = vVal(i)
    i = i + 1
    tPrmIncData.IncInoSit = vVal(i)
    i = i + 1
    tPrmIncData.IncOnoSit = vVal(i)
    i = i + 1
    tPrmIncData.IncTyp = vVal(i)
    i = i + 1
    tPrmIncData.IncHOpSit = vVal(i)
    i = i + 1
    tPrmIncData.IncHOnSit = vVal(i)
    i = i + 1
    tPrmIncData.IncHIpSit = vVal(i)
    i = i + 1
    tPrmIncData.IncHInSit = vVal(i)
    
    
    
    Exit Sub

IncMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub IncMstStore(sPrmKey As String, sPrmValue As String, tPrmIncData As IncMstRec)

    
    sPrmKey = Format(CDouble(tPrmIncData.IncCod), "@@") & Chr(5)
    
    sPrmValue = tPrmIncData.IncCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncIprSit & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncOprSit & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncEtcDep & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncInoSit & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncOnoSit & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncHOpSit & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncHOnSit & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncHIpSit & Chr(5)
    sPrmValue = sPrmValue & tPrmIncData.IncHInSit & Chr(5)
    
End Sub

    
Public Sub InsMstLoad(sPrmValue As String, tPrmInsData As InsMstRec)

    On Error GoTo InsMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmInsData.InsCod = vVal(i)
    i = i + 1
    tPrmInsData.InsHspTyp = vVal(i)
    i = i + 1
    tPrmInsData.InsMdcTyp = vVal(i)
    i = i + 1
    tPrmInsData.InsAdpDte = vVal(i)
    i = i + 1
    tPrmInsData.InsCodNam = vVal(i)
    i = i + 1
    tPrmInsData.InsConYon = vVal(i)
    i = i + 1
    tPrmInsData.InsFeeYon = vVal(i)
    i = i + 1
    tPrmInsData.InsFeeLvl = vVal(i)
    i = i + 1
    tPrmInsData.InsOpoRat = vVal(i)
    i = i + 1
    tPrmInsData.InsOpbRat = vVal(i)
    i = i + 1
    tPrmInsData.InsIpoRat = vVal(i)
    i = i + 1
    tPrmInsData.InsIpbRat = vVal(i)
    i = i + 1
    tPrmInsData.InsHadRat = vVal(i)
    i = i + 1
    tPrmInsData.InsLmtHig = vVal(i)
    i = i + 1
    tPrmInsData.InsLmtOwn = vVal(i)
    i = i + 1
    tPrmInsData.InsLmt70 = vVal(i)
    i = i + 1
    tPrmInsData.InsCasAmt = vVal(i)
    i = i + 1
    tPrmInsData.InsCodTyp = vVal(i)
    i = i + 1
    tPrmInsData.InsCutCod = vVal(i)
    i = i + 1
    tPrmInsData.InsNonYon = vVal(i)
    i = i + 1
    tPrmInsData.InsConCor = vVal(i)
    i = i + 1
    tPrmInsData.InsDgsOpo = vVal(i)
    i = i + 1
    tPrmInsData.InsDgsOpb = vVal(i)
    i = i + 1
    tPrmInsData.InsLmtOut = vVal(i)
    i = i + 1
    tPrmInsData.InsCasOut = vVal(i)
    i = i + 1
    tPrmInsData.InsLmtDig = vVal(i)
    i = i + 1
    tPrmInsData.InsReqCod = vVal(i)
    
    Exit Sub

InsMstLoad_ErrorTraping:
    Resume Next

End Sub
    
    '********************************************************************
    '       보험정보 찾기 (진료과 정보의 병원구분과 진료구분을 적용)
    '
    '   sPrmInsCod : 보험 코드
    '   sPrmHspTyp : 병원구분(1.의원, 2.병원, 3.종합병원, 4.대학병원)
    '   sPrmMdcTyp : 진료구분(1.의과, 2.치과, 3.정신과, 4.한방과)
    '   sPrmDate   : 적용일자 (ex. 19961231)
    '********************************************************************
    '

Public Sub InsMstRead(sPrmInsCod As String, sPrmHspTyp As String, sPrmMdcTyp As String, sPrmDate As String, InsData As InsMstRec)
    
    Dim sInsCurKey As String, sInsCmpKey As String, sInsRetVal As String
    
    sInsCmpKey = sPrmInsCod & Chr(5)
    sInsCurKey = sInsCmpKey & sPrmHspTyp & Chr(5) & sPrmMdcTyp & Chr(5) & sPrmDate & Chr(5)
    
    sInsCurKey = mSetPrev("InsMst", sInsCurKey)
    sInsCurKey = mReadPrev("InsMst", sInsCurKey, sInsCmpKey, sInsRetVal)
    
    Call InsMstLoad(sInsRetVal, InsData)
    
    Exit Sub

InsMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub InsMstStore(sPrmKey As String, sPrmValue As String, tPrmInsData As InsMstRec)
    
    sPrmKey = tPrmInsData.InsCod & Chr(5)
    sPrmKey = sPrmKey & tPrmInsData.InsHspTyp & Chr(5)
    sPrmKey = sPrmKey & tPrmInsData.InsMdcTyp & Chr(5)
    sPrmKey = sPrmKey & tPrmInsData.InsAdpDte & Chr(5)
    
    sPrmValue = tPrmInsData.InsCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsConYon & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsFeeYon & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsFeeLvl & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsOpoRat & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsOpbRat & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsIpoRat & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsIpbRat & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsHadRat & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsLmtHig & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsLmtOwn & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsLmt70 & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsCasAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsCodTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsCutCod & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsNonYon & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsConCor & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsDgsOpo & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsDgsOpb & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsLmtOut & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsCasOut & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsLmtDig & Chr(5)
    sPrmValue = sPrmValue & tPrmInsData.InsReqCod & Chr(5)
        
End Sub
    
Public Sub ItmMstLoad(sPrmValue As String, tPrmItmData As ItmMstRec)

    On Error GoTo ItmMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmItmData.ItmCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmAstCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmCodNam = vVal(i)
    i = i + 1
    tPrmItmData.ItmWrkCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmWrkYon = vVal(i)
    i = i + 1
    tPrmItmData.ItmCarCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmCarYon = vVal(i)
    i = i + 1
    tPrmItmData.ItmIncCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmGudCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmGudYon = vVal(i)
    i = i + 1
    tPrmItmData.ItmAdpDte = vVal(i)
    i = i + 1
    tPrmItmData.ItmExpDte = vVal(i)
    
    Exit Sub

ItmMstLoad_ErrorTraping:
    Resume Next

End Sub


Public Sub ItmHstLoad(sPrmValue As String, tPrmItmData As ItmHstRec)

    On Error GoTo ItmHstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmItmData.ItmCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmAstCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmAdpKey = vVal(i)
    i = i + 1
    tPrmItmData.ItmCodNam = vVal(i)
    i = i + 1
    tPrmItmData.ItmWrkCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmWrkYon = vVal(i)
    i = i + 1
    tPrmItmData.ItmCarCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmCarYon = vVal(i)
    i = i + 1
    tPrmItmData.ItmIncCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmGudCod = vVal(i)
    i = i + 1
    tPrmItmData.ItmGudYon = vVal(i)
    i = i + 1
    tPrmItmData.ItmAdpDte = vVal(i)
    i = i + 1
    tPrmItmData.ItmExpDte = vVal(i)
    Exit Sub

ItmHstLoad_ErrorTraping:
    Resume Next

End Sub

    
Public Sub ItmMstStore(sPrmKey As String, sPrmValue As String, tPrmItmData As ItmMstRec)

    
    sPrmKey = tPrmItmData.ItmCod & Chr(5)
    sPrmKey = sPrmKey & tPrmItmData.ItmAstCod & Chr(5)
    
    sPrmValue = tPrmItmData.ItmCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmWrkCod & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmWrkYon & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmCarCod & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmCarYon & Chr(5)
    'sPrmValue = sPrmValue & Format(CDouble(tPrmItmData.ItmIncCod), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmIncCod & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmGudCod & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmGudYon & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmExpDte & Chr(5)
    
End Sub


Public Sub ItmHstStore(sPrmKey As String, sPrmValue As String, tPrmItmData As ItmHstRec)

    
    sPrmKey = tPrmItmData.ItmCod & Chr(5)
    sPrmKey = sPrmKey & tPrmItmData.ItmAstCod & Chr(5)
    sPrmKey = sPrmKey & tPrmItmData.ItmAdpKey & Chr(5)
    
    sPrmValue = tPrmItmData.ItmCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmWrkCod & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmWrkYon & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmCarCod & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmCarYon & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmItmData.ItmIncCod), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmGudCod & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmGudYon & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmItmData.ItmExpDte & Chr(5)
    
End Sub

    
Public Sub LabMstLoad(sPrmValue As String, tPrmLabData As LabMstRec)

    On Error GoTo LabMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmLabData.LabCod = vVal(i)
    i = i + 1
    tPrmLabData.LabSeq = vVal(i)
    i = i + 1
    tPrmLabData.LabSubCod = vVal(i)
    i = i + 1
    tPrmLabData.LabCodTyp = vVal(i)
    i = i + 1
    tPrmLabData.LabCodNam = vVal(i)
    i = i + 1
    tPrmLabData.LabTubCod = vVal(i)
    i = i + 1
    tPrmLabData.LabSpmCod = vVal(i)
    i = i + 1
    tPrmLabData.LabSpmReq = vVal(i)
    i = i + 1
    tPrmLabData.LabComMax = vVal(i)
    i = i + 1
    tPrmLabData.LabComLow = vVal(i)
    i = i + 1
    tPrmLabData.LabMalMax = vVal(i)
    i = i + 1
    tPrmLabData.LabMalLow = vVal(i)
    i = i + 1
    tPrmLabData.LabFmlMax = vVal(i)
    i = i + 1
    tPrmLabData.LabFmlLow = vVal(i)
    i = i + 1
    tPrmLabData.LabMzhUnt = vVal(i)
    i = i + 1
    tPrmLabData.LabSclYon = vVal(i)
    
    Exit Sub

LabMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub LabMstStore(sPrmKey As String, sPrmValue As String, tPrmLabData As LabMstRec)

    
    sPrmKey = tPrmLabData.LabCod & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmLabData.LabSeq), "@@") & Chr(5)
    
    sPrmValue = tPrmLabData.LabSubCod & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabCodTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabTubCod & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabSpmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabSpmReq & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabComMax & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabComLow & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabMalMax & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabMalLow & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabFmlMax & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabFmlLow & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabMzhUnt & Chr(5)
    sPrmValue = sPrmValue & tPrmLabData.LabSclYon & Chr(5)
    
End Sub

    
Public Sub MgdMstLoad(sPrmValue As String, tPrmData As MgdMstRec)

    On Error GoTo MgdMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.MgdCod = vVal(i)
    i = i + 1
    tPrmData.MgdAdpDte = vVal(i)
    i = i + 1
    tPrmData.MgdNam = vVal(i)
    i = i + 1
    tPrmData.MgdInsCod = vVal(i)
    i = i + 1
    tPrmData.MgdNatCod = vVal(i)
    i = i + 1
    tPrmData.MgdCarCod = vVal(i)
    i = i + 1
    tPrmData.MgdWrkCod = vVal(i)
    i = i + 1
    tPrmData.MgdGenCod = vVal(i)
    i = i + 1
    tPrmData.MgdDifCod = vVal(i)
    i = i + 1
    tPrmData.MgdMgdSeq = vVal(i)
    'eversky 모양이 들어간다.
    i = i + 1
    tPrmData.MgdMgdPic = vVal(i)
    i = i + 1
    tPrmData.MgdCruCod = vVal(i)    '직원식대코드
    i = i + 1
    tPrmData.MgdSecCod = vVal(i)    '3월1일보호1종식이변경처리
    Exit Sub

MgdMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub MgdMstStore(sPrmKey As String, sPrmValue As String, tPrmMgdData As MgdMstRec)

    
    sPrmKey = tPrmMgdData.MgdCod & Chr(5)
    sPrmKey = sPrmKey & tPrmMgdData.MgdAdpDte & Chr(5)
    
    sPrmValue = tPrmMgdData.MgdNam & Chr(5)
    sPrmValue = sPrmValue & tPrmMgdData.MgdInsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmMgdData.MgdNatCod & Chr(5)
    sPrmValue = sPrmValue & tPrmMgdData.MgdCarCod & Chr(5)
    sPrmValue = sPrmValue & tPrmMgdData.MgdWrkCod & Chr(5)
    sPrmValue = sPrmValue & tPrmMgdData.MgdGenCod & Chr(5)
    sPrmValue = sPrmValue & tPrmMgdData.MgdDifCod & Chr(5)
    sPrmValue = sPrmValue & Format(Trim(tPrmMgdData.MgdMgdSeq), "@@@") & Chr(5)
    'eversky 모양 추가로 새로운 항목 추가
    sPrmValue = sPrmValue & tPrmMgdData.MgdMgdPic & Chr(5)
    sPrmValue = sPrmValue & tPrmMgdData.MgdCruCod & Chr(5)  '직원식대 코드
    sPrmValue = sPrmValue & tPrmMgdData.MgdSecCod & Chr(5)  '3월1일보호1종식이변경처리
    
End Sub

    
Public Sub MsgMstLoad(sPrmValue As String, tPrmMgdData As MsgMstRec)

    On Error GoTo IcdMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmMgdData.MsgCod = vVal(i)
    i = i + 1
    tPrmMgdData.MsgCodNam = vVal(i)
    
    Exit Sub

IcdMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub MsgMstStore(sPrmKey As String, sPrmValue As String, tPrmMgdData As MsgMstRec)

    
    sPrmKey = tPrmMgdData.MsgCod & Chr(5)
    sPrmValue = tPrmMgdData.MsgCodNam & Chr(5)
    
End Sub

    
Public Sub MthMstLoad(sPrmValue As String, tPrmData As MthMstRec)

    On Error GoTo MthMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.MthCod = vVal(i)
    i = i + 1
    tPrmData.MthAdpDte = vVal(i)
    i = i + 1
    tPrmData.MthNam = vVal(i)
    i = i + 1
    tPrmData.MthFeeCod = vVal(i)
    i = i + 1
    tPrmData.MthInpLmt = vVal(i)
    i = i + 1
    tPrmData.MthOutLmt = vVal(i)
    i = i + 1
    tPrmData.MthInpOvr = vVal(i)
    i = i + 1
    tPrmData.MthOutOvr = vVal(i)
    '수기료수정 @:^)
    i = i + 1
    tPrmData.MthCinLmt = vVal(i)
    i = i + 1
    tPrmData.MthCouLmt = vVal(i)
    i = i + 1
    tPrmData.MthCinOvr = vVal(i)
    i = i + 1
    tPrmData.MthCouOvr = vVal(i)
    i = i + 1
    tPrmData.MthSinLmt = vVal(i)
    i = i + 1
    tPrmData.MthSouLmt = vVal(i)
    i = i + 1
    tPrmData.MthSinOvr = vVal(i)
    i = i + 1
    tPrmData.MthSouOvr = vVal(i)
    i = i + 1
    tPrmData.MthBinLmt = vVal(i)
    i = i + 1
    tPrmData.MthBouLmt = vVal(i)
    i = i + 1
    tPrmData.MthBinOvr = vVal(i)
    i = i + 1
    tPrmData.MthBouOvr = vVal(i)
    '/수기료수정 @:^)
    
    Exit Sub

MthMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub MthMstStore(sPrmKey As String, sPrmValue As String, tPrmMthData As MthMstRec)

    
    sPrmKey = tPrmMthData.MthCod & Chr(5)
    sPrmKey = sPrmKey & tPrmMthData.MthAdpDte & Chr(5)
    
    sPrmValue = tPrmMthData.MthNam & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthFeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthInpLmt & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthOutLmt & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthInpOvr & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthOutOvr & Chr(5)
    
    sPrmValue = sPrmValue & tPrmMthData.MthCinLmt & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthCouLmt & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthCinOvr & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthCouOvr & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthSinLmt & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthSouLmt & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthSinOvr & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthSouOvr & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthBinLmt & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthBouLmt & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthBinOvr & Chr(5)
    sPrmValue = sPrmValue & tPrmMthData.MthBouOvr & Chr(5)
    
    
End Sub

    
    
Public Sub OpiMstLoad(sPrmValue As String, tPrmOpiData As OpiMstRec)

    On Error GoTo OpiMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmOpiData.OpiDepCod = vVal(i)
    i = i + 1
    tPrmOpiData.OpiCod = vVal(i)
    i = i + 1
    tPrmOpiData.OpiCodDsc = vVal(i)
    
    Exit Sub

OpiMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OpiMstStore(sPrmKey As String, sPrmValue As String, tPrmOpiData As OpiMstRec)

    
    sPrmKey = tPrmOpiData.OpiDepCod & Chr(5)
    sPrmKey = sPrmKey & tPrmOpiData.OpiCod & Chr(5)
    
    sPrmValue = tPrmOpiData.OpiCodDsc & Chr(5)
End Sub

    
Public Sub OprMstLoad(sPrmValue As String, tPrmOprData As OprMstRec)

    On Error GoTo OprMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmOprData.OprDepCod = vVal(i)
    i = i + 1
    tPrmOprData.OprCod = vVal(i)
    i = i + 1
    tPrmOprData.OprCodDsc = vVal(i)
    i = i + 1
    tPrmOprData.OprSeeCod = vVal(i)
    i = i + 1
    tPrmOprData.OprUseTms = vVal(i)
    i = i + 1
    tPrmOprData.OprIcdYon = vVal(i)
    
    
    Exit Sub

OprMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub OprMstStore(sPrmKey As String, sPrmValue As String, tPrmOprData As OprMstRec)

    
    sPrmKey = tPrmOprData.OprDepCod & Chr(5)
    sPrmKey = sPrmKey & tPrmOprData.OprCod & Chr(5)
    
    sPrmValue = tPrmOprData.OprCodDsc & Chr(5)
    sPrmValue = sPrmValue & tPrmOprData.OprSeeCod & Chr(5)
    sPrmValue = sPrmValue & tPrmOprData.OprUseTms & Chr(5)
    sPrmValue = sPrmValue & tPrmOprData.OprIcdYon & Chr(5)
    
End Sub

    
Public Sub RomMstLoad(sPrmValue As String, tPrmRomData As RomMstRec)

    On Error GoTo RomMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmRomData.RomWrdCod = vVal(i)
    i = i + 1
    tPrmRomData.RomCod = vVal(i)
    i = i + 1
    tPrmRomData.RomCodNam = vVal(i)
    i = i + 1
    tPrmRomData.RomDepCod = vVal(i)
    i = i + 1
    tPrmRomData.RomBasBed = vVal(i)
    i = i + 1
    tPrmRomData.RomActBed = vVal(i)
    i = i + 1
    tPrmRomData.RomRemBed = vVal(i)
    i = i + 1
    tPrmRomData.RomSexCod = vVal(i)
    i = i + 1
    tPrmRomData.RomTyp = vVal(i)
    i = i + 1
    tPrmRomData.RomGrdCod = vVal(i)
    i = i + 1
    tPrmRomData.RomStsCod = vVal(i)
    i = i + 1
    tPrmRomData.RomEqpInf = vVal(i)
    
    Exit Sub

RomMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub RomMstStore(sPrmKey As String, sPrmValue As String, tPrmRomData As RomMstRec)

    
    sPrmKey = tPrmRomData.RomWrdCod & Chr(5)
    sPrmKey = sPrmKey & tPrmRomData.RomCod & Chr(5)
    
    sPrmValue = tPrmRomData.RomCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomBasBed & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomActBed & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomRemBed & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomSexCod & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomGrdCod & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomStsCod & Chr(5)
    sPrmValue = sPrmValue & tPrmRomData.RomEqpInf & Chr(5)
    
End Sub

    
Public Sub RptMstLoad(sPrmValue As String, tPrmData As RptMstRec)

    On Error GoTo RptMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmData.RptJobTyp = vVal(i)
    i = i + 1
    tPrmData.RptSeqNum = vVal(i)
    i = i + 1
    tPrmData.RptTypFlg = vVal(i)
    i = i + 1
    tPrmData.RptNam = vVal(i)
    i = i + 1
    tPrmData.RptExeNam = vVal(i)
    i = i + 1
    tPrmData.RptSgnCnt = vVal(i)
    i = i + 1
    tPrmData.RptSgnNam = vVal(i)
    
    Exit Sub

RptMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub RptMstStore(sPrmKey As String, sPrmValue As String, tPrmData As RptMstRec)

    
    sPrmKey = tPrmData.RptJobTyp & Chr(5)
    
    
    sPrmKey = sPrmKey & tPrmData.RptSeqNum & Chr(5)
    
    sPrmValue = tPrmData.RptTypFlg & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RptNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RptExeNam & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RptSgnCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmData.RptSgnNam & Chr(5)
    
End Sub

    
Public Sub SeeHstLoad(sPrmValue As String, tPrmSeeData As SeeHstRec)

    On Error GoTo SeeHstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmSeeData.SeeOdrCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAdpKey = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSotCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSlpDep = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSlpCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCodTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeEngNam = vVal(i)
    i = i + 1
    tPrmSeeData.SeeKorNam = vVal(i)
    i = i + 1
    tPrmSeeData.SeeElcCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeItmCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAstCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeePhrTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSlpTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSclRat = vVal(i)
    i = i + 1
    tPrmSeeData.SeeDivYon = vVal(i)
    i = i + 1
    tPrmSeeData.SeeDrgCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUsgCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeMthCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeRepYon = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAddCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCalTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUntQty = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUntCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSpmCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeMakCmp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSpcAmt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeLftCnt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAdpDte = vVal(i)
    i = i + 1
    tPrmSeeData.SeeExpDte = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUidCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUpdDtm = vVal(i)
    i = i + 1
    tPrmSeeData.SeeComNam = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAddNon = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCodDiv = vVal(i)
    i = i + 1
    tPrmSeeData.SeeRelCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAdmCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSotTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeTotQty = vVal(i)
    i = i + 1
    tPrmSeeData.SeeTotTms = vVal(i)
    i = i + 1
    tPrmSeeData.SeeEffect = vVal(i)
    i = i + 1
    tPrmSeeData.SeeInsAmt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCarAmt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeWrkAmt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCodYon = vVal(i)
    
    
    
    Exit Sub

SeeHstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub SeeHstStore(sPrmKey As String, sPrmValue As String, tPrmSeeData As SeeHstRec)

    
    sPrmKey = tPrmSeeData.SeeOdrCod & Chr(5)
    sPrmKey = sPrmKey & tPrmSeeData.SeeAdpKey & Chr(5)
    
    sPrmValue = tPrmSeeData.SeeSotCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSlpDep & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSlpCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCodTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeEngNam & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeKorNam & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeElcCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeItmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAstCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeePhrTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSlpTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSclRat & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeDivYon & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeDrgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUsgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeMthCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeRepYon & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAddCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCalTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUntQty & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUntCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSpmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeMakCmp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSpcAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeLftCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeComNam & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAddNon & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCodDiv & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeRelCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAdmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSotTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeTotQty & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeTotTms & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeEffect & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCarAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeWrkAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCodYon & Chr(5)
    
    
End Sub

    
Public Sub SeeMstLoad(sPrmValue As String, tPrmSeeData As SeeMstRec)

    On Error GoTo SeeMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    i = i + 1
    tPrmSeeData.SeeOdrCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSotCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSlpDep = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSlpCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCodTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeEngNam = vVal(i)
    i = i + 1
    tPrmSeeData.SeeKorNam = vVal(i)
    i = i + 1
    tPrmSeeData.SeeElcCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeItmCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAstCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeePhrTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSlpTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSclRat = vVal(i)
    i = i + 1
    tPrmSeeData.SeeDivYon = vVal(i)
    i = i + 1
    tPrmSeeData.SeeDrgCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUsgCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeMthCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeRepYon = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAddCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCalTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUntQty = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUntCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSpmCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeMakCmp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSpcAmt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeLftCnt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAdpDte = vVal(i)
    i = i + 1
    tPrmSeeData.SeeExpDte = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUidCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeUpdDtm = vVal(i)
    i = i + 1
    tPrmSeeData.SeeComNam = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAddNon = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCodDiv = vVal(i)
    i = i + 1
    tPrmSeeData.SeeRelCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeAdmCod = vVal(i)
    i = i + 1
    tPrmSeeData.SeeSotTyp = vVal(i)
    i = i + 1
    tPrmSeeData.SeeTotQty = vVal(i)
    i = i + 1
    tPrmSeeData.SeeTotTms = vVal(i)
    i = i + 1
    tPrmSeeData.SeeEffect = vVal(i)
    i = i + 1
    tPrmSeeData.SeeInsAmt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCarAmt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeWrkAmt = vVal(i)
    i = i + 1
    tPrmSeeData.SeeCodYon = vVal(i)
    
    Exit Sub

SeeMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub SeeMstRead(sPrmCod As String, tPrmSeeMst As SeeMstRec)
    
    Dim sCurKey As String
    Dim sRetVal As String
    
    sCurKey = sPrmCod & Chr(5)
    sCurKey = mSetReadEqual("SeeMst", sCurKey, sRetVal)
    Call SeeMstLoad(sRetVal, tPrmSeeMst)
    
    Exit Sub

SeeMstLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub SeeMstReadByAdpDte(sPrmSeeCod As String, sPrmDate As String, SeeData As SeeMstRec)

    Dim SeeHstData As SeeHstRec
    Dim sSeeMstCurKey As String, sSeeMstCmpKey As String, sSeeMstRetVal As String

    sSeeMstCmpKey = sPrmSeeCod & Chr(5)
    sSeeMstCurKey = sSeeMstCmpKey
    sSeeMstCurKey = mSetReadNext("SeeMst", sSeeMstCurKey, sSeeMstCmpKey, sSeeMstRetVal)
    
    If sSeeMstCurKey = "" Then Exit Sub
    
    Call SeeMstLoad(sSeeMstRetVal, SeeData)
    
    '------------------------------------------------------------
    '- 적용일에 부합되면 히스토리를 읽을 필요 없이 Exit 한다.
    '------------------------------------------------------------
    If Left(sPrmDate, 8) >= Left((SeeData.SeeAdpDte), 8) And Left(sPrmDate, 8) <= Left((SeeData.SeeExpDte), 8) Then
        Exit Sub
    End If
    
    If Trim(SeeData.SeeRelCod) <> "" Then
        MsgBox SeeData.SeeKorNam & "(" & SeeData.SeeOdrCod & ")는 " & _
               SeeData.SeeRelCod & "로 대체 되었습니다."
    End If
    
    '-------------------------------------------------------
    '- 적용일범위를 벗어나면 수가정보의 History를 읽는다.
    '-------------------------------------------------------
    sSeeMstCmpKey = sPrmSeeCod & Chr(5)
    sSeeMstCurKey = sSeeMstCmpKey & sPrmDate & Chr(5)
    sSeeMstCurKey = mSetPrev("SeeHst", sSeeMstCurKey)
    sSeeMstCurKey = mReadPrev("SeeHst", sSeeMstCurKey, sSeeMstCmpKey, sSeeMstRetVal)
            
    'Bug가 맞는데 일단은 그냥 둔다.
    'If sSeeMstCurKey = "" Then Exit Sub
    
    Call SeeHstLoad(sSeeMstRetVal, SeeHstData)
    Call SeeHstStore(sSeeMstCurKey, sSeeMstRetVal, SeeHstData)

    '------------------------------------------------------
    '- History도 적용일에 부합되는지 check한다(970918)
    '------------------------------------------------------
    If Left(sPrmDate, 8) >= Left((SeeHstData.SeeAdpDte), 8) And Left(sPrmDate, 8) <= Left((SeeHstData.SeeExpDte), 8) Then
        sSeeMstRetVal = sPrmSeeCod & Chr(5) & sSeeMstRetVal
        Call SeeMstLoad(sSeeMstRetVal, SeeData)
    Else
        sSeeMstRetVal = ""
        Call SeeMstLoad(sSeeMstRetVal, SeeData)
    End If

End Sub
    
Public Sub SeeMstStore(sPrmKey As String, sPrmValue As String, tPrmSeeData As SeeMstRec)

    
    sPrmKey = tPrmSeeData.SeeOdrCod & Chr(5)
    
    sPrmValue = tPrmSeeData.SeeSotCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSlpDep & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSlpCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCodTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeEngNam & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeKorNam & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeElcCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeItmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAstCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeePhrTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSlpTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSclRat & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeDivYon & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeDrgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUsgCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeMthCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeRepYon & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAddCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCalTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUntQty & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUntCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSpmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeMakCmp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSpcAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeLftCnt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeExpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUidCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeUpdDtm & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeComNam & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAddNon & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCodDiv & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeRelCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeAdmCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeSotTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeTotQty & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeTotTms & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeEffect & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeInsAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCarAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeWrkAmt & Chr(5)
    sPrmValue = sPrmValue & tPrmSeeData.SeeCodYon & Chr(5)
    
    
End Sub

    
Public Sub SimMstLoad(sPrmValue As String, tPrmSimData As SimMstRec)

    On Error GoTo SimMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmSimData.SimOdrCod = vVal(i)
    i = i + 1
    tPrmSimData.SimRepCod = vVal(i)
    i = i + 1
    tPrmSimData.SimLowQty = vVal(i)
    i = i + 1
    tPrmSimData.SimHigQty = vVal(i)
    i = i + 1
    tPrmSimData.SimAvgQty = vVal(i)
    i = i + 1
    tPrmSimData.SimRefCmd = vVal(i)
    i = i + 1
    tPrmSimData.SimIcdCod = vVal(i)
    
    
    Exit Sub

SimMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub SimMstStore(sPrmKey As String, sPrmValue As String, tPrmSimData As SimMstRec)

    
    sPrmKey = tPrmSimData.SimOdrCod & Chr(5)
    
    sPrmValue = tPrmSimData.SimRepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSimData.SimLowQty & Chr(5)
    sPrmValue = sPrmValue & tPrmSimData.SimHigQty & Chr(5)
    sPrmValue = sPrmValue & tPrmSimData.SimAvgQty & Chr(5)
    sPrmValue = sPrmValue & tPrmSimData.SimRefCmd & Chr(5)
    sPrmValue = sPrmValue & tPrmSimData.SimIcdCod & Chr(5)
    
End Sub

    
    
Public Sub SxyMstLoad(sPrmValue As String, tPrmSxyData As SxyMstRec)

    On Error GoTo SxyMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmSxyData.SxyElcCod = vVal(i)
    i = i + 1
    tPrmSxyData.SxyOdrSeq = vVal(i)
    i = i + 1
    tPrmSxyData.SxyOdrCod = vVal(i)
    i = i + 1
    tPrmSxyData.SxyOdrQty = vVal(i)
    
    Exit Sub

SxyMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub SxyMstStore(sPrmKey As String, sPrmValue As String, tPrmSxyData As SxyMstRec)

    
    sPrmKey = tPrmSxyData.SxyElcCod & Chr(5)
    sPrmKey = sPrmKey & Format(tPrmSxyData.SxyOdrSeq, "@@") & Chr(5)
    
    sPrmValue = tPrmSxyData.SxyOdrCod & Chr(5)
    sPrmValue = sPrmValue & tPrmSxyData.SxyOdrQty & Chr(5)
    
End Sub

    
Public Sub TabMstLoad(sPrmValue As String, tPrmTabData As TabMstRec)

    On Error GoTo TabMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmTabData.TabCod = vVal(i)
    i = i + 1
    tPrmTabData.TabCodNam = vVal(i)
    i = i + 1
    tPrmTabData.TabUpdYon = vVal(i)
    
    Exit Sub

TabMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub TabMstStore(sPrmKey As String, sPrmValue As String, tPrmTabData As TabMstRec)

    
    sPrmKey = tPrmTabData.TabCod & Chr(5)
    
    sPrmValue = tPrmTabData.TabCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmTabData.TabUpdYon & Chr(5)
    
End Sub

    
Public Sub UidMstLoad(sPrmValue As String, tPrmUidData As UidMstRec)

    On Error GoTo UidMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1

'    i = i + 1
'    tPrmUidData.UidCod = vVal(i)
'    i = i + 1
'    tPrmUidData.UidNam = vVal(i)
'    i = i + 1
'    tPrmUidData.UidPwd = vVal(i)
'    i = i + 1
'    tPrmUidData.UidDepCod = vVal(i)
'    i = i + 1
'    tPrmUidData.UidSecLev = vVal(i)
'    i = i + 1
'    tPrmUidData.UidEmpNum = vVal(i)
'    i = i + 1
'    tPrmUidData.UidPrtCod = vVal(i)
'    i = i + 1
'    tPrmUidData.UiddtrYon = vVal(i)
'    i = i + 1
'    tPrmUidData.UidSpcYon = vVal(i)
'    i = i + 1
'    tPrmUidData.UidAssLev = vVal(i)
'    i = i + 1
'    tPrmUidData.UidPosCod = vVal(i)
'    i = i + 1
'    tPrmUidData.UidSgnDir = vVal(i)
'    i = i + 1
'    tPrmUidData.UidSgnFle = vVal(i)
'    i = i + 1
'    tPrmUidData.UidLicNum = vVal(i)
'    i = i + 1
'    tPrmUidData.UidTelNum = vVal(i)
'    i = i + 1
'    tPrmUidData.UidMalAdd = vVal(i)
'    i = i + 1
'    tPrmUidData.UidAdpDte = vVal(i)
'    i = i + 1
'    tPrmUidData.UidExpDte = vVal(i)

    i = i + 1
    tPrmUidData.UidCod = vVal(i)
    i = i + 1
    tPrmUidData.UidNam = vVal(i)
    i = i + 1
    tPrmUidData.UidPwd = vVal(i)
    i = i + 1
    tPrmUidData.UidDepCod = vVal(i)
    i = i + 1
    tPrmUidData.UidSecLev = vVal(i)
    i = i + 1
    tPrmUidData.UidEmpNum = vVal(i)
    i = i + 1
    tPrmUidData.UidPrtCod = vVal(i)
    i = i + 1
    tPrmUidData.UidDtrYon = vVal(i)
    i = i + 1
    tPrmUidData.UidSpcYon = vVal(i)
    i = i + 1
    tPrmUidData.UidAssLev = vVal(i)
    i = i + 1
    tPrmUidData.UidPosCod = vVal(i)
    i = i + 1
    tPrmUidData.UidSgnDir = vVal(i)
    i = i + 1
    tPrmUidData.UidSgnFle = vVal(i)
    i = i + 1
    tPrmUidData.UidLicNum = vVal(i)
    i = i + 1
    tPrmUidData.UidTelNum = vVal(i)
    i = i + 1
    tPrmUidData.UidMalAdd = vVal(i)
    i = i + 1
    tPrmUidData.UidAdpDte = vVal(i)
    i = i + 1
    tPrmUidData.UidEndDte = vVal(i)
    i = i + 1
    tPrmUidData.UidSpcNum = vVal(i)


    Exit Sub

UidMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub UidMstStore(sPrmKey As String, sPrmValue As String, tPrmUidData As UidMstRec)

    
    sPrmKey = tPrmUidData.UidCod & Chr(5)
    
'    sPrmValue = tPrmUidData.UidNam & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidPwd & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidDepCod & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidSecLev & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidEmpNum & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidPrtCod & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UiddtrYon & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidSpcYon & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidAssLev & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidPosCod & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidSgnDir & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidSgnFle & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidLicNum & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidTelNum & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidMalAdd & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidAdpDte & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidExpDte & Chr(5)

    sPrmValue = tPrmUidData.UidNam & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidPwd & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSecLev & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidEmpNum & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidPrtCod & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidDtrYon & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidAssLev & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidPosCod & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSgnDir & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSgnFle & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidLicNum & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidTelNum & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidMalAdd & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidEndDte & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSpcNum & Chr(5)
    
End Sub

Public Sub UidHstLoad(sPrmValue As String, tPrmUidData As UidHstRec)

    On Error GoTo UidHstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
'    i = i + 1
'    tPrmUidData.UidCod = vVal(i)
'    i = i + 1
'    tPrmUidData.UidAdpKey = vVal(i)
'    i = i + 1
'    tPrmUidData.UidNam = vVal(i)
'    i = i + 1
'    tPrmUidData.UidPwd = vVal(i)
'    i = i + 1
'    tPrmUidData.UidDepCod = vVal(i)
'    i = i + 1
'    tPrmUidData.UidSecLev = vVal(i)
'    i = i + 1
'    tPrmUidData.UidEmpNum = vVal(i)
'    i = i + 1
'    tPrmUidData.UidPrtCod = vVal(i)
'    i = i + 1
'    tPrmUidData.UiddtrYon = vVal(i)
'    i = i + 1
'    tPrmUidData.UidSpcYon = vVal(i)
'    i = i + 1
'    tPrmUidData.UidAssLev = vVal(i)
'    i = i + 1
'    tPrmUidData.UidPosCod = vVal(i)
'    i = i + 1
'    tPrmUidData.UidSgnDir = vVal(i)
'    i = i + 1
'    tPrmUidData.UidSgnFle = vVal(i)
'    i = i + 1
'    tPrmUidData.UidLicNum = vVal(i)
'    i = i + 1
'    tPrmUidData.UidTelNum = vVal(i)
'    i = i + 1
'    tPrmUidData.UidMalAdd = vVal(i)
'    i = i + 1
'    tPrmUidData.UidAdpDte = vVal(i)
'    i = i + 1
'    tPrmUidData.UidExpDte = vVal(i)

    i = i + 1
    tPrmUidData.UidCod = vVal(i)
    i = i + 1
    tPrmUidData.UidAdpKey = vVal(i)
    i = i + 1
    tPrmUidData.UidNam = vVal(i)
    i = i + 1
    tPrmUidData.UidPwd = vVal(i)
    i = i + 1
    tPrmUidData.UidDepCod = vVal(i)
    i = i + 1
    tPrmUidData.UidSecLev = vVal(i)
    i = i + 1
    tPrmUidData.UidEmpNum = vVal(i)
    i = i + 1
    tPrmUidData.UidPrtCod = vVal(i)
    i = i + 1
    tPrmUidData.UidDtrYon = vVal(i)
    i = i + 1
    tPrmUidData.UidSpcYon = vVal(i)
    i = i + 1
    tPrmUidData.UidAssLev = vVal(i)
    i = i + 1
    tPrmUidData.UidPosCod = vVal(i)
    i = i + 1
    tPrmUidData.UidSgnDir = vVal(i)
    i = i + 1
    tPrmUidData.UidSgnFle = vVal(i)
    i = i + 1
    tPrmUidData.UidLicNum = vVal(i)
    i = i + 1
    tPrmUidData.UidTelNum = vVal(i)
    i = i + 1
    tPrmUidData.UidMalAdd = vVal(i)
    i = i + 1
    tPrmUidData.UidAdpDte = vVal(i)
    i = i + 1
    tPrmUidData.UidEndDte = vVal(i)
    i = i + 1
    tPrmUidData.UidSpcNum = vVal(i)
    
    
    Exit Sub

UidHstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub UidHstStore(sPrmKey As String, sPrmValue As String, tPrmUidData As UidHstRec)

    
    sPrmKey = tPrmUidData.UidCod & Chr(5) _
            & tPrmUidData.UidAdpKey & Chr(5)
    
'    sPrmValue = tPrmUidData.UidNam & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidPwd & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidDepCod & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidSecLev & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidEmpNum & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidPrtCod & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UiddtrYon & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidSpcYon & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidAssLev & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidPosCod & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidSgnDir & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidSgnFle & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidLicNum & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidTelNum & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidMalAdd & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidAdpDte & Chr(5)
'    sPrmValue = sPrmValue & tPrmUidData.UidExpDte & Chr(5)
    
    
    sPrmValue = tPrmUidData.UidNam & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidPwd & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidDepCod & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSecLev & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidEmpNum & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidPrtCod & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidDtrYon & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSpcYon & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidAssLev & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidPosCod & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSgnDir & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSgnFle & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidLicNum & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidTelNum & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidMalAdd & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidAdpDte & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidEndDte & Chr(5)
    sPrmValue = sPrmValue & tPrmUidData.UidSpcNum & Chr(5)
End Sub

    
Public Sub UsgMstLoad(sPrmValue As String, tPrmUsgData As UsgMstRec)

    On Error GoTo UsgMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmUsgData.UsgCod = vVal(i)
    i = i + 1
    tPrmUsgData.UsgFulDsc = vVal(i)
    i = i + 1
    tPrmUsgData.UsgCodNam = vVal(i)
    i = i + 1
    tPrmUsgData.UsgOdrTms = vVal(i)
    i = i + 1
    tPrmUsgData.UsgMthCod = vVal(i)
    i = i + 1
    tPrmUsgData.UsgDspSeq = vVal(i)
    i = i + 1
    tPrmUsgData.UsgDspGrp = vVal(i)
    
    i = i + 1
    tPrmUsgData.UsgActTim = vVal(i)
    
    '20030214 lek add for 횟수에 대한 주요 용법임을 보여줌
    i = i + 1
    tPrmUsgData.UsgMainYon = vVal(i)
    
Exit Sub

UsgMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub UsgMstStore(sPrmKey As String, sPrmValue As String, tPrmUsgData As UsgMstRec)

    
    sPrmKey = tPrmUsgData.UsgCod & Chr(5)
    
    sPrmValue = tPrmUsgData.UsgFulDsc & Chr(5)
    sPrmValue = sPrmValue & tPrmUsgData.UsgCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmUsgData.UsgOdrTms & Chr(5)
    sPrmValue = sPrmValue & tPrmUsgData.UsgMthCod & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmUsgData.UsgDspSeq), "@@@@") & Chr(5)
    sPrmValue = sPrmValue & Format(CDouble(tPrmUsgData.UsgDspGrp), "@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmUsgData.UsgActTim & Chr(5)
    
    sPrmValue = sPrmValue & tPrmUsgData.UsgMainYon & Chr(5) '20030214 lek add for 횟수에 대한 기본 용법 표시
    
End Sub
    
    
Public Sub WmnMstLoad(sPrmValue As String, tPrmWmnData As wmnMstRec)

    On Error GoTo WmnMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmWmnData.WmnWrdCod = vVal(i)
    i = i + 1
    tPrmWmnData.WmnSexTyp = vVal(i)
    i = i + 1
    tPrmWmnData.WmnManTot = vVal(i)
    i = i + 1
    tPrmWmnData.WmnDepTyp = vVal(i)
    i = i + 1
    tPrmWmnData.WmnWrdTyp = vVal(i)
    
    Exit Sub

WmnMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub WmnMstStore(sPrmKey As String, sPrmValue As String, tPrmWmnData As wmnMstRec)

    
    sPrmKey = tPrmWmnData.WmnWrdCod & Chr(5)
    
    sPrmValue = tPrmWmnData.WmnSexTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmWmnData.WmnManTot & Chr(5)
    sPrmValue = sPrmValue & tPrmWmnData.WmnDepTyp & Chr(5)
    sPrmValue = sPrmValue & tPrmWmnData.WmnWrdTyp & Chr(5)
    
End Sub

    
Public Sub WrdMstLoad(sPrmValue As String, tPrmCstData As WrdMstRec)

    On Error GoTo WrdMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmCstData.WrdCod = vVal(i)
    i = i + 1
    tPrmCstData.WrdCodNam = vVal(i)
    i = i + 1
    tPrmCstData.WrdAsgBed = vVal(i)
    i = i + 1
    tPrmCstData.WrdAprBed = vVal(i)
    i = i + 1
    tPrmCstData.WrdOcpBed = vVal(i)
    i = i + 1
    tPrmCstData.WrdMonDay = vVal(i)
    i = i + 1
    tPrmCstData.WrdAnnDay = vVal(i)
    i = i + 1
    tPrmCstData.WrdSnsDte = vVal(i)
    i = i + 1
    tPrmCstData.WrdBasInf = vVal(i)
    
    Exit Sub

WrdMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub WrdMstStore(sPrmKey As String, sPrmValue As String, tPrmCstData As WrdMstRec)

    
    sPrmKey = tPrmCstData.WrdCod & Chr(5)
    
    sPrmValue = tPrmCstData.WrdCodNam & Chr(5)
    sPrmValue = sPrmValue & tPrmCstData.WrdAsgBed & Chr(5)
    sPrmValue = sPrmValue & tPrmCstData.WrdAprBed & Chr(5)
    sPrmValue = sPrmValue & tPrmCstData.WrdOcpBed & Chr(5)
    sPrmValue = sPrmValue & tPrmCstData.WrdMonDay & Chr(5)
    sPrmValue = sPrmValue & tPrmCstData.WrdAnnDay & Chr(5)
    sPrmValue = sPrmValue & tPrmCstData.WrdSnsDte & Chr(5)
    sPrmValue = sPrmValue & tPrmCstData.WrdBasInf & Chr(5)
    
End Sub

    
Public Sub ZipMstLoad(sPrmValue As String, tPrmZipData As ZipMstRec)

    On Error GoTo ZipMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    
    i = i + 1
    tPrmZipData.ZipCod = vVal(i)
    i = i + 1
    tPrmZipData.ZipLrgNam = vVal(i)
    i = i + 1
    tPrmZipData.ZipMdlNam = vVal(i)
    i = i + 1
    tPrmZipData.ZipSmlNam = vVal(i)
    i = i + 1
    tPrmZipData.ZipLclAra = vVal(i)
    
    Exit Sub

ZipMstLoad_ErrorTraping:
    Resume Next

End Sub
    
Public Sub ZipMstStore(sPrmKey As String, sPrmValue As String, tPrmZipData As ZipMstRec)

    
    sPrmKey = tPrmZipData.ZipCod & Chr(5)
    
    sPrmValue = tPrmZipData.ZipLrgNam & Chr(5)
    sPrmValue = sPrmValue & tPrmZipData.ZipMdlNam & Chr(5)
    sPrmValue = sPrmValue & tPrmZipData.ZipSmlNam & Chr(5)
    sPrmValue = sPrmValue & tPrmZipData.ZipLclAra & Chr(5)
    
End Sub

    
Public Sub MgdMstRead(sPrmCod As String, sPrmDte As String, tPrmMgdData As MgdMstRec)

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sRetVal As String
    
    sCmpKey = sPrmCod & Chr(5)
    sCurKey = sCmpKey & "99999999" & Chr(5)
    sCurKey = mSetPrev("MgdMst", sCurKey)
    Do
        sCurKey = mReadPrev("MgdMst", sCurKey, sCmpKey, sRetVal)
        If sCurKey = "" Then Exit Do
        
        Call MgdMstLoad(sRetVal, tPrmMgdData)
        
        If CLong(sPrmDte) >= CLong(tPrmMgdData.MgdAdpDte) Then
            Exit Sub
        Else
            Call MgdMstLoad("", tPrmMgdData)
        End If
    Loop
End Sub

Public Sub ChtManMstLoad(sPrmValue As String, tPrmChtManData As ChtManMstRec)
    Dim i As Integer

    i = 1
    tPrmChtManData.ChtManRomNum = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmChtManData.ChtManRakNum = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmChtManData.ChtManCabNum = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmChtManData.ChtManDtlNum = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmChtManData.ChtManChtNum = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmChtManData.ChtManPatNam = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmChtManData.ChtManResNum = piece(sPrmValue, Chr(5), i)
    i = i + 1
    tPrmChtManData.ChtManCurStt = piece(sPrmValue, Chr(5), i)

End Sub

Public Sub ChtManMstStore(sPrmKey As String, sPrmValue As String, tPrmChtManData As ChtManMstRec)
    
    sPrmKey = tPrmChtManData.ChtManRomNum & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmChtManData.ChtManRakNum), "@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmChtManData.ChtManCabNum), "@@@") & Chr(5)
    sPrmKey = sPrmKey & Format(CDouble(tPrmChtManData.ChtManDtlNum), "@@@@") & Chr(5)
    
    sPrmValue = Format(CDouble(tPrmChtManData.ChtManChtNum), "@@@@@@@@") & Chr(5)
    sPrmValue = sPrmValue & tPrmChtManData.ChtManPatNam & Chr(5)
    sPrmValue = sPrmValue & tPrmChtManData.ChtManResNum & Chr(5)
    sPrmValue = sPrmValue & tPrmChtManData.ChtManCurStt

End Sub

Public Sub DrsMstLoad(sPrmValue As String, DrsData As DrsMstRec)

    On Error GoTo DrsMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 70)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    With DrsData
    i = i + 1
    .DrsSotTyp = vVal(i)  'DrsMstKey  슬립구분
    i = i + 1
    .DrsSotCod = vVal(i)  'DrsMstKey  슬립종목(항목)
    i = i + 1
    .DrsSitCod = vVal(i)  'DrsMstKey  작성부서
    i = i + 1
    .DrsDtrCod = vVal(i)  'DrsMstKey  의사코드
    i = i + 1
    .DrsSlpCod = vVal(i)  'DrsMstKey  슬립종류(범주)
    i = i + 1
    .DrsOdrSeq = vVal(i)  'DrsMstKey  처방 Seq
    i = i + 1
    .DrsOdrCod = vVal(i)  '1          처방 코드
    i = i + 1
    .DrsCodNam = vVal(i)  '2          처방명
    i = i + 1
    .DrsOdrQty = vVal(i)  '3          투여량
    i = i + 1
    .DrsOdrTms = vVal(i)  '4          횟수
    i = i + 1
    .DrsOdrDay = vVal(i)  '5          일수
    i = i + 1
    .DrsUsgCod = vVal(i)  '6          용법
    i = i + 1
    .DrsSpmCod = vVal(i)  '7          검체코드
    i = i + 1
    .DrsSpcYon = vVal(i)  '8          특기여부
    i = i + 1
    .DrsSpcCmt = vVal(i)  '9          특기사항
    i = i + 1
    .DrsDgsRol = vVal(i)  '10         방사선촬영부위(Left, Right)
    i = i + 1
    .DrsAdpTyp = vVal(i)  '11         적용구분        ----> 96/05/07 추가 (GrpMst에 있는 Field)
    i = i + 1
    .DrsMthCod = vVal(i)  '12         행위구분        ----> 96/05/07 추가 (GrpMst에 있는 Field)
    i = i + 1
    .DrsDgsYon = vVal(i)  '13         상병코드여부    ----> 96/05/07 추가 (GrpMst에 있는 Field)
    i = i + 1
    .DrsInsYon = vVal(i)  '14         급여구분        ----> 96/05/07 추가 (GrpMst에 있는 Field)
    i = i + 1
    .DrsSlpDep = vVal(i)  '15
    i = i + 1
    .DrsDgsEtc = vVal(i)  '16         추가사항
    i = i + 1
    .DrsSlpSeq = vVal(i)  'DrsMstKey  슬립순서
    i = i + 1
    .DrsRepYon = vVal(i)  '           의사사용량
    i = i + 1
    .DrsItmCod = vVal(i)  '19         항목정보 ====> 2001/11/30 james 추가 (SeeMst의 SeeItmCod)

    End With
    Exit Sub
    
DrsMstLoad_ErrorTraping:
    Resume Next

End Sub

Public Sub DrsMstStore(sCurKey As String, sRetVal As String, DrsData As DrsMstRec)

    With DrsData
        sCurKey = .DrsSotTyp & Chr(5)
        sCurKey = sCurKey & .DrsSotCod & Chr(5)              'DrsMstKey  슬립종목(항목)
        sCurKey = sCurKey & .DrsSitCod & Chr(5)              'DrsMstKey  작성부서
        sCurKey = sCurKey & .DrsDtrCod & Chr(5)              'DrsMstKey  의사코드
        sCurKey = sCurKey & Format(Trim(.DrsSlpCod), "@@@@@") & Chr(5)              'DrsMstKey  슬립종류(범주)
        sCurKey = sCurKey & Format(Trim(.DrsOdrSeq), "@@@@@") & Chr(5)             'DrsMstKey  처방 Seq
        
        sRetVal = .DrsOdrCod & Chr(5)               '1          처방 코드
        sRetVal = sRetVal & .DrsCodNam & Chr(5)     '2          처방명
        sRetVal = sRetVal & .DrsOdrQty & Chr(5)     '3          투여량
        sRetVal = sRetVal & .DrsOdrTms & Chr(5)     '4          횟수
        sRetVal = sRetVal & .DrsOdrDay & Chr(5)     '5          일수
        sRetVal = sRetVal & .DrsUsgCod & Chr(5)     '6          용법
        sRetVal = sRetVal & .DrsSpmCod & Chr(5)     '7          검체코드
        sRetVal = sRetVal & .DrsSpcYon & Chr(5)     '8          특기여부
        sRetVal = sRetVal & .DrsSpcCmt & Chr(5)     '9          특기사항
        sRetVal = sRetVal & .DrsDgsRol & Chr(5)     '10         방사선촬영부위(Left, Right)
        sRetVal = sRetVal & .DrsAdpTyp & Chr(5)     '11         적용구분        ----> 96/05/07 추가 (GrpMst에 있는 Field)
        sRetVal = sRetVal & .DrsMthCod & Chr(5)     '12         행위구분        ----> 96/05/07 추가 (GrpMst에 있는 Field)
        sRetVal = sRetVal & .DrsDgsYon & Chr(5)     '13         상병코드여부    ----> 96/05/07 추가 (GrpMst에 있는 Field)
        sRetVal = sRetVal & .DrsInsYon & Chr(5)     '14         급여구분        ----> 96/05/07 추가 (GrpMst에 있는 Field)
        sRetVal = sRetVal & .DrsSlpDep & Chr(5)     '15
        sRetVal = sRetVal & .DrsDgsEtc & Chr(5)     '16         추가사항
        sRetVal = sRetVal & Format(Trim(.DrsSlpSeq), "@@@") & Chr(5)     'DrsMstKey  슬립순서
        sRetVal = sRetVal & .DrsRepYon & Chr(5)     '           의사사용량
        sRetVal = sRetVal & .DrsItmCod & Chr(5)     '19         항목정보 ====> 2001/11/30 james 추가 (SeeMst의 SeeItmCod)
    End With
    
End Sub

Public Sub KgoMstLoad(sPrmValue As String, KgoData As KgoMstRec)

    On Error GoTo KgoMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 10)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    With KgoData
    
    i = i + 1
    .KgoSeeCod = vVal(i)    '수가코드
    i = i + 1
    .KgoUntCod = vVal(i)    'Kg 단위 1,2,...
    i = i + 1
    .KgoOdrQty = vVal(i)    '
    i = i + 1
    .KgoDtrQty = vVal(i)
    i = i + 1
    .KgoSpcRem = vVal(i)
    i = i + 1
    .KgoUpdDtm = vVal(i)
    i = i + 1
    .KgoUidCod = vVal(i)
    
    End With

KgoMstLoad_ErrorTraping:
    Exit Sub

End Sub
Public Sub OutMstLoad(sPrmValue As String, OutData As OutMstRec)

    On Error GoTo OutMstLoad_ErrorTraping

    Dim i       As Integer
    Dim vVal()  As String

    If sPrmValue = "" Then
        sPrmValue = sPrmValue & Replicate(Chr(5), 10)
    End If

    vVal = Split(sPrmValue & Replicate(Chr(5), 2), Chr(5))

    i = -1
    
    With OutData
    
    i = i + 1
    .OutOdrDte = vVal(i)
    i = i + 1
    .OutNum = vVal(i)     '
    i = i + 1
    .OutUpdDtm = vVal(i)     '
    
    
    End With

OutMstLoad_ErrorTraping:
    Exit Sub

End Sub


Public Sub KgoMstStore(sCurKey As String, sRetVal As String, KgoData As KgoMstRec)

    With KgoData

    sCurKey = .KgoSeeCod & Chr(5)    '수가코드
    
    sRetVal = .KgoUntCod & Chr(5)   'Kg 단위 1,2,...
    sRetVal = sRetVal & .KgoOdrQty & Chr(5)
    sRetVal = sRetVal & .KgoDtrQty & Chr(5)
    sRetVal = sRetVal & .KgoSpcRem & Chr(5)
    sRetVal = sRetVal & .KgoUpdDtm & Chr(5)
    sRetVal = sRetVal & .KgoUidCod & Chr(5)
    
    End With
    
End Sub

Public Sub OutMstStore(sCurKey As String, sRetVal As String, OutData As OutMstRec)

    With OutData

    sCurKey = .OutOdrDte & Chr(5)
    
    sRetVal = .OutNum & Chr(5)    'Kg 단위 1,2,...
    sRetVal = sRetVal & .OutUpdDtm & Chr(5)
    
    
    End With
    
End Sub

