--drop    table  mosxpslh
;
create table mosxpslh                                     -- 처방내역
(
//hos_org_no,pt_no,adms_ymd,med_ymd,prcp_cd,prcp_cd_seq
   hos_org_no              varchar2(10)     not null,     -- 병원기관번호
   pt_no                   varchar2(8)      not null,     -- 환자번호
   adms_ymd                varchar2(8)      not null,     -- 입원일자(입원)/00000000(외래)
   med_ymd                 varchar2(8)      not null,     -- 진료일자
   prcp_cd                 varchar2(12)     not null,     -- 처방코드
   prcp_cd_seq             numeric(3)       not null,     -- 처방코드일련번호
   med_dpt_cd              varchar2(6)      not null,     -- 진료과코드
   calc_ymd                varchar2(8)      not null,     -- 계산일자 (처방확인에서 이용)
   req_ymd                 varchar2(8)      not null,     -- 청구일자 (청구:실시일자 기준 청구)
   mn_sub_typ_cd           varchar2(1)      not null,     -- 주부유형코드 (M:주유형,S:부유형)
   prcp_ip_dt              varchar2(12)     not null,     -- 처방입력일시
   prcp_typ_cd             varchar2(1)      not null,     -- 처방구분코드
                                                          -- O:원처방
                                                          -- C:재원중인 환자의 외래 비급여처방
                                                          -- D:전달후삭제처방
                                                          -- L:전달전삭제처방
                                                          -- S:심사에서삭제한처방
   calc_typ_cd             varchar2(1)      not null,     -- 계산구분코드 (M:Minus 계산처리)
   prcp_stus_cd            varchar2(1)      not null,     -- 처방상태코드
                                                          -- 0:처방(Raw)
                                                          -- 1:수납(Pay)
                                                          -- 2:채취(Extract)
                                                          -- 3:접수(Accept)
                                                          -- 4:(검사진행 : 임상, 임시결과 : 지원)
                                                          -- 5:일부보고(Part)
                                                          -- 6:최종결과(Finish)
                                                          -- 7:수정결과(Modify)
                                                          -- 9:검사 최초 I/F결과 : pks 20081208 추가
                                                          -- A:Hold처방 -> Release 시 0:처방상태
                                                          -- -:Chk(처방확인)
   prcp_knd_cd             varchar2(1)      not null,     -- 처방종류코드 (O:외래,E:응급,I:입원)
   prcp_cap                numeric(8,4)     not null,     -- 처방용량(마취:시간)
   prcp_unit               varchar2(6)          null,     -- 처방단위
   prcp_cnt                numeric(2)       not null,     -- 처방횟수(마취:분)
   prcp_d_cnt              numeric(3)       not null,     -- 처방일수
   mth_cd                  varchar2(4)          null,     -- 방법코드
   pay_typ_cd              varchar2(1)      not null,     -- 급여구분코드
                                                          -- X:비급여      , Y:급여
                                                          -- N:급여본인100%, -:None
   op_typ_cd               varchar2(1)      not null,     -- 수술구분코드 2008.10.20, 변경
                                                          -- 수술(estm_cls = 'MS', estm_mn_cd4 = '9')
                                                          --   0 : 주수술
                                                          --   1 : 제2의 수술
                                                          --   2 : 재수술(주된수술)
                                                          --   3 : 재수술(제2의수술)
                                                          -- 마취(estm_cls = 'L', anes_time_flag  = 'Y')
                                                          --   0 : 해당없음   
                                                          --   5 : 심폐체외순환법
                                                          --   6 : 일측폐환기법
                                                          --   7 : 고빈도제트환기법
                                                          --   8 : 개흉적 심장수술
                                                          --   9 : 뇌정양, 뇌질환에 대한 개두술
                                                          -- 혈관조영촬영(estm_cls = 'HC', vesl_pho_flag = 'Y')
                                                          --   0 : 해당없음
                                                          --   1 : 단측
                                                          --   2 : 양측
   emer_typ_cd             varchar2(1)      not null,     -- 응급구분여부
                                                          -- E:응급,D:퇴원약,O:외출약,P:P.R.N,-:None
                                                          -- S:무료, p:액팅한 PRN, M:Self Med
                                                          -- S:원무 금액산정안함, 청구 금액산정함 
                                                          -- M:원무, 청구 금액산정안함  
   nig_hol_med_typ_cd      varchar2(1)      not null,     -- 야간/휴일/임상구분코드
                                                          -- N:야간,L:임상,H:휴일,-:None
   spmed_yn                varchar2(1)      not null,     -- 특진여부 (Y:특진,N:일반)
   trfu_typ_cd             varchar2(1)      not null,     -- 수혈구분코드
                                                          -- P:예정수혈,T:당일수혈,C:혈액은행출고,
                                                          -- -:None
   pwr_port_typ_cd         varchar2(1)      not null,     -- 가루약(Pulverize)/이동(Portabl)구분코드
                                                          -- Y:가루약(Pulverize),P:Portable,-:None
   prcp_cmm_dpt            varchar2(6)      not null,     -- 처방전달부서
   prcp_dr_empno           varchar2(12)     not null,     -- 처방의사사번
   hop_dt                  varchar2(12)         null,     -- 희망일시
   exec_dr_empno           varchar2(12)     not null,     -- 실시의사사변
   prcp_ord                numeric(3)       not null,     -- 처방순서 (화면에 보여주는 순서)
                                                          -- 일반검진인경우->dmd_item_cd 값 Setting
   spp_dpt_rsc_yn          varchar2(1)      not null,     -- 지원부서예약여부
                                                          -- Y:검사 예약완료,-:None
   acp_dpt                 varchar2(6)          null,     -- 접수부서
   smp_dt                  varchar2(12)         null,     -- 채취일시 (주사=주사실시일시)
   smp_no                  varchar2(12)         null,     -- 채취번호 (ex:02-1234567-1-1)
   acp_dt                  varchar2(12)         null,     -- 접수일시
   acp_no                  numeric(5)           null,     -- 접수번호(약처방전교부번호)
                                                          -- 진단검사:채취일련번호
   tril_ymd                varchar2(8)          null,     -- 실시일자(약처방전교부일자)
   nurs_cfr_stus_cd        varchar2(1)          null,     -- 간호확인상태코드
                                                          -- 'Y':확인, Default:'-'
   hos_yn                  varchar2(1)      not null,     -- 원내여부 (-:원내,Y:원내,N:원외)
   hos_rsn_cd              varchar2(2)      not null,     -- 원내사유코드
   hoso_phm_rcv_yn         varchar2(1)      not null,     -- 원외약수령여부(Y:원외약수령여부,-:None)
   tel_prcp_yn             varchar2(1)      not null,     -- 전화처방여부  (T:전화처방,-:None)
   cnl_rsn_cd              varchar2(2)      not null,     -- 취소사유코드/반납의뢰
                                                          -- (01:처방변경,02:부작용..)
                                                          -- 반납확인시 새로운 처방에는 없어야함.
   add_prcp_empno          varchar2(12)         null,     -- 추가처방담당자사번
                                                          -- 지원부서추가처방담당자사번
   gnl_add_typ_cd          varchar2(1)          null,     -- 정규/추가/타추구분코드/전입처방여부
                                                          -- 0:정규,1:추가,2:타추,3:전입처방,
                                                          -- 4:D/C 확인 처방
   calc_yn                 varchar2(1)      not null,     -- 계산여부 (Y:계산)
   blodo_typ_cd            varchar2(1)      not null,     -- 헌혈증구분코드 (3:320cc,4:400cc,-:None)
   rep_prcp_typ_cd         varchar2(1)      not null,     -- 대표처방구분코드
                                                          -- 0:대표(종검),1:풀린처방(종검에서 풀린처방),2:일반
   rv_user_empno           varchar2(12)         null,     -- 심사담당자사번
   rv_stus_cd              varchar2(1)          null,     -- 심사상태코드(P:심사,-:None)
   rv_dt                   varchar2(12)         null,     -- 심사일시
   lst_yn                  varchar2(1)      not null,     -- 최종여부(Y/-)
                                                          -- 사전에서 삭제(P)일 경우 
                                                          -- 이전 처방 상태를 저장한다.
   req_yn                  varchar2(1)      not null,     -- 청구유무(Y:청구/A:추가/N)
   edt_dt                  varchar2(12)     not null,     -- 수정일시
   user_empno              varchar2(12)     not null,     -- 담당자사번
   input_gubun             varchar2(2)      not null,     -- D0:정규,D1:추가,T0:타과
                                                          -- (아래는 현재 미사용)
                                                          -- D0:정규,D1:추가,NR:간호사,
                                                          -- TB:타병동(회복실,수술실),
                                                          -- TG:타과추가처방(타진료과),
                                                          -- MS:진료지원부서,SS:심사
   input_part              varchar2(6)      not null,     -- (input_gubun)D0,D1일때 환자진료과,
                                                          -- T0일때 타진료과
                                                          -- (아래는 현재 미사용)
                                                          -- (input_gubun)D0,D1,NR,TB일때 환자진료과
                                                          -- TG일때 타진료과, MS일때 진료지원부서
   gr_ser                  numeric(3)       not null,     -- 진단검사:연속처방의 묶음 일련번호
                                                          -- 현재 미사용
   gtt                     numeric(3)       not null,     -- 주사속도
   micro_gtt               numeric(3)       not null,     -- 주사속도
   spe_prep                varchar2(1)      not null,     -- 특별조제 1:항암조제, 2:무균조제
                                                          --          3:직불수납, -:Default
                                                          -- 현재미사용(다음으로사용)
                                                          -- 검진구분 1:1차검진 ,2:2차검진,
                                                          --          3:구강검진,E:암검진
   smp_empno               varchar2(12)     not null,     -- 채취자
                                                          -- 검사자 (기타지원부서 검사자)
   spp_dpt_rsv_ymd         varchar2(12)     not null,     -- 지원부서 예약일시
   result_ymd              varchar2(12)     not null,     -- 결과일시
   dc_typ                  varchar2(1)      not null,     -- DC구분(D:반납, -:정상)
                                                          -- S:오더정지(현재 미사용)
   prcp_dc_seq             numeric(3)       not null,     -- DC코드일련번호
                                                          -- 현재 미사용
   wrk_no                  numeric(6)       not null,     -- 작업번호
   breakfast               varchar2(1)          null,     -- 마약처방전 출력 여부
   lunch                   varchar2(1)          null,     -- 중식(Y:식사/y:집계/N:안함)
                                                          -- 외래의 원외처방이 입원으로 전환시(O)
                                                          -- 2007.06 HY
   dinner                  varchar2(1)          null,     -- 석식(Y:식사/y:집계/N:안함)
                                                          -- 2007.09.17 전일처방과 틀릴경우 'Y'
   midnight                varchar2(1)          null,     -- 야식(Y:식사/y:집계/N:안함)
   wo_gubn                 varchar2(1)          null,     -- 양/한방구분(W:양 O:한)
                                                          -- NULL일때는 'W'
   nurs_cfr_dt             varchar2(12)         null,     -- 간호확인일시
   nurs_cfr_empno          varchar2(12)         null,     -- 간호확인담당자
   nurs_exe_dt             varchar2(12)         null,     -- 간호실시일시
   nurs_exe_empno          varchar2(12)         null,     -- 간호실시담당자
   op_yn                   varchar2(1)          null,     -- 수술 처방 (Pre Order:P, Post Order:Y)
   rate_cal_fg             varchar2(1)          null,     -- 보호정신과 행위별 계산 유무 
                                                          -- (0:정액금액에 포함, 1:행위별 계산)
   act_a_cnt               numeric(3)           null,     -- Acting all Count (tims * days)
   act_r_cnt               numeric(3)           null,     -- Acting Check Count
   tot_d_cnt               numeric(3)           null,     -- 처방일수(총일수)
   nb_pt_no                varchar2(8)          null,     -- cij 2008.11.18 : 신생아 이동 처방 신생아 환자번호
                                                          -- 신생아일 경우 엄마번호
   nb_prcp_cd_seq          numeric(3)           null      -- 신생아 이동처방 처방코드일련번호
                                                          -- 신생아일 경우 엄마처방코드일련번호 
-- order_gubun             varchar2(2)      not null,     -- A1:Diet, A2:V/S, A3:활동사항, A4:BW,
                                                          -- A5:I/O, A9:기타, D1:내복약, D2:외용약,
                                                          -- D3:주사약, D4:FLUID, D5:처치약품,
                                                          -- F1:OP, F2:TREATMENT, B1:LAB, B2:BLOOD,
                                                          -- E1:PATHOLOGY, C1,C2:X-RAY, G1:METERIALS
--   spc_cd                  varchar2(6)      not null      -- 검체코드
-- mix_yn                  varchar2(1)      not null,     -- MIX(0:보통, 1:MIX)
-- prn_typ_cd              varchar2(1)      not null,     -- PRN 구분(1:PRN, 2:PREPARE, 3:PREMED,
                                                          --          4:V/O, 5:퇴원후 오더)
-- bil_plc                 varchar2(6)      not null,     -- 가야할곳
-- slip_kind               varchar2(3)      not null,     -- 서식지종류
-- admr_gubn               varchar2(1)      not null,     -- 투약번호 구분
-- wklst_yn                varchar2(1)      not null,     -- worklist 대상(1:Yes, 0:No)
)
pctfree        5
pctused       85
initrans       4
tablespace  TMOS
            storage (initial         10M
                     next             5M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create unique index mosxpslh_ux1 on mosxpslh(hos_org_no,pt_no,adms_ymd,med_ymd,prcp_cd,prcp_cd_seq)
pctfree       10
initrans       4
tablespace  XMOS
            storage (initial          5M
                     next             3M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create index mosxpslh_ix1 on mosxpslh(med_ymd,pt_no,adms_ymd,hos_org_no,prcp_cd,prcp_cd_seq,prcp_typ_cd)
pctfree       10
initrans       4
tablespace  XMOS
            storage (initial          5M
                     next             3M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create index mosxpslh_ix2 on mosxpslh(smp_no,hos_org_no,prcp_stus_cd)
pctfree       10
initrans       4
tablespace  XMOS
            storage (initial          5M
                     next             3M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create index mosxpslh_ix3 on mosxpslh(hos_org_no, tril_ymd, acp_no)
pctfree       10
initrans       4
tablespace  XMOS
            storage (initial          5M
                     next             3M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create index mosxpslh_ix4 on mosxpslh(SUBSTR(acp_dt,1,8), hos_org_no, smp_no)
pctfree       10
initrans       4
tablespace  XMOS
            storage (initial          5M
                     next             3M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create index mosxpslh_ix5 on mosxpslh(SUBSTR(hop_dt,1,8), hos_org_no)
pctfree       10
initrans       4
tablespace  XMOS
            storage (initial          5M
                     next             3M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create index mosxpslh_ix6 on mosxpslh(SUBSTR(med_ymd,1,6),hos_org_no)
pctfree       10
initrans       4
tablespace  XMOS
            storage (initial          5M
                     next             3M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;

create index mosxpslh_ix7 on mosxpslh(SUBSTR(spp_dpt_rsv_ymd,1,8),hos_org_no)
pctfree       10
initrans       4
tablespace  XMOS
            storage (initial          5M
                     next             3M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;

exit;
