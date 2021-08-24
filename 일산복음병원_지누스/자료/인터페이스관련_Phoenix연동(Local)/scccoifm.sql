--drop    table  scccoifm
;
create table scccoifm                                     -- 진단 검사 의학과 검사 코드 마스타
(
   hos_org_no              varchar2(8)      not null,     -- 병원 기호 
   exam_cd                 varchar2(12)     not null,     -- 검사 cd
                                                          -- 앞 첫자리가 L 이면 검사 코드
   exam_typ                varchar2(4)          null,     -- LL : 일반 검사 코드
                                                          -- LD : Diff Count 코드
                                                          -- LC : CBC 코드 
                                                          -- LM : 미생물 배양 코드
                                                          -- LY : 미생물 염색 코드
                                                          -- LR : Micro 결과 검사
                                                          -- LS : 계산 검사 
                                                          -- LB : Anti-Body Test
                                                          -- LA : Abo typing  검사
                                                          -- RM : 균주 코드
                                                          -- RA : 항생제 코드
                                                          -- RR : Micro 검사 결과 코드
   exam_nm                 varchar2(80)         null,     -- 검사 name
   exam_abbr               varchar2(30)         null,     -- 검사 약어명
   gr_typ                  varchar2(1)          null,     -- 그룹 코드 여부
   trust_typ               varchar2(4)          null,     -- 검사실 구분(scccodem.typ_cd:'04')
   stat_cd                 varchar2(1)          null,     -- 통계 코드 여부 Y/N--- M   
                                                          -- sccststm
   ctn_typ                 varchar2(1)          null,     -- 연속 검사 유무
                                                          -- 'Y'/'N'
   spc_cd                  varchar2(4)          null,     -- 검체 코드(scccodem.typ_cd:'02')
   spc_smp_amt             varchar2(10)         null,     -- 검체 채취량    
   spc_smp_unit            varchar2(3)          null,     -- 검체 단위
   spc_smp_typ             varchar2(4)          null,     -- 검체 채취자 구분
                                                          -- scccodem.typ_cd:'05'
   rslt_typ_cd             varchar2(4)          null,     -- 검사 결과 유형 (S)
                                                          -- scccodem.typ_cd:'06'
                                                          -- P:원형결과 scccprfm 
   int_len                 varchar2(1)          null,     -- 정수부 길이
   dec_len                 varchar2(1)          null,     -- 소수부 길이
   rslt_unit               varchar2(20)         null,     -- 결과 단위
   nml_typ_cd              varchar2(4)          null,     -- 정상치 유형 (S) 
                                                          -- scccdm.typ_cd:'07'
   nml_m_max               varchar2(10)         null,     -- 정상치 남자 상한
   nml_f_max               varchar2(10)         null,     -- 정상치 여자 상한
   nml_m_min               varchar2(10)         null,     -- 정상치 남자 하한
   nml_f_min               varchar2(10)         null,     -- 정상치 여자 하한
   qc_typ_cd               varchar2(4)          null,     -- 정도 관리 기준
                                                          -- scccdm.typ_cd:'08'
   dlt_max                 varchar2(10)         null,     -- delta 상한
   dlt_min                 varchar2(10)         null,     -- delta 하한
   pnc_max                 varchar2(10)         null,     -- panic 상한
   pnc_min                 varchar2(10)         null,     -- panic 하한
   exam_mach_cd            varchar2(4)          null,     -- 검사 장비 코드
                                                          -- scccodem.typ_cd:'09'
   qc_vld_trm              varchar2(4)          null,     -- 정도 관리 유효기간
   exam_cau_cd             varchar2(30)         null,     -- 검사 주의 사항 코드
                                                          -- SCL 검사코드
   exam_mth_cd             varchar2(4)          null,     -- 검사 방법 코드(scccdm.typ_cd:'11')
   exam_rslt_rept_yn       varchar2(1)          null,     -- 검사 결과 보고 여부
   lmt_trm_day             varchar2(2)          null,     -- 검사 소요 기간 (일)
                                                          -- 사용하지 않음
   lmt_trm_tm              varchar2(2)          null,     -- 검사 소요 기간 (시간)
   lmt_trm_mint            varchar2(2)          null,     -- 검사 소요 기간 (분)
   bs_exam_cd              varchar2(12)         null,     -- 기본 검사 코드
   bs_rslt_val             varchar2(30)         null,     -- 검사 기본 결과 값  (S)
   user_empno              varchar2(12)     not null,     -- 사용자사번
   use_typ                 varchar2(1)      not null,     -- 사용여부(Y:사용,D:정지)
   fr_dt                   varchar2(12)     not null,     -- 시작일자
   to_dt                   varchar2(12)     not null,     -- 종료일자
   spc_bott                varchar2(4)          null,     -- 검체 용기
                                                          -- scccodem.typ_cd:'13'
   slip_knd                varchar2(3)          null,     -- Slip종류
                                                          -- hzcbsdtm.cd_typ:'AST', hzcbsdtm.bas_cd:'C002'
   slip_sub                varchar2(4)          null,     -- Slip세부종류
                                                          -- hzcbsdtm.cd_typ:'AST', hzcbsdtm.bas_cd:'AB91'
   imm_trans_spc           varchar2(1)          null,     -- 즉시전달 검체(Y/N)
   reserve_spc             varchar2(1)          null,     -- 예약 검체(Y/N)
   ptbs_app_item           varchar2(1)          null,     -- 환자기본정보에 적용해야할 항목
                                                          -- VHI 중 하나이고 결과유형의 etc1이 Y/N일 경우 적용됨
                                                          -- -: 없음
                                                          -- V: 항생제검사에서
                                                                VRE(vancomycin resistant enterococcus)가 나온 환자
                                                          -- H: HIV 환자
                                                          -- I: 매독,간염 환자
   blod_type               varchar2(5)          null,     -- 혈액Type
                                                          -- scccodem.typ_cd:'14'
   emer_able_cd            varchar2(5)          null,     -- 응급가능여부
                                                          -- -: 없음, 1:근무시간 외에 불가능, 2:전체 불가능
   rcpt_info_cd            varchar2(5)          null,     -- 진료과/영수증 정보
                                                          -- hzcbsdtm(AST/R001)
)
pctfree        5
pctused       85
initrans       4
tablespace  TSCC
            storage (initial          10M
                     next             5M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create unique index scccoifm_ux1 on scccoifm(hos_org_no,exam_cd,use_typ,fr_dt)
pctfree       10
initrans       4
tablespace  XSCC
            storage (initial          2M
                     next             1M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create        index scccoifm_ix1 on scccoifm(exam_cd,use_typ,hos_org_no,fr_dt)
pctfree       10
initrans       4
tablespace  XSCC
            storage (initial          2M
                     next             1M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create        index scccoifm_ix2 on scccoifm(SUBSTR(exam_cd, 1, 3),use_typ,hos_org_no,fr_dt)
pctfree       10
initrans       4
tablespace  XSCC
            storage (initial          2M
                     next             1M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create        index scccoifm_ix3 on scccoifm(bs_exam_cd,use_typ,hos_org_no,fr_dt)
pctfree       10
initrans       4
tablespace  XSCC
            storage (initial          2M
                     next             1M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create        index scccoifm_ix4 on scccoifm(exam_nm,use_typ,hos_org_no,fr_dt)
pctfree       10
initrans       4
tablespace  XSCC
            storage (initial          2M
                     next             1M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create        index scccoifm_ix5 on scccoifm(exam_mach_cd,gr_typ,hos_org_no,fr_dt)
pctfree       10
initrans       4
tablespace  XSCC
            storage (initial          2M
                     next             1M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;

exit;
