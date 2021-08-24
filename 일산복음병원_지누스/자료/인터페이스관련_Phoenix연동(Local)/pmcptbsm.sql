--drop    table  pmcptbsm
;
create table pmcptbsm                                      -- 환자인적내역
(
   hos_org_no              varchar2(10)     not null,      -- 병원기관번호
   pt_no                   varchar2(8)      not null,      -- 환자번호
   pt_nm                   varchar2(30)     not null,      -- 환자성명
   bth_y                   varchar2(2)      not null,      -- 출생년도
   ssn_1                   varchar2(6)      not null,      -- 주민등록번호1
   ssn_2                   varchar2(7)      not null,      -- 주민등록번호2
   pt_no_vld_yn            varchar2(1)      not null,      -- 환자번호유효여부 (Y/N)
   tel_no_1                varchar2(3)      not null,      -- 전화번호1
   tel_no_2                varchar2(4)      not null,      -- 전화번호2
   tel_no_3                varchar2(4)      not null,      -- 전화번호3
   hp_no_1                 varchar2(3)          null,      -- 휴대폰번호1
   hp_no_2                 varchar2(4)          null,      -- 휴대폰번호2
   hp_no_3                 varchar2(4)          null,      -- 휴대폰번호3
   mail_addr               varchar2(50)         null,      -- 메일주소
   zpcd_1                  varchar2(3)      not null,      -- 우편번호1
   zpcd_2                  varchar2(3)      not null,      -- 우편번호2
   zpcd_seq                numeric(5)       not null,      -- 우편번호순서
   lw_addr                 varchar2(60)     not null,      -- 하위주소
   ocp_cd                  varchar2(2)          null,      -- 직업코드 (의무기록)
                                                           -- 방사선 촬영여부로 활용(Y/N)
   rgn_cd                  varchar2(2)          null,      -- 종교코드 (의무기록)
                                                           -- pks20080814 : Net 환자여부(Y)
   adms_yn                 varchar2(1)      not null,      -- 재원여부 (Y:재원, N:재원아님)
   indi_ucolt_yn           varchar2(1)      not null,      -- 개인미수여부 (Y/N)
   abo_blotyp              varchar2(2)          null,      -- abo혈액형
   rh_blotyp               varchar2(2)          null,      -- rh혈액형
                                                           -- cij 2008.09.30 varchar(1)?->varchar(2)
   fs_chos_ymd             varchar2(8)      not null,      -- 최초내원일자
   lst_chos_ymd            varchar2(8)      not null,      -- 최종내원일자
   seri_no                 varchar2(10)         null,      -- 중증환자고유번호
                                                           -- Table로 변경됨.
   /* --------------------[추가]----------------------------------------------------------------- */
   pt_nat                  varchar2(6)          null,      -- 환자국적(외국인일 경우)
                                                           -- 기초코드 HZC-C010
   brth_ymd                varchar2(8)      not null,      -- 생년월일
   brth_wo_flag            varchar2(1)      not null,      -- 0:양력, 1:음력
   wrk_ymd                 varchar2(8)      not null,      -- 작업일자
   wrk_tm                  varchar2(6)      not null,      -- 작업시간
   wrk_empno               varchar2(12)     not null,      -- 작업자

   /* --------------------------------------------------------------------------------------------
    * 진단검사 결과입력시 저장
    * pks20090116 추가
    * -------------------------------------------------------------------------------------------- */
   vre_yn                  varchar2(1)          null,      -- 항생제 검사에서 VRE가 나온 검체(Y/N)
   hiv_yn                  varchar2(1)          null,      -- HIV 양성 환자(Y/N)
   hpts_yn                 varchar2(1)          null       -- 매독 or 간염 보균 환자(Y/N)
)
pctfree        5
pctused       85
initrans       4
tablespace  TPMC
            storage (initial         10M
                     next             5M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create unique index pmcptbsm_ux1 on pmcptbsm(hos_org_no,pt_no)
pctfree       10
initrans       4
tablespace  XPMC
            storage (initial          5M
                     next             3M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create index pmcptbsm_ix1 on pmcptbsm(hos_org_no,ssn_1,ssn_2)
pctfree       10
initrans       4
tablespace  XPMC
            storage (initial          5M
                     next             3M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create index pmcptbsm_ix2 on pmcptbsm(hos_org_no,pt_nm)
pctfree       10
initrans       4
tablespace  XPMC
            storage (initial          5M
                     next             3M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create synonym outmed.pmcptbsm for med.pmcptbsm;
grant select, insert, update on pmcptbsm to outmed;

exit;
