--drop    table  scrprexh
;
create table scrprexh                                     -- 접수 테이블 
(
   hos_org_no              varchar2(8)     not null,      -- 병원 기호
   smp_no                  varchar2(12)    not null,      -- 검체번호
   prcp_seq                numeric(2)      not null,      -- 검체내 처방 seq(mosxpslh's acp_no)
   exam_seq                numeric(2)      not null,      -- 검사코드 seq
   parn_seq                numeric(2)      not null,      -- 상위코드 seq 
   prcp_yn                 varchar2(1)     not null,      -- 처방코드여부
   gr_sng_cd               varchar2(1)     not null,      -- G/S
   cd                      varchar2(20)    not null,      -- 처방,검사,균주,항생제,결과코드
   pt_no                   varchar2(8)     not null,      -- 환자 등록번호
   smp_stus                varchar2(1)     not null,      -- 0 : 접수 1 : 예비 2 : 일부
                                                          -- 3 : 최종 4 : 수정 6 : Interface 
   acp_dt                  varchar2(12)    not null,      -- 접수 일자
   acp_empno               varchar2(12)    not null,      -- 접수자 사번
   rept_dt                 varchar2(12)        null,      -- 보고 일자
   rept_empno              varchar2(12)        null,      -- 보고자 사번
   excel_yn                varchar2(1)         null,      -- 외주검사 Excel 저장여부
   scl_key                 varchar2(30)        null,      -- SCL 연결KEY
   slip_knd                varchar2(3)         null,      -- Slip종류
   wrk_no                  numeric(5)          null       -- 작업번호
)
pctfree        5
pctused       85
initrans       4
tablespace  TSRA
            storage (initial          10M
                     next             5M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create unique index scrprexh_ux1 on scrprexh(hos_org_no,smp_no,prcp_seq,exam_seq)
pctfree       10
initrans       4
tablespace  XSRA
            storage (initial          2M
                     next             1M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create index scrprexh_ix1 on scrprexh(smp_no, prcp_seq, hos_org_no)
pctfree       10
initrans       4
tablespace  XSRA
            storage (initial          2M
                     next             1M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create index scrprexh_ix2 on scrprexh(SUBSTR(cd, 1, 3), prcp_yn, hos_org_no)
pctfree       10
initrans       4
tablespace  XSRA
            storage (initial          2M
                     next             1M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create index scrprexh_ix3 on scrprexh(SUBSTR(acp_dt, 1, 8), hos_org_no, pt_no, smp_no)
pctfree       10
initrans       4
tablespace  XSRA
            storage (initial          2M
                     next             1M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create index scrprexh_ix4 on scrprexh(hos_org_no,SUBSTR(rept_dt,1,8))
pctfree       10
initrans       4
tablespace  XSRA
            storage (initial          2M
                     next             1M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create index scrprexh_ix5 on scrprexh(scl_key,hos_org_no)
pctfree       10
initrans       4
tablespace  XSRA
            storage (initial          2M
                     next             1M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create index scrprexh_ix6 on scrprexh(SUBSTR(acp_dt, 1, 8), hos_org_no, slip_knd, wrk_no)
pctfree       10
initrans       4
tablespace  XSRA
            storage (initial          2M
                     next             1M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create synonym outmed.scrprexh for med.scrprexh;
grant select, insert, update on scrprexh to outmed;

exit;
