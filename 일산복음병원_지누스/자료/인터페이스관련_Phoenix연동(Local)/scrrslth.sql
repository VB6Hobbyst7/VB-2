--drop    table  scrrslth
;
create table scrrslth                                     -- 검사 결과
(
   hos_org_no              varchar2(8)     not null,      -- 병원 기호
   smp_no                  varchar2(12)    not null,      -- 검체번호
   prcp_seq                numeric(2)      not null,      -- 검체내 처방 seq
   exam_seq                numeric(2)      not null,      -- 검사코드 seq
   cd                      varchar2(20)    not null,      -- 처방,검사,균주,항생제,결과코드
   rept_seq                numeric(3)      not null,      -- 출력순서
   pt_no                   varchar2(8)     not null,      -- 환자번호
   dla_stus                varchar2(2)     not null,      -- 델타 상태
                                                          -- 첫자리   정상 - N, 비정상  - Y
                                                          -- 둘째자리 정상 - N, High - H, Low - L
   pnc_stus                varchar2(2)     not null,      -- 패닉 상태
                                                          -- 첫자리   정상 - N, 비정상  - Y
                                                          -- 둘째자리 정상 - N, High - H, Low - L
   exam_stus               varchar2(1)     not null,      -- 검사 상태
                                                          -- 0 : 접수 1 : 예비 2 : 일부
                                                          -- 3 : 최종 4 : 수정 6 : Interface
   exam_rslt               varchar2(100)       null,      -- 검사 결과
   exam_rslt_cd            varchar2(2)         null,      -- 검사 결과 코드
   mach_rslt               varchar2(80)        null,      -- 장비 검사 결과
   mach_rslt_cd            varchar2(2)         null,      -- 장비 결과 코드
   exam_dt                 varchar2(12)        null,      -- 검사일시
   exam_empno              varchar2(12)        null,      -- 검사자사번
   etc1                    varchar2(10)        null,      -- size(sign) : 2009.01.06 jangmc
   etc2                    varchar2(10)        null       -- 기타2
)
pctfree        5
pctused       85
initrans       4
tablespace  TSCC
            storage (initial          20M
                     next             10M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create unique index scrrslth_ux1 on scrrslth(hos_org_no,smp_no,prcp_seq,exam_seq,rept_seq)
pctfree       10
initrans       4
tablespace  XSCC
            storage (initial          10M
                     next             5M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create index scrrslth_ix1 on scrrslth(hos_org_no, pt_no, cd, exam_dt)
pctfree       10
initrans       4
tablespace  XSCC
            storage (initial          10M
                     next             5M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create index scrrslth_ix2 on scrrslth(pt_no, exam_dt, hos_org_no)
pctfree       10
initrans       4
tablespace  XSCC
            storage (initial          10M
                     next             5M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create synonym outmed.scrrslth for med.scrrslth;
grant select, insert, update on scrrslth to outmed;

exit;
