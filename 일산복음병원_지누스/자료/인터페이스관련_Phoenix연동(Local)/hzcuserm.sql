--drop    table  hzcuserm
;
create table hzcuserm                                     -- 사용자관리
(
   hos_org_no              varchar2(10)     not null,     -- 병원기관번호
   empno                   varchar2(12)     not null,     -- 사번
   fr_ymd                  varchar2(8)      not null,     -- 시작일자
   to_ymd                  varchar2(8)      not null,     -- 종료일자
   pwd                     varchar2(12)     not null,     -- 암호
   empnm                   varchar2(20)     not null,     -- 성명
   dpt_cd                  varchar2(6)      not null,     -- 진료과코드(부서코드)
   ocptyp                  varchar2(1)      not null,     -- 직종
                                                          -- D:의사, N:간호사, T:기사,
                                                          -- P:약사, E:사무원
   spmed_yn                varchar2(1)          null,     -- 선택진료여부(Y:선택진료,N:비선택진료)
   lcns_no                 varchar2(6)          null,     -- 면허번호
   pos                     varchar2(1)      not null,     -- 소속 -> 'emc전송제외 의사(Y/-)'로 변경
                                                          -- 간호사(병동,외래,수술실,회복실…)
   use_yn                  varchar2(1)      not null,     -- 사용여부(Y:사용,N:못함)
   use_dut                 varchar2(1)      not null,     -- 사용업무
                                                          -- H : 원무 R : 심사 M : 관리 C : 경영
                                                          -- D : 처방 N : 간호 A : 지원 
   edt_dt                  varchar2(12)     not null,     -- 수정일시
   user_empno              varchar2(12)     not null,     -- 담당자사번
   tool_list               varchar2(900)    not null,     -- ToolBar 리스트
   cert_yn                 varchar2(1)      not null,     -- 인증여부(Y/N)
   secure_pwd              varchar2(10)     not null,     -- 보안암호(개인)
   sign_pwd                varchar2(10)     not null,     -- 인증암호
   ssn_1                   varchar2(6)          null,     -- 주민번호 1(자격확인)
   ssn_2                   varchar2(7)          null,     -- 주민번호 2(자격확인)
   emc_send_cnt            numeric(5)       not null,     -- 응급의료센타 전송 갯수
   emc_send_week           varchar2(1)      not null      -- 응급의료센타 전송 요일
)
pctfree        5
pctused       85
initrans       4
tablespace  TRDO
            storage (initial         10M
                     next             5M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;
create unique index hzcuserm_ux1 on hzcuserm(hos_org_no,empno,fr_ymd)
pctfree       10
initrans       4
tablespace  XRDO
            storage (initial          5M
                     next             3M
                     pctincrease      0
                     minextents       1
                     freelists        5
                     freelist groups  1)
;

exit;
