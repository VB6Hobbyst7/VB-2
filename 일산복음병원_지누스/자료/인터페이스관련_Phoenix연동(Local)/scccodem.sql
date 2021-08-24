--drop    table  scccodem
;
create table scccodem                                     -- 진단 검사 의학과 코드 마스타
(
   hos_org_no              varchar2(8)      not null,     -- 병원 기호 
   typ_cd                  varchar2(4)      not null,     -- 구분
                                                          -- 01 : 검사 코드 구분
                                                          -- 02 : 검체 코드 
                                                          -- 03 : 각종 코드 매칭
                                                          -- 04 : 검사실 구분
                                                          -- 05 : 채취자 구분 
                                                          -- 06 : 결과 유형 
                                                                : ON,FT,AS - Hard Coding
                                                          -- 07 : 정상치 유형 
                                                          -- 08 : 정도 관리 유형
                                                          -- 09 : 장비 코드 
                                                          -- 10 : 주의 사항 
                                                          -- 11 : 검사 방법 
                                                          -- 12 : Diff Key Mapping
                                                          -- 13 : 검체 용기
                                                          -- 14 : 혈액제제 분류
                                                          -- 15 : 혈액형 분류
   cd                      varchar2(10)     not null,     -- 코드
   cd_desc                 varchar2(80)         null,     -- 코드 DESC 
   tribu_cd                varchar2(1)      not null,     -- 분류
                                                          -- '1':관리용
                                                          -- '2':업무용
   user_empno              varchar2(12)     not null,     -- 검체 코드는 바코드 출력 장수
   use_typ                 varchar2(1)      not null,     -- 사용구분(Y:사용,D:삭제)
   sub_yn                  varchar2(1)          null, 
   pat_typ_cd              varchar2(4)          null,     -- 상위코드
   etc1                    varchar2(10)         null,     -- 09:장비코드 - Slip(L01,L02) -- 사용하지 않음
                                                          -- 06:결과유형 - "AS": 항생제검사
                                                                         - "BB": 혈액불출
                                                             AS:항생제검사 - S : 하한미만
                                                                           - I : 평균치 
                                                                           - R : 상한초과
                                                             BB:혈액불출   - Y : 완료
                                                          -- 14:혈액제제 - 유효일수 
   etc2                    varchar2(10)         null,     -- 출력순서(2009.02.02 jangmc)
   fr_dt                   varchar2(12)     not null,     -- 시작일자
   to_dt                   varchar2(12)     not null      -- 종료일자
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
create unique index scccodem_ux1 on scccodem(hos_org_no,typ_cd,cd,pat_typ_cd,use_typ,fr_dt)
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
