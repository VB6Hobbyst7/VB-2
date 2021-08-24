--drop    table  scrprexh
;
create table scrprexh                                     -- ���� ���̺� 
(
   hos_org_no              varchar2(8)     not null,      -- ���� ��ȣ
   smp_no                  varchar2(12)    not null,      -- ��ü��ȣ
   prcp_seq                numeric(2)      not null,      -- ��ü�� ó�� seq(mosxpslh's acp_no)
   exam_seq                numeric(2)      not null,      -- �˻��ڵ� seq
   parn_seq                numeric(2)      not null,      -- �����ڵ� seq 
   prcp_yn                 varchar2(1)     not null,      -- ó���ڵ忩��
   gr_sng_cd               varchar2(1)     not null,      -- G/S
   cd                      varchar2(20)    not null,      -- ó��,�˻�,����,�׻���,����ڵ�
   pt_no                   varchar2(8)     not null,      -- ȯ�� ��Ϲ�ȣ
   smp_stus                varchar2(1)     not null,      -- 0 : ���� 1 : ���� 2 : �Ϻ�
                                                          -- 3 : ���� 4 : ���� 6 : Interface 
   acp_dt                  varchar2(12)    not null,      -- ���� ����
   acp_empno               varchar2(12)    not null,      -- ������ ���
   rept_dt                 varchar2(12)        null,      -- ���� ����
   rept_empno              varchar2(12)        null,      -- ������ ���
   excel_yn                varchar2(1)         null,      -- ���ְ˻� Excel ���忩��
   scl_key                 varchar2(30)        null,      -- SCL ����KEY
   slip_knd                varchar2(3)         null,      -- Slip����
   wrk_no                  numeric(5)          null       -- �۾���ȣ
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
