--drop    table  scrrslth
;
create table scrrslth                                     -- �˻� ���
(
   hos_org_no              varchar2(8)     not null,      -- ���� ��ȣ
   smp_no                  varchar2(12)    not null,      -- ��ü��ȣ
   prcp_seq                numeric(2)      not null,      -- ��ü�� ó�� seq
   exam_seq                numeric(2)      not null,      -- �˻��ڵ� seq
   cd                      varchar2(20)    not null,      -- ó��,�˻�,����,�׻���,����ڵ�
   rept_seq                numeric(3)      not null,      -- ��¼���
   pt_no                   varchar2(8)     not null,      -- ȯ�ڹ�ȣ
   dla_stus                varchar2(2)     not null,      -- ��Ÿ ����
                                                          -- ù�ڸ�   ���� - N, ������  - Y
                                                          -- ��°�ڸ� ���� - N, High - H, Low - L
   pnc_stus                varchar2(2)     not null,      -- �д� ����
                                                          -- ù�ڸ�   ���� - N, ������  - Y
                                                          -- ��°�ڸ� ���� - N, High - H, Low - L
   exam_stus               varchar2(1)     not null,      -- �˻� ����
                                                          -- 0 : ���� 1 : ���� 2 : �Ϻ�
                                                          -- 3 : ���� 4 : ���� 6 : Interface
   exam_rslt               varchar2(100)       null,      -- �˻� ���
   exam_rslt_cd            varchar2(2)         null,      -- �˻� ��� �ڵ�
   mach_rslt               varchar2(80)        null,      -- ��� �˻� ���
   mach_rslt_cd            varchar2(2)         null,      -- ��� ��� �ڵ�
   exam_dt                 varchar2(12)        null,      -- �˻��Ͻ�
   exam_empno              varchar2(12)        null,      -- �˻��ڻ��
   etc1                    varchar2(10)        null,      -- size(sign) : 2009.01.06 jangmc
   etc2                    varchar2(10)        null       -- ��Ÿ2
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
