--drop    table  scccodem
;
create table scccodem                                     -- ���� �˻� ���а� �ڵ� ����Ÿ
(
   hos_org_no              varchar2(8)      not null,     -- ���� ��ȣ 
   typ_cd                  varchar2(4)      not null,     -- ����
                                                          -- 01 : �˻� �ڵ� ����
                                                          -- 02 : ��ü �ڵ� 
                                                          -- 03 : ���� �ڵ� ��Ī
                                                          -- 04 : �˻�� ����
                                                          -- 05 : ä���� ���� 
                                                          -- 06 : ��� ���� 
                                                                : ON,FT,AS - Hard Coding
                                                          -- 07 : ����ġ ���� 
                                                          -- 08 : ���� ���� ����
                                                          -- 09 : ��� �ڵ� 
                                                          -- 10 : ���� ���� 
                                                          -- 11 : �˻� ��� 
                                                          -- 12 : Diff Key Mapping
                                                          -- 13 : ��ü ���
                                                          -- 14 : �������� �з�
                                                          -- 15 : ������ �з�
   cd                      varchar2(10)     not null,     -- �ڵ�
   cd_desc                 varchar2(80)         null,     -- �ڵ� DESC 
   tribu_cd                varchar2(1)      not null,     -- �з�
                                                          -- '1':������
                                                          -- '2':������
   user_empno              varchar2(12)     not null,     -- ��ü �ڵ�� ���ڵ� ��� ���
   use_typ                 varchar2(1)      not null,     -- ��뱸��(Y:���,D:����)
   sub_yn                  varchar2(1)          null, 
   pat_typ_cd              varchar2(4)          null,     -- �����ڵ�
   etc1                    varchar2(10)         null,     -- 09:����ڵ� - Slip(L01,L02) -- ������� ����
                                                          -- 06:������� - "AS": �׻����˻�
                                                                         - "BB": ���׺���
                                                             AS:�׻����˻� - S : ���ѹ̸�
                                                                           - I : ���ġ 
                                                                           - R : �����ʰ�
                                                             BB:���׺���   - Y : �Ϸ�
                                                          -- 14:�������� - ��ȿ�ϼ� 
   etc2                    varchar2(10)         null,     -- ��¼���(2009.02.02 jangmc)
   fr_dt                   varchar2(12)     not null,     -- ��������
   to_dt                   varchar2(12)     not null      -- ��������
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
