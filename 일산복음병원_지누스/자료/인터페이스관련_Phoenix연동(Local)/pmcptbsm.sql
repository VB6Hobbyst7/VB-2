--drop    table  pmcptbsm
;
create table pmcptbsm                                      -- ȯ����������
(
   hos_org_no              varchar2(10)     not null,      -- ���������ȣ
   pt_no                   varchar2(8)      not null,      -- ȯ�ڹ�ȣ
   pt_nm                   varchar2(30)     not null,      -- ȯ�ڼ���
   bth_y                   varchar2(2)      not null,      -- ����⵵
   ssn_1                   varchar2(6)      not null,      -- �ֹε�Ϲ�ȣ1
   ssn_2                   varchar2(7)      not null,      -- �ֹε�Ϲ�ȣ2
   pt_no_vld_yn            varchar2(1)      not null,      -- ȯ�ڹ�ȣ��ȿ���� (Y/N)
   tel_no_1                varchar2(3)      not null,      -- ��ȭ��ȣ1
   tel_no_2                varchar2(4)      not null,      -- ��ȭ��ȣ2
   tel_no_3                varchar2(4)      not null,      -- ��ȭ��ȣ3
   hp_no_1                 varchar2(3)          null,      -- �޴�����ȣ1
   hp_no_2                 varchar2(4)          null,      -- �޴�����ȣ2
   hp_no_3                 varchar2(4)          null,      -- �޴�����ȣ3
   mail_addr               varchar2(50)         null,      -- �����ּ�
   zpcd_1                  varchar2(3)      not null,      -- �����ȣ1
   zpcd_2                  varchar2(3)      not null,      -- �����ȣ2
   zpcd_seq                numeric(5)       not null,      -- �����ȣ����
   lw_addr                 varchar2(60)     not null,      -- �����ּ�
   ocp_cd                  varchar2(2)          null,      -- �����ڵ� (�ǹ����)
                                                           -- ��缱 �Կ����η� Ȱ��(Y/N)
   rgn_cd                  varchar2(2)          null,      -- �����ڵ� (�ǹ����)
                                                           -- pks20080814 : Net ȯ�ڿ���(Y)
   adms_yn                 varchar2(1)      not null,      -- ������� (Y:���, N:����ƴ�)
   indi_ucolt_yn           varchar2(1)      not null,      -- ���ι̼����� (Y/N)
   abo_blotyp              varchar2(2)          null,      -- abo������
   rh_blotyp               varchar2(2)          null,      -- rh������
                                                           -- cij 2008.09.30 varchar(1)?->varchar(2)
   fs_chos_ymd             varchar2(8)      not null,      -- ���ʳ�������
   lst_chos_ymd            varchar2(8)      not null,      -- ������������
   seri_no                 varchar2(10)         null,      -- ����ȯ�ڰ�����ȣ
                                                           -- Table�� �����.
   /* --------------------[�߰�]----------------------------------------------------------------- */
   pt_nat                  varchar2(6)          null,      -- ȯ�ڱ���(�ܱ����� ���)
                                                           -- �����ڵ� HZC-C010
   brth_ymd                varchar2(8)      not null,      -- �������
   brth_wo_flag            varchar2(1)      not null,      -- 0:���, 1:����
   wrk_ymd                 varchar2(8)      not null,      -- �۾�����
   wrk_tm                  varchar2(6)      not null,      -- �۾��ð�
   wrk_empno               varchar2(12)     not null,      -- �۾���

   /* --------------------------------------------------------------------------------------------
    * ���ܰ˻� ����Է½� ����
    * pks20090116 �߰�
    * -------------------------------------------------------------------------------------------- */
   vre_yn                  varchar2(1)          null,      -- �׻��� �˻翡�� VRE�� ���� ��ü(Y/N)
   hiv_yn                  varchar2(1)          null,      -- HIV �缺 ȯ��(Y/N)
   hpts_yn                 varchar2(1)          null       -- �ŵ� or ���� ���� ȯ��(Y/N)
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
