--drop    table  scccoifm
;
create table scccoifm                                     -- ���� �˻� ���а� �˻� �ڵ� ����Ÿ
(
   hos_org_no              varchar2(8)      not null,     -- ���� ��ȣ 
   exam_cd                 varchar2(12)     not null,     -- �˻� cd
                                                          -- �� ù�ڸ��� L �̸� �˻� �ڵ�
   exam_typ                varchar2(4)          null,     -- LL : �Ϲ� �˻� �ڵ�
                                                          -- LD : Diff Count �ڵ�
                                                          -- LC : CBC �ڵ� 
                                                          -- LM : �̻��� ��� �ڵ�
                                                          -- LY : �̻��� ���� �ڵ�
                                                          -- LR : Micro ��� �˻�
                                                          -- LS : ��� �˻� 
                                                          -- LB : Anti-Body Test
                                                          -- LA : Abo typing  �˻�
                                                          -- RM : ���� �ڵ�
                                                          -- RA : �׻��� �ڵ�
                                                          -- RR : Micro �˻� ��� �ڵ�
   exam_nm                 varchar2(80)         null,     -- �˻� name
   exam_abbr               varchar2(30)         null,     -- �˻� ����
   gr_typ                  varchar2(1)          null,     -- �׷� �ڵ� ����
   trust_typ               varchar2(4)          null,     -- �˻�� ����(scccodem.typ_cd:'04')
   stat_cd                 varchar2(1)          null,     -- ��� �ڵ� ���� Y/N--- M   
                                                          -- sccststm
   ctn_typ                 varchar2(1)          null,     -- ���� �˻� ����
                                                          -- 'Y'/'N'
   spc_cd                  varchar2(4)          null,     -- ��ü �ڵ�(scccodem.typ_cd:'02')
   spc_smp_amt             varchar2(10)         null,     -- ��ü ä�뷮    
   spc_smp_unit            varchar2(3)          null,     -- ��ü ����
   spc_smp_typ             varchar2(4)          null,     -- ��ü ä���� ����
                                                          -- scccodem.typ_cd:'05'
   rslt_typ_cd             varchar2(4)          null,     -- �˻� ��� ���� (S)
                                                          -- scccodem.typ_cd:'06'
                                                          -- P:������� scccprfm 
   int_len                 varchar2(1)          null,     -- ������ ����
   dec_len                 varchar2(1)          null,     -- �Ҽ��� ����
   rslt_unit               varchar2(20)         null,     -- ��� ����
   nml_typ_cd              varchar2(4)          null,     -- ����ġ ���� (S) 
                                                          -- scccdm.typ_cd:'07'
   nml_m_max               varchar2(10)         null,     -- ����ġ ���� ����
   nml_f_max               varchar2(10)         null,     -- ����ġ ���� ����
   nml_m_min               varchar2(10)         null,     -- ����ġ ���� ����
   nml_f_min               varchar2(10)         null,     -- ����ġ ���� ����
   qc_typ_cd               varchar2(4)          null,     -- ���� ���� ����
                                                          -- scccdm.typ_cd:'08'
   dlt_max                 varchar2(10)         null,     -- delta ����
   dlt_min                 varchar2(10)         null,     -- delta ����
   pnc_max                 varchar2(10)         null,     -- panic ����
   pnc_min                 varchar2(10)         null,     -- panic ����
   exam_mach_cd            varchar2(4)          null,     -- �˻� ��� �ڵ�
                                                          -- scccodem.typ_cd:'09'
   qc_vld_trm              varchar2(4)          null,     -- ���� ���� ��ȿ�Ⱓ
   exam_cau_cd             varchar2(30)         null,     -- �˻� ���� ���� �ڵ�
                                                          -- SCL �˻��ڵ�
   exam_mth_cd             varchar2(4)          null,     -- �˻� ��� �ڵ�(scccdm.typ_cd:'11')
   exam_rslt_rept_yn       varchar2(1)          null,     -- �˻� ��� ���� ����
   lmt_trm_day             varchar2(2)          null,     -- �˻� �ҿ� �Ⱓ (��)
                                                          -- ������� ����
   lmt_trm_tm              varchar2(2)          null,     -- �˻� �ҿ� �Ⱓ (�ð�)
   lmt_trm_mint            varchar2(2)          null,     -- �˻� �ҿ� �Ⱓ (��)
   bs_exam_cd              varchar2(12)         null,     -- �⺻ �˻� �ڵ�
   bs_rslt_val             varchar2(30)         null,     -- �˻� �⺻ ��� ��  (S)
   user_empno              varchar2(12)     not null,     -- ����ڻ��
   use_typ                 varchar2(1)      not null,     -- ��뿩��(Y:���,D:����)
   fr_dt                   varchar2(12)     not null,     -- ��������
   to_dt                   varchar2(12)     not null,     -- ��������
   spc_bott                varchar2(4)          null,     -- ��ü ���
                                                          -- scccodem.typ_cd:'13'
   slip_knd                varchar2(3)          null,     -- Slip����
                                                          -- hzcbsdtm.cd_typ:'AST', hzcbsdtm.bas_cd:'C002'
   slip_sub                varchar2(4)          null,     -- Slip��������
                                                          -- hzcbsdtm.cd_typ:'AST', hzcbsdtm.bas_cd:'AB91'
   imm_trans_spc           varchar2(1)          null,     -- ������� ��ü(Y/N)
   reserve_spc             varchar2(1)          null,     -- ���� ��ü(Y/N)
   ptbs_app_item           varchar2(1)          null,     -- ȯ�ڱ⺻������ �����ؾ��� �׸�
                                                          -- VHI �� �ϳ��̰� ��������� etc1�� Y/N�� ��� �����
                                                          -- -: ����
                                                          -- V: �׻����˻翡��
                                                                VRE(vancomycin resistant enterococcus)�� ���� ȯ��
                                                          -- H: HIV ȯ��
                                                          -- I: �ŵ�,���� ȯ��
   blod_type               varchar2(5)          null,     -- ����Type
                                                          -- scccodem.typ_cd:'14'
   emer_able_cd            varchar2(5)          null,     -- ���ް��ɿ���
                                                          -- -: ����, 1:�ٹ��ð� �ܿ� �Ұ���, 2:��ü �Ұ���
   rcpt_info_cd            varchar2(5)          null,     -- �����/������ ����
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
