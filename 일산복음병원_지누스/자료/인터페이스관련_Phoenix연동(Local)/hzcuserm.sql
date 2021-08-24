--drop    table  hzcuserm
;
create table hzcuserm                                     -- ����ڰ���
(
   hos_org_no              varchar2(10)     not null,     -- ���������ȣ
   empno                   varchar2(12)     not null,     -- ���
   fr_ymd                  varchar2(8)      not null,     -- ��������
   to_ymd                  varchar2(8)      not null,     -- ��������
   pwd                     varchar2(12)     not null,     -- ��ȣ
   empnm                   varchar2(20)     not null,     -- ����
   dpt_cd                  varchar2(6)      not null,     -- ������ڵ�(�μ��ڵ�)
   ocptyp                  varchar2(1)      not null,     -- ����
                                                          -- D:�ǻ�, N:��ȣ��, T:���,
                                                          -- P:���, E:�繫��
   spmed_yn                varchar2(1)          null,     -- �������Ῡ��(Y:��������,N:��������)
   lcns_no                 varchar2(6)          null,     -- �����ȣ
   pos                     varchar2(1)      not null,     -- �Ҽ� -> 'emc�������� �ǻ�(Y/-)'�� ����
                                                          -- ��ȣ��(����,�ܷ�,������,ȸ���ǡ�)
   use_yn                  varchar2(1)      not null,     -- ��뿩��(Y:���,N:����)
   use_dut                 varchar2(1)      not null,     -- ������
                                                          -- H : ���� R : �ɻ� M : ���� C : �濵
                                                          -- D : ó�� N : ��ȣ A : ���� 
   edt_dt                  varchar2(12)     not null,     -- �����Ͻ�
   user_empno              varchar2(12)     not null,     -- ����ڻ��
   tool_list               varchar2(900)    not null,     -- ToolBar ����Ʈ
   cert_yn                 varchar2(1)      not null,     -- ��������(Y/N)
   secure_pwd              varchar2(10)     not null,     -- ���Ⱦ�ȣ(����)
   sign_pwd                varchar2(10)     not null,     -- ������ȣ
   ssn_1                   varchar2(6)          null,     -- �ֹι�ȣ 1(�ڰ�Ȯ��)
   ssn_2                   varchar2(7)          null,     -- �ֹι�ȣ 2(�ڰ�Ȯ��)
   emc_send_cnt            numeric(5)       not null,     -- �����ǷἾŸ ���� ����
   emc_send_week           varchar2(1)      not null      -- �����ǷἾŸ ���� ����
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
