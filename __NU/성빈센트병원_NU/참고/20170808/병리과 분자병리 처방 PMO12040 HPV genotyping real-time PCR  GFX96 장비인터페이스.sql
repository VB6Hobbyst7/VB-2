/*    �ý��� database ���� ������ ������ ����� Ȯ�� ���� ���� �մϴ�.(�н������ ��ȣȭ ������ ���� �ʽ��ϴ�. �����ȣ������ Ȯ���� �մϴ�.)
      ����Ʈ ORACLE �α��� ���̵� / �н����� = lisu02 / lisu02 
1. Tnsname ���� ����

# ����Ʈ���� �(REAL)  ����ȭ 1ȣ�� ���н� 2ȣ��
REAL_HIS017_CTF=
    (DESCRIPTION=
        (ADDRESS_LIST=
            (ADDRESS=(PROTOCOL=TCP)(HOST=172.17.81.151)(PORT=1521))
            (ADDRESS=(PROTOCOL=TCP)(HOST=172.17.81.152)(PORT=1521))
            (FAILOVER=ON)
            (LOAD_BALANCE=OFF)
        )
        (CONNECT_DATA=
            (SERVICE_NAME=HIS017)
        )
    )

# ����Ʈ���� ����(TEST)
HIS017_EDU=
  (DESCRIPTION=
    (ADDRESS=
      (PROTOCOL=TCP)
      (HOST=172.17.102.183)
      (PORT=1525)
    )
    (CONNECT_DATA=
      (SERVICE_NAME=HIS017)
    )
  )
*/


/* -- ������ �Ҽ�  ����� �α��� ���� ��û�� ��� 20801950 / �����
param=['20801950']
����Ʈ���� ������ �����ڷῡ ��ϵ� �е鸸 ������� ��ȸ�˴ϴ�.
*/ 
select usab.instcd, -- �����ȣ 
       usab.userid, -- ����ھ��̵�
       usab.userabbr as usernm -- ����ڼ���
  from lis.lpcmusab usab,  -- ������ ����ڰ��� 
       com.zsumusrb usrb   -- �ý��� ���������
 where usab.instcd     = '017' 
   and usab.delflagcd  = '0'
   and usab.userid  = :arg_userid
   and usab.userid  = usrb.userid
   and to_char(sysdate,'yyyymmdd')    between usrb.userfromdd and usrb.usertodd

 

/* ==========
��Ʈ�� ����ó���� �κ� �����ּ���. nU�� ����Ŭ�� Full scan �ɰ�� DBA���� ���ܵ˴ϴ�. 
��Ʈ�� ����ó���� �κ� �����ּ���. nU�� ����Ŭ�� Full scan �ɰ�� DBA���� ���ܵ˴ϴ�.
��Ʈ�� ����ó���� �κ� �����ּ���. nU�� ����Ŭ�� Full scan �ɰ�� DBA���� ���ܵ˴ϴ�.
��Ʈ�� ����ó���� �κ� �����ּ���. nU�� ����Ŭ�� Full scan �ɰ�� DBA���� ���ܵ˴ϴ�.
��Ʈ�� ����ó���� �κ� �����ּ���. nU�� ����Ŭ�� Full scan �ɰ�� DBA���� ���ܵ˴ϴ�.
��Ʈ�� ����ó���� �κ� �����ּ���. nU�� ����Ŭ�� Full scan �ɰ�� DBA���� ���ܵ˴ϴ�.
��Ʈ�� ����ó���� �κ� �����ּ���. nU�� ����Ŭ�� Full scan �ɰ�� DBA���� ���ܵ˴ϴ�.
��Ʈ�� ����ó���� �κ� �����ּ���. nU�� ����Ŭ�� Full scan �ɰ�� DBA���� ���ܵ˴ϴ�.
==========*/ 


/* -- �������� ������ ������ȸ */
/* -- �������� ������ ������ȸ */
/* -- �������� ������ ������ȸ */
/* -- �������� ������ ������ȸ */
/* -- �������� ������ ������ȸ */
/* param=[017, to_char(sysdate-1,'yyyymmdd'), to_char(sysdate,'yyyymmdd')]
������ �������� ���������� ������ ������ ��ȸ�˴ϴ�. ��ȸȭ�� ó�� �Ǵ� �Ʒ����� ó���� ��� �ʿ��� �����Դϴ�.
*/ 
select /*+ leading(acpt) */
    acpt.instcd,        --�����ȣ
    acpt.acptdd,        --��������
    acpt.acptno,        --������ȣ
    acpt.acptitemno,    --�����׸��ȣ
    acpt.ptno,          --������ȣ
    acpt.pid,           --��Ϲ�ȣ
    acpt.testcd,        --�˻��ڵ�
    test.testengnm,     --�����˻��
    spcm.spcnm,         --��ü��
    acpt.prcpgenrflag,  --�Կ�/�ܷ�����
    acpt.orddeptcd,     --������ڵ�
    dept.deptengabbr,   --�������
    acpt.prcpdd,        --ó������,
    acpt.execprcpuniqno, --�ǽ�ó�����Ϲ�ȣ
    acpt.prcpno,        --ó���ȣ
    ptbs.hngnm,         --ȯ�ڸ�
    ptbs.sex,           --����
    ptbs.brthdd,        --����
    com.fn_zz_getage(ptbs.rrgstno1, ptbs.rrgstno2, acpt.acptdd, 'A', ptbs.brthdd) as age  -- �������ڱ��� ����
from lis.lpjmacpt acpt, lis.lpcmtest test, lis.lpcmspcm spcm, pam.pmcmptbs ptbs, com.zsdddept dept
where acpt.instcd = :arg_instcd --'����Ʈ����(017)�� �����Դϴ�.
and acpt.acptdd between  :arg_acptfromdd and :arg_acpttodd 
and acpt.testcd = 'PMO12040' -- �˻� ó���ڵ� PMO12040 ���� GFX96 ���� �ٸ�ó�� ó���� in ���� ����
and acpt.acptstatcd = '0' --���������ڵ�(0:����,1:���,2:������,3:�������, 4:������� ����
and acpt.instcd          = test.instcd
and acpt.testcd          = test.testcd
and acpt.instcd          = spcm.instcd
and acpt.spccd           = spcm.spccd
and acpt.instcd          = ptbs.instcd
and acpt.pid             = ptbs.pid
and acpt.instcd          = dept.instcd
and acpt.orddeptcd       = dept.deptcd
and acpt.prcpdd between dept.valifromdd and dept.valitodd




/* -- ��������� ó�� ���� */
/* -- ��������� ó�� ���� */
/* -- ��������� ó�� ���� */
/* -- ��������� ó�� ���� */
/* -- ��������� ó�� ���� */
/* -- ��������� ó�� ���� */
/* -- ��������� ó�� ���� ������ */
/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml setseqnolock 
param=[017, 4] */
update lis.lpcmseqn
       set lastgenrno = 1
 where instcd        = :arg_instcd   
   and seqgenryy = '1900'
   and seqflagcd  = :arg_seqflagcd
        
/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getlastseqno 
param=[017, 2017, 4]*/
select coalesce(lastgenrno+1, 1) as lastgenrno
  from lis.lpcmseqn
 where instcd     = :arg_instcd
   and seqgenryy  = :arg_seqgenryy 
   and seqflagcd  = :arg_seqflagcd

/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml setseqnogenr 
param=[32715, 10602673, 017, 2017, 4]*/
update lis.lpcmseqn
   set lastgenrno = :arg_lastgenrno ,
          lastupdtdt   = sysdate,
          lastupdtrid  = :arg_userid 
 where instcd        = :arg_instcd 
   and seqgenryy = :arg_seqgenryy
   and seqflagcd  = :arg_seqflagcd 
        
/* himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml instestrslt 
param=[M17003176, 20170724, 32715, 017, 
17488137, 142613, null, null, null, 
HPV High Risk Type : Positive (18+, 68+, 31+++)
HPV Low  Risk Type : Positive (70+, 61+) , 
null, 
null, null, ������ü other, HPV genotyping real-time PCR, null, 0, 
0, 0, 0, 0, 0, 0, null, 
10602673, 
null, 0, null, null, null, 
null, null, 10602673, 
null, null, 10602673, null]
*/
insert into lis.lprmrslt (ptno,       rsltrgstdd,    rsltrgstno,     rsltrgsthistno, instcd,
                          pid,        rsltrgsttm,    grostestrecdd,  grostestrectm,  grostestrecid,    
                          diagcnts,
                          readdd,     readtm,        readid,         extrpartcnts,   extrmthdcnts,     diagcd,
                          spckeepflagcd, rslthideflagcd, cncrjudgflagcd, conccaseflagcd, preprsltflagcd, ugcyalertflagcd, cnstcd, 
                          rsltrgstid, cnclflagcd,    cnclresncd,     cncldd,         cncltm,
                          grospic,    keybloc,       tissbloct,      tissblocnt,     readgrade,
                          cnclid,     delflagcd, 
                          fstrgstdt,  fstrgstrid,
                          lastupdtdt, lastupdtrid)
                  values (:arg_ptno,       :arg_rsltrgstdd,    :arg_rsltrgstno,     1,        :arg_instcd,
                          :arg_pid,        :arg_rsltrgsttm,    :arg_grostestrecdd,  :arg_grostestrectm,  :arg_grostestrecid,    
                          :arg_diagcnts,
                          :arg_readdd,     :arg_readtm,        :arg_readid,         :arg_extrpartcnts,   :arg_extrmthdcnts,     :arg_diagcd,
                          :arg_spckeepflagcd, :arg_rslthideflagcd, :arg_cncrjudgflagcd, :arg_conccaseflagcd, :arg_preprsltflagcd, :arg_ugcyalertflagcd, :arg_cnstcd, 
                          CASE WHEN :arg_cellrsltrgstid IS NULL THEN :arg_userid ELSE :arg_cellrsltrgstid END ,       '-',       '-',       '-',       '-',
                          :arg_grospic,    :arg_keybloc,       :arg_tissbloct,      :arg_tissblocnt,     :arg_readgrade,
                          '-',       '0', 
                          sysdate,  case when :arg_cellrsltrgstid is null then :arg_userid else :arg_cellrsltrgstid end ,
                          sysdate,  case when :arg_cellrsltrgstid is null then :arg_userid else :arg_cellrsltrgstid end )
        
/* himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml instestrsltcnts 
param=[
M17003176, 
20170724, 
32715, 
017, 
17488137, 
[Methods]
   Seegene HPV Real-time PCR (Anyplex II HPV 28 Detection Real-time PCR)

[Result]
HPV High Risk Type : POSITIVE (18+, 68+, 31+++)
HPV Low  Risk Type : POSITIVE (70+, 61+) , null, Adequate, null, null, null, 3. Comment
�������� ���̷��� (Human papilloma virus)�� �ڱð�ξ��� �ֿ� �������ڷ� �˷��� �ֽ��ϴ�. �ϰ��� ���ü� ������ ���� �����豺 (high risk)�� �����豺 (low risk)�� ���еǸ�, �����豺 HPV�� �밳 �ð��� ������ �ҽǵǰų� �縶�� ���� �缺������ ������ �Ǵ� �ݸ�, �����豺�� �ڱð�ξ��� ���߽�Ű�µ� �����մϴ� (N Engl J Med. 2003 348:518).

�� �� ��ǰ�� 19���� �����豺 HPV (16, 18, 26, 31, 33, 25, 29, 45, 51, 52, 53, 56, 58, 59, 66, 68, 73, 82��)�� 9���� �����豺 HPV (6, 11, 40, 42, 43, 44, 54, 61, 70��), ���δ������� Ÿ���ٻ��� �����մϴ�. 
�� Viral load���� +++:10^5 copies/reaction, ++:10^5~10^2 copies/reaction, +:10^2 copies/reaction�� �󵵷� �ؼ��� �� �ֽ��ϴ�. �� �� ��+���� �ſ� ���� �󵵷� ���� �ñ�, ��ü ä�� ���¿� ���� �ݺ� �˻� �� �������� ���� �� �ֽ��ϴ�.
�� PCR �˻�� ��ü �� �ռ��� ���ų� �������� ��ü �Ǽ� �Ǵ� ���� ���������� �����ϴ� ��� �������� ���� �� �ֽ��ϴ�. ����, PCR �˻�� ������ ������ �˻��ϹǷ� �����հ� ����� ������ �ȵǾ� ���缺�� ���ɼ��� �ֽ��ϴ�. ��� �ؼ� �� �ӻ� ���� �������� �Ǵ��Ͻñ� �ٶ��ϴ�.
�� ��� �˻�� �˻� ���, �þ��� �������� �� �˻� ����� ������ �����ǿ� ���� Ȯ�εǾ����ϴ�.
   (�˻� �����: �����), 
10602673, 
10602673]
*/
insert into lis.lprmcnts (ptno, rsltrgstdd, rsltrgstno, rsltrgsthistno, instcd, pid,  
                          rsltcnts1,  rsltcnts2,  rsltcnts3,    
                          rsltcnts4,  rsltcnts5,  rsltcnts6,
                          cmtcnts,    delflagcd,
                          fstrgstdt,  fstrgstrid,
                          lastupdtdt, lastupdtrid)
                   values(:arg_ptno, :arg_rsltrgstdd, :arg_rsltrgstno, :arg_rsltrgsthistno, :arg_instcd, :arg_pid,  
                          :arg_rsltcnts1,  :arg_rsltcnts2,  :arg_rsltcnts3,    
                          :arg_rsltcnts4,  :arg_rsltcnts5,  :arg_rsltcnts6,
                          :arg_cmtcnts,    '0',
                          sysdate, :arg_fstrgstrid,
                          sysdate, :arg_lastupdtrid)

            

/* himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml updlastdiag 
param=[null, null, null, 
0, 
10602673, 
017, 
M17003176, 
17488137] */      
            update lis.lprmrslt 
               set readdd     = :arg_readdd,          readtm      = :arg_readtm,   readid = :arg_readid, 
                   cnclflagcd = '-',               cnclresncd  = '-',  
                   cncldd     = '-',               cncltm      = '-',        cnclid = '-',
                   ugcyalertflagcd = :arg_ugcyalertflagcd,
                   lastupdtdt = sysdate, lastupdtrid = :arg_lastupdtrid
             where instcd         = :arg_instcd
               and ptno           = :arg_ptno
               and pid            = :arg_pid
               and rsltrgsthistno = 1
               and delflagcd      = '0'       
        
/* himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml updlpcmpnis  
param=[������ü other, 
HPV genotyping real-time PCR, 
HPV High Risk Type : Positive (18+, 68+, 31+++)
HPV Low  Risk Type : Positive (70+, 61+) , 
null, 
10602673, 
017, 
M17003176] */
            update lis.lpcmpnis
               set extrpartcnts = :arg_extrpartcnts,
                   extrmthdcnts = :arg_extrmthdcnts,
                   diagcnts     = :arg_diagcnts,
                   diagcd       = :arg_diagcd,
                   lastupdtdt  = sysdate,
                   lastupdtrid = :arg_lastupdtrid
             where instcd    = :arg_instcd
               and ptno      = :arg_ptno
               and delflagcd = '0'

/* himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml updexersltcomfirm 
param=[
N, 
10602673, 
017, 
M17003176, 
17488137, 
20170724]*/        
update lis.lpjmacpt
   set rsltstatcd = nvl(:arg_rsltstatcd, 'Y')
     , lastupdtrid = :arg_lastupdtrid
     , lastupdtdt = sysdate
 where instcd = :arg_instcd
   and ptno   = :arg_ptno
   and pid    = :arg_pid
   and acptdd = nvl(:arg_acptdd, acptdd)

/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getReqHospYn         
param=[017, 
M17003176]*/
select 
         acpt.ptno
        ,acpt.testflagcd
        ,cnts.cmtcnts 
        ,(select cdnm 
             from com.zbcmcode code
           where code.cdgrupid = 'A0607'
                and code.cdid = acpt.instcd) as hospnm
        ,nvl((select cdnm 
                from com.zbcmcode code
               where code.cdgrupid = 'A0607'
                 and code.cdid = acpt.reqinstcd),
             (select coophospnm 
                from ast.arhmchsp chsp
               where chsp.instcd = acpt.instcd
                 and chsp.coophospcd = acpt.reqinstcd)) as coophospnm
         ,test.workflagcd        
    from lis.lpjmacpt acpt
         ,lis.lprmcnts cnts
         ,lis.lpcmtest test
    where acpt.instcd = cnts.instcd
        and acpt.ptno   = cnts.ptno
        and acpt.instcd = test.instcd
        and acpt.testcd = test.testcd
        and acpt.instcd = :arginstcd
        and acpt.ptno   = :arg_ptno
    


/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getptnoprcpinfo 
param=[53, 53, 53, 
53, 53, 53, 
53, 53, 53, 
53, 53, 
53, 
017, 
M17003176, 
17488137]
*/
          SELECT acpt.instcd, acpt.prcpdd, acpt.pid, acpt.prcpno, acpt.execprcpuniqno,
                   MIN(acpt.prcpgenrflag) AS prcpgenrflag,  'I' AS biztretflagcd,  
                   CASE WHEN ''||:arg_scrno||'' = '32' THEN '700' 
                        WHEN ''||:arg_scrno||'' = '52' THEN '740' 
                        WHEN ''||:arg_scrno||'' = '53' THEN '700' 
                   END AS newprcpstatcd, 
                   CASE WHEN ''||:arg_scrno||'' = '32' THEN '700' 
                        WHEN ''||:arg_scrno||'' = '52' THEN '740' 
                        WHEN ''||:arg_scrno||'' = '53' THEN '700' 
                   END AS bizflagcd, 
                   CASE WHEN ''||:arg_scrno||'' = '32' THEN '700' 
                        WHEN ''||:arg_scrno||'' = '52' THEN '740' 
                        WHEN ''||:arg_scrno||'' = '53' THEN '700' 
                   END AS tretflagcd, 
                   CASE WHEN :arg_scrno = '32' THEN MAX(pnis.makeenddd) ELSE TO_CHAR(SYSDATE,'YYYYMMDD') END AS tretdd,
                   CASE WHEN :arg_scrno = '32' THEN MAX(pnis.makeendtm) ELSE TO_CHAR(SYSDATE,'HH24MISS') END AS trettm,
                   :arg_scrno AS scrno
                 , acpt.prcpgenrflag AS prcpgenrflagcd
              FROM lis.lpjmacpt acpt
                 , lis.lpcmpnis pnis
             WHERE acpt.instcd      = :arg_instcd
               AND acpt.ptno        = :arg_ptno
               AND acpt.pid         = :arg_pid
               AND acpt.acptstatcd IN ('0', '2', '3', '4', '9')
               AND acpt.instcd      = pnis.instcd
               AND acpt.ptno        = pnis.ptno
               AND pnis.delflagcd   = '0'
             GROUP BY acpt.instcd, acpt.prcpdd, acpt.pid, acpt.prcpno, 
                      acpt.execprcpuniqno, acpt.prcpgenrflag
        
/* himed/his/lis/lib/plgyprcpcommonmgt/dao/sqls/plgyprcpcommondao_sqls.xml getprcpstat 
param=[017, 17488137, 20170724, 1151787391]*/
        select prcpstatcd 
        from (
            SELECT b.prcpstatcd
              FROM emr.mmodexip a, emr.mmohiprc b   -- �Կ�
             WHERE a.instcd         = :arg_instcd
               AND a.pid            = :arg_pid
               AND a.prcpdd         = :arg_prcpdd
               AND a.execprcpuniqno = :arg_execprcpuniqno
               AND a.execprcphistcd = 'O'
               AND a.instcd         = b.instcd
               AND a.pid            = b.pid
               AND a.prcpdd         = b.prcpdd
               AND a.prcpno         = b.prcpno
               AND a.prcphistno     = b.prcphistno
               AND b.prcphistcd     = 'O'
               AND b.prcpclscd      = 'D2'
               AND b.tempprcpflag   = 'N'
              union all
            SELECT b.prcpstatcd
              FROM emr.mmodexop a, emr.mmohoprc b   -- �ܷ�
             WHERE a.instcd         = :arg_instcd
               AND a.pid            = :arg_pid
               AND a.prcpdd         = :arg_prcpdd
               AND a.execprcpuniqno = :arg_execprcpuniqno
               AND a.execprcphistcd = 'O'
               AND a.instcd         = b.instcd
               AND a.pid            = b.pid
               AND a.prcpdd         = b.prcpdd
               AND a.prcpno         = b.prcpno
               AND a.prcphistno     = b.prcphistno
               AND b.prcphistcd     = 'O'
               AND b.prcpclscd      = 'D2'
               AND b.tempprcpflag   = 'N' )
             where rownum = 1
        
/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml setprcpstat  
param=[710, 10602673, 017, 17488137, 20170724, 1151787391]
*/
--�Կ��ϰ��
            UPDATE 
                   emr.mmohiprc
               SET prcpstatcd  = :arg_prcpstatcd,
                   lastupdtdt  = SYSDATE,
                   lastupdtrid = :arg_lastupdtrid
             WHERE (instcd, pid, prcpdd, prcpno, prcphistno) IN
                   (SELECT instcd, pid, prcpdd, prcpno, prcphistno
                      FROM emr.mmodexip
                     WHERE instcd         = :arg_instcd
                       AND pid            = :arg_pid
                       AND prcpdd         = :arg_prcpdd
                       AND execprcpuniqno = :arg_execprcpuniqno
                       AND execprcphistcd = 'O'
                   )
               AND prcphistcd   = 'O'
               AND prcpclscd    = 'D2'
               AND tempprcpflag = 'N'

--�ܷ��ϰ��
            UPDATE 
                   emr.mmohoprc
               SET prcpstatcd  = :arg_prcpstatcd,
                   lastupdtdt  = SYSDATE,
                   lastupdtrid = :arg_lastupdtrid
             WHERE (instcd, pid, prcpdd, prcpno, prcphistno) IN
                   (SELECT instcd, pid, prcpdd, prcpno, prcphistno
                      FROM emr.mmodexop
                     WHERE instcd         = :arg_instcd
                       AND pid            = :arg_pid
                       AND prcpdd         = :arg_prcpdd
                       AND execprcpuniqno = :arg_execprcpuniqno
                       AND execprcphistcd = 'O'
                   )
               AND prcphistcd   = 'O'
               AND prcpclscd    = 'D2'
               AND tempprcpflag = 'N'
               
/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml setexecprcpstat      
param=[710, 10602673, 017, 17488137, 20170724, 1151787391]
*/
--�Կ��ϰ��
            UPDATE 
                   emr.mmodexip a
               SET a.execprcpstatcd = :arg_execprcpstatcd,
                   a.lastupdtdt     = SYSDATE,
                   a.lastupdtrid    = :arg_lastupdtrid
             WHERE a.instcd         = :arg_instcd
               AND a.pid            = :arg_pid
               AND a.prcpdd         = :arg_prcpdd
               AND a.execprcpuniqno = :arg_execprcpuniqno
               AND a.execprcphistcd = 'O'

--�ܷ��ϰ��
            UPDATE 
                   emr.mmodexop a
               SET a.execprcpstatcd = :arg_execprcpstatcd,
                   a.lastupdtdt     = SYSDATE,
                   a.lastupdtrid    = :arg_lastupdtrid
             WHERE a.instcd         = :arg_instcd
               AND a.pid            = :arg_pid
               AND a.prcpdd         = :arg_prcpdd
               AND a.execprcpuniqno = :arg_execprcpuniqno
               AND a.execprcphistcd = 'O'

/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getDoPrcpTret 
param=[017, 1151787391, 20170724, 710]
*/
SELECT COUNT(prcpdd) AS tretcnt
  FROM emr.mmodexpt
 WHERE instcd         = :arg_instcd
   AND execprcpuniqno = :arg_execprcpuniqno
   AND prcpdd         = :arg_prcpdd
   AND tretflagcd     = :arg_tretflagcd
        
/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml insDoPrcpTret 
param=[20170724, 1151787391, 
710, 017, 
20170724, 142613, null, 10602673, null, 
10602673, 
10602673] */
INSERT INTO emr.mmodexpt (prcpdd,       execprcpuniqno, 
                          tretflagcd,   instcd,            
                          tretdd,       trettm,    tretpsnid, 
                          fstrgstrid,   fstrgstdt, 
                          lastupdtrid,  lastupdtdt)
                  VALUES (:arg_prcpdd,      CAST(:arg_execprcpuniqno AS INTEGER), 
                          :arg_tretflagcd,   :arg_instcd,            
                          :arg_tretdd,       :arg_trettm,    CASE WHEN :arg_cnfmid IS NULL THEN :arg_userid ELSE :arg_cnfmid END , 
                          :arg_userid,   SYSDATE, 
                          :arg_userid,   SYSDATE)

/* himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml getptnoinfo         
param=[017, M17003176, 17488137]
*/
             SELECT acpt.instcd, acpt.pid, acpt.ptno, 
                   acpt.prcpgenrflag,     acpt.workflagcd,
                   NVL(acpt.readdd, '-') AS readdd, 
                   NVL(acpt.readtm, '-') AS readtm, 
                   NVL(acpt.readdrid,'-') AS readdrid, acpt.diagnm,
                   TRIM(COALESCE(usrb.usernm, '')) AS reptdrnm,
                   CASE WHEN acpt.workflagcd = '1001' THEN acpt.grosrslt1
                        WHEN acpt.workflagcd = '1002' THEN acpt.grosrslt1
                        ELSE NULL
                   END AS grosrslt1
              FROM (
                    SELECT acpt.instcd, acpt.pid, acpt.ptno, 
                           acpt.prcpgenrflag, ptnf.workflagcd,
                           acpt.readdd,   acpt.readtm,
                           pnis.readdrid, pnis.diagcnts AS diagnm,
                           NVL(cnts.rsltcnts1,' ') AS grosrslt1
                      FROM lis.lpjmacpt acpt, lis.lpcmptnf ptnf, lis.lpcmpnis pnis,
                           lis.lprmcnts cnts
                     WHERE acpt.instcd      = :arg_instcd
                       AND acpt.ptno        = :arg_ptno
                       AND acpt.pid         = :arg_pid
                       AND acpt.acptstatcd IN ('0', '2', '3', '4', '9')
                       AND acpt.instcd      = ptnf.instcd
                       AND acpt.ptnocd      = ptnf.ptnocd
                       AND acpt.instcd      = pnis.instcd
                       AND acpt.ptno        = pnis.ptno
                       AND acpt.pid         = pnis.pid
                       AND pnis.delflagcd   = '0'
                       AND acpt.instcd      = cnts.instcd
                       AND acpt.ptno        = cnts.ptno
                       AND acpt.pid         = cnts.pid
                       AND cnts.rsltrgsthistno = 1
                       AND ROWNUM              = 1
                   ) acpt LEFT OUTER JOIN com.zsumusrb usrb
                ON acpt.readdrid = usrb.userid
               AND acpt.readdd   BETWEEN usrb.userfromdd AND usrb.usertodd
          
/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml gettestrsltinfo 
param=[017, M17003176, 017, M17003176, 017, M17003176, 17488137, 017, M17003176, 17488137]
*/
            SELECT a.instcd, a.ptno, a.pid,
                   SUBSTR(a.ptno,1,LENGTH(a.ptno)-6)||'-'||
                   SUBSTR(a.ptno,  LENGTH(a.ptno)-5, 6) AS dispptno,
                   a.grostestrecdd,  a.grostestrectm, a.grostestrecid,
                   cnts.rsltcnts1,   cnts.rsltcnts2,  cnts.rsltcnts3, 
                   cnts.rsltcnts4,   cnts.rsltcnts5,  cnts.rsltcnts6,   cnts.cmtcnts,
                   NVL(a.readdd, '-') AS readdd,         a.spckeepflagcd, a.conccaseflagcd, a.preprsltflagcd, a.ugcyalertflagcd,
                   a.rslthideflagcd, a.grospic,       a.tissbloct,      a.tissblocnt,     a.keybloc,
                   CASE WHEN a.preprsltflagcd = '1'THEN '�������Դϴ�.' ELSE '' END AS preprsltcnts
                 , a.readtm
                 , a.readid
                 , COALESCE(
                       (SELECT COUNT(cssd.csteno) 
                          FROM lis.lpbmcssd cssd
                         WHERE cssd.instcd    = :arg_instcd
                           AND cssd.ptno      = :arg_ptno
                           AND cssd.csteno    > '000'
                           AND cssd.slidno    = '000'
                           AND cssd.spchistno = 1
                       ), 0
                   ) AS cstecnt
                 , COALESCE(
                       (SELECT COUNT(cssd.csteno) 
                          FROM lis.lpbmcssd cssd
                         WHERE cssd.instcd    = :arg_instcd
                           AND cssd.ptno      = :arg_ptno
                           AND cssd.slidno    > '000'
                           AND cssd.spchistno = 1
                       ), 0
                   ) AS slidcnt
              FROM lis.lprmrslt a, lis.lprmcnts cnts
             WHERE a.instcd         = :arg_instcd
               AND a.ptno           = :arg_ptno
               AND a.pid            = :arg_pid
               AND a.rsltrgsthistno = 1
               AND a.delflagcd      = '0'    
               AND a.instcd         = cnts.instcd
               AND a.ptno           = cnts.ptno
               AND a.rsltrgstdd     = cnts.rsltrgstdd
               AND a.rsltrgstno     = cnts.rsltrgstno
               AND a.rsltrgsthistno = cnts.rsltrgsthistno
               AND a.pid            = cnts.pid
               AND cnts.delflagcd   = '0'
             UNION ALL
            SELECT bfpa.instcd, bfpa.ptno, bfpa.pid,
                   SUBSTR(bfpa.ptno,1,LENGTH(bfpa.ptno)-6)||'-'||
                   SUBSTR(bfpa.ptno,  LENGTH(bfpa.ptno)-5, 6) AS dispptno,
                   bfpa.diagdd AS grostestrecdd, TO_CHAR(SYSDATE, 'HH24MISS') AS grostestrectm, '-' AS grostestrecid,
                   TO_CLOB(' ') AS rsltcnts1, TO_CLOB(' ') AS rsltcnts2, bfrt.rsltcnts AS rsltcnts3,
                   TO_CLOB(' ') AS rsltcnts4, TO_CLOB(' ') AS rsltcnts5, TO_CLOB(' ') AS rsltcnts6,  TO_CLOB(' ') AS cmtcnts,
                   bfpa.diagdd AS readdd, '0' AS spckeepflagcd, '0' AS conccaseflagcd, '0' AS preprsltflagcd, '0' AS ugcyalertflagcd,
                   '0' AS rslthideflagcd, ' ' AS grospic,       ' ' AS tissbloct,      ' ' AS tissblocnt,     ' ' AS keybloc,
                   ' ' AS preprsltcnts
                 , '-' AS readtm
                 , '-' AS readid
                 , 0 AS cstecnt
                 , 0 AS slidcnt
              FROM lis.lprmbfpa bfpa, lis.lprmbfrt bfrt
             WHERE bfpa.instcd     = :arg_instcd
               AND bfpa.ptno       = :arg_ptno
               AND bfpaa.pid       = :arg_pid
               AND bfpa.rgsthistno = 1
               AND bfpa.instcd     = bfrt.instcd
               AND bfpa.ptno       = bfrt.ptno
               AND bfpa.rgstdd     = bfrt.rgstdd
               AND bfpa.rgstno     = bfrt.rgstno
               AND bfrt.rgsthistno = 1
        

/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getpatinfo 
param=[017, M17003176, 017, M17003176, 17488137, 017, M17003176, 017, M17003176]
*/
            SELECT a.instcd, a.ptno, a.pid, a.patnm
                 , a.acptdd, 
                   REPLACE(a.readdd, '-', '') AS readdd,
                   a.rrgstno, a.sexage, a.grosdrid,  a.spcnm,
                   TRIM(COALESCE(c.usernm,'-')) AS grosdrnm,
                   TRIM(COALESCE(d.usernm,'-')) AS readdrnm,
                   SUBSTR(a.ptno,1,LENGTH(a.ptno)-6)||'-'||
                   SUBSTR(a.ptno,  LENGTH(a.ptno)-5, 6) AS dispptno,
                   a.workflagcd, a.spckeepflagcd, a.grostestdd,   a.grostesttm
                 , a.acpttm
                 , NVL((SELECT usrb.usernm
	                      FROM com.zsumusrb usrb
	                     WHERE usrb.userid = a.acptid
                     	   AND a.acptdd BETWEEN usrb.userfromdd AND usrb.usertodd 
                        ), '-') AS acptnm
                 , a.diagcnts
                 , a.prcpdd
                 , a.orddeptcd
                 , a.testflagcd
                 , a.coophospnm
                 , a.statsworkflagcd
                 , COALESCE(
                       (SELECT COUNT(cssd.csteno) 
                          FROM lis.lpbmcssd cssd
                         WHERE cssd.instcd    = :arg_instcd
                           AND cssd.ptno      = :arg_ptno
                           AND cssd.csteno    > '000'
                           AND cssd.slidno    = '000'
                           AND cssd.spchistno = 1
                       ), 0
                   ) AS cstecnt
                 , a.orddd
                 , a.orddrid
                 , a.fstrgstrid
                 , a.prcpgenrflag
                 , a.delivedd
                 , a.delivetm
                 , a.delivenm
                 , fstreaddrid
                 , decode(fstreaddrid, null, null, com.fn_zs_getusernm(fstreaddrid, to_char(sysdate, 'yyyymmdd')))   as fstreaddrnm  -- 201312 1���ǵ��� �߰�.
                 , rsltstatcd
                 , a.prcpno -- [SR20150708000187] hamtn 20150727 ��Ź�˻� �߰�������ſ��� ��Ź�Ƿڵ����� ���°� ������ ���� �߰�
              FROM (
                    SELECT a.instcd, a.ptno, a.pid, b.hngnm AS patnm, acpt.readdd, acpt.readtm,
                           b.rrgstno1||'-'||b.rrgstno2 AS rrgstno,
                           com.fn_zz_getsex(b.rrgstno1, b.rrgstno2, '2')||'/'||
                           com.fn_zz_getage(b.rrgstno1, b.rrgstno2, a.acptdd, 'A', '-') AS sexage,
                           a.grosdrid, a.readdrid, c.spcnm, ptnf.workflagcd,
                           a.spckeepflagcd, a.grostestdd,   a.grostesttm
                         , acpt.acptdd
                         , acpt.acpttm
                         , acpt.acptid
                         , a.diagcnts
                         , acpt.prcpdd
                         , acpt.orddeptcd
                         , acpt.testflagcd
                         , acpt.coophospnm
                         , acpt.statsworkflagcd
                         , acpt.orddd
                         , acpt.orddrid
                         , acpt.fstrgstrid
                         , acpt.prcpgenrflag
                         , prtn.delivedd
                         , prtn.delivetm
                         , (SELECT usrb.usernm
                              FROM com.zsumusrb usrb
                             WHERE usrb.userid = prtn.deliveid
                               AND acpt.prcpdd BETWEEN usrb.userfromdd AND usrb.usertodd
                           ) AS delivenm
                         , a.fstreaddrid as fstreaddrid   -- 201312 1���ǵ��� �߰�.
                         , rsltstatcd    as rsltstatcd     -- 201402 ���ο��� �߰�
                         , acpt.prcpno -- [SR20150708000187]  20150727 ��Ź�˻� �߰�������ſ��� ��Ź�Ƿڵ����� ���°� ������ ���� �߰�
                      FROM lis.lpcmpnis a, pam.pmcmptbs b, lis.lpcmspcm c, lis.lpcmptnf ptnf
                           ,(SELECT acpt.instcd
                                  , acpt.acptdd
                                  , acpt.acpttm
                                  , acpt.acptid
                                  , acpt.prcpdd
                                  , acpt.orddeptcd
                                  , acpt.testflagcd
                                  , NVL((SELECT coophospnm
    				                       FROM ast.arhmchsp
    				                      WHERE instcd     = acpt.instcd
    				                        AND deldd      = '00000000'
    				                        AND coophospcd = acpt.reqinstcd
    				                        AND pathoyn    = 'Y'
    				                    ), DECODE(acpt.reqinstcd, '-' , '', (SELECT basecdnm
                                                                              FROM lis.llfmbasc
                                                                             WHERE basecdid = 'LL058'
                                                                               AND basecd2  = acpt.reqinstcd))) AS coophospnm
				                  , acpt.statsworkflagcd
				                  , acpt.orddd
                                  , acpt.orddrid
                                  , acpt.fstrgstrid
                                  , acpt.prcpgenrflag
                                  , acpt.trandd
                           		  , acpt.tranno
                           		  , acpt.prcpno
                                  , acpt.readdd
                                  , acpt.readtm
                                  , acpt.rsltstatcd
                               FROM (SELECT acpt.instcd
                                          , acpt.acptdd
                                          , acpt.acpttm
                                          , acpt.acptid
                                          , acpt.prcpdd
                                  		  , acpt.orddeptcd
                                  		  , acpt.reqinstcd
                                  		  , acpt.testflagcd
                                  		  , test.statsworkflagcd
                                  		  , acpt.orddd
                                  		  , acpt.orddrid  
                                  		  , acpt.fstrgstrid
                                  		  , acpt.prcpgenrflag
                                  		  , acpt.trandd
                                  		  , acpt.tranno
                                  		  , acpt.prcpno
                                          , rslt.readdd 
                                          , rslt.readtm
                                          , acpt.rsltstatcd
                                       FROM lis.lpjmacpt acpt, lis.lpcmtest test, lis.lprmrslt rslt 
                                      WHERE acpt.instcd  = :arg_instcd
                                        AND acpt.ptno    = :arg_ptno
						            	 AND acpt.pid     = :arg_pid
                                        AND acpt.acptstatcd IN ('0', '2', '3', '4', '9')
                                        AND acpt.instcd   = test.instcd
                                        AND acpt.testcd   = test.testcd         
                                        AND  acpt.instcd      = rslt.instcd(+)
                                        AND acpt.ptno        = rslt.ptno(+)
                                        AND acpt.pid         = rslt.pid(+)
                                        AND rslt.rsltrgsthistno(+) = '1'   
                                        ORDER BY DECODE(NVL(acpt.trandd, '-'), '-', '9', acpt.trandd), acpt.fstrgstdt
                                    ) acpt
                             WHERE ROWNUM = 1
                            ) acpt LEFT OUTER JOIN lis.lpjmprtn prtn
		                                        ON acpt.instcd      = prtn.instcd
		                                       AND acpt.trandd      = prtn.trandd
		                                       AND acpt.tranno      = prtn.tranno
		                                       AND acpt.prcpdd      = prtn.prcpdd
		                                       AND acpt.prcpno      = prtn.prcpno
		                                       AND prtn.prcpstatcd  NOT IN ('1', '3', '6', '7')
                     WHERE a.instcd    = :arg_instcd
                       AND a.ptno      = :arg_ptno
                       AND a.delflagcd = '0'
                       AND a.instcd    = b.instcd
                       AND a.pid       = b.pid
                       AND a.instcd    = c.instcd
                       AND a.spccd     = c.spccd
                       AND a.instcd    = ptnf.instcd
                       AND a.ptnocd    = ptnf.ptnocd
                   ) a LEFT OUTER JOIN com.zsumusrb c
                ON a.grosdrid = c.userid
               AND a.acptdd   BETWEEN c.userfromdd AND c.usertodd
                   LEFT OUTER JOIN com.zsumusrb d
                ON a.readdrid = d.userid
               AND a.readdd   BETWEEN d.userfromdd AND d.usertodd
             UNION ALL
            SELECT bfpa.instcd, bfpa.ptno, bfpa.pid, bfpa.patnm, bfpa.acptdd,
                   bfpa.diagdd AS readdd,  bfpa.rrgstno1||bfpa.rrgstno2 AS rrgstno, bfpa.sex||bfpa.age,
                   '-' AS grosdrid, bfrt.tcnts AS spcnm, '' AS grosdrnm, bfpa.diagdrnm AS readdrnm,
                   SUBSTR(bfpa.ptno,1,LENGTH(bfpa.ptno)-6)||'-'||
                   SUBSTR(bfpa.ptno,  LENGTH(bfpa.ptno)-5, 6) AS dispptno,
                   CASE WHEN SUBSTR(bfpa.ptno,1,1) = 'S' THEN '1001'
                        WHEN SUBSTR(bfpa.ptno,1,1) = 'C' THEN '1003'
                        WHEN SUBSTR(bfpa.ptno,1,1) = 'B' THEN '1012'
                        ELSE '��Ÿ'
                   END AS workflagcd,
                   '-' AS spckeepflagcd, bfpa.diagdd AS grostestdd, TO_CHAR(SYSDATE, 'HH24MISS') AS grostesttm
                 , '-' AS acpttm
                 , '-' AS acptnm
                 , '-' AS diagcnts
                 , '-' AS prcpdd
                 , '-' AS orddeptcd
                 , '-' AS testflagcd
                 , '-' AS coophospnm
                 , '-' AS statsworkflagcd
                 , 0   AS cstecnt
                 , '-' AS orddd
                 , '-' AS orddrid
                 , '-' AS fstrgstrid
                 , '-' AS prcpgenrflag
                 , '-' AS delivedd
                 , '-' AS delivetm
                 , '-' AS delivenm
                 , '-' as fstreaddrid
                 , '-' as fstreaddrnm
                 , 'N' as rsltstatcd
                 , 0   AS prcpno -- [SR20150708000187]  20150727 ��Ź�˻� �߰�������ſ��� ��Ź�Ƿڵ����� ���°� ������ ���� �߰�
              FROM lis.lprmbfpa bfpa, lis.lprmbfrt bfrt
             WHERE bfpa.instcd     = :arg_dutplceinstcd
               AND bfpa.ptno       = :arg_ptno
               AND bfpa.rgsthistno = 1
               AND bfpa.instcd     = bfrt.instcd
               AND bfpa.ptno       = bfrt.ptno
               AND bfpa.rgstdd     = bfrt.rgstdd
               AND bfpa.rgstno     = bfrt.rgstno
               AND bfrt.rgsthistno = 1


/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml gethoirtestcdlist 
param=[017, 17488137, M17003176, 017, 17488137, M17003176, 017, 17488137, M17003176]|
*/   
             SELECT acpt.instcd, acpt.pid, acpt.orddd, acpt.testnm,
                   acpt.prcpdd,  acpt.prcpgenrflag, acpt.orddrnm,   acpt.orddeptnm, acpt.deptengabbr,
                   TRIM(COALESCE(usrb.usernm,' ')) AS atdoctnm,     acpt.reqdt,
                   acpt.acpttm, acpt.prcprgsttm, rslt.readtm, TO_CHAR(rslt.lastupdtdt,'YYYY-MM-DD HH24:MI:SS')  AS rgstdd
              FROM (
                    SELECT
                           min(row_col.val) testnm,
                           MIN(a.prcpdd)  AS prcpdd,  MIN(a.prcpgenrflag) AS prcpgenrflag,
                           MIN(d.usernm)  AS orddrnm, MIN(c.depthngnm)    AS orddeptnm, MIN(c.deptengabbr)    AS deptengabbr,
                           MIN(a.instcd)  AS instcd,  MIN(a.pid)          AS pid,
                           MIN(a.orddd)   AS orddd,   MIN(a.acpttm)   AS acpttm,
                           MIN(a.prcprgsttm)   AS prcprgsttm,
                           MIN(a.ptno)   AS ptno,
                           MIN(
                               NVL(
                                   (SELECT TO_CHAR(TO_DATE(trhd.reqdd||trhd.reqtm, 'YYYYMMDDHH24MISS'), 'YYYY-MM-DD HH24:MI:SS')
                                      FROM lis.llcmtrte trte, lis.llcmtrhd trhd
                                     WHERE trte.instcd      = a.instcd
                                       AND trte.acptdd      = a.acptdd
                                       AND trte.acptno      = a.acptno
                                       AND trte.ptno        = a.ptno
                                       AND trte.acptstatcd IN ('0', '1')
                                       AND trte.instcd      = trhd.instcd
                                       AND trte.trustinstcd = trhd.trustinstcd
                                       AND trte.deptflagcd  = trhd.deptflagcd
                                       AND trte.reqdd       = trhd.reqdd
                                       AND trte.reqno       = trhd.reqno
                                   ), '-'
                               )
                           ) AS reqdt
                      FROM lis.lpjmacpt a, lis.lpcmtest b,
                           com.zsdddept c, com.zsumusrb d,
				(
					SELECT
					        substr(sys_connect_by_path(val,'��'),2) AS val
					FROM
					        (
					        SELECT 	a.instcd, a.pid, a.prcpdd, a.prcpgenrflag, a.orddeptcd,
					                a.orddrid, a.ptno, a.prcprgsttm, a.acpttm, a.orddd,
					                a.acptno, a.acptdd,
					        	TRIM(b.testengnm) val,
					        	row_number() OVER ( ORDER BY DECODE( NVL(a.trandd, '-'), '-', '9', a.trandd ), a.fstrgstrid) rn,
					        	COUNT (*) OVER () cnt
						    FROM lis.lpjmacpt a, lis.lpcmtest b
			             		/************** testnm �� �������� ���� ���� ********/
			             WHERE a.instcd         = :arg_instcd
							  AND a.pid         = :arg_pid
							  AND a.ptno        = :arg_ptno
							  AND a.acptstatcd IN ('0', '2', '3', '4', '9')
							  AND a.instcd      = b.instcd
							  AND a.testcd      = b.testcd
					        )
					WHERE 	level = cnt
						start with rn = 1
						connect by  prior rn = rn-1
				)row_col
			WHERE
			    /************** testnm �̿��� ����  �������� ���� ���� ********/
                    a.instcd      	= :arg_instcd
			    AND a.pid         	= :arg_pid
			    AND a.ptno        	= :arg_ptno
			    AND a.acptstatcd IN ('0', '2', '3', '4', '9')
                AND a.instcd      = b.instcd
                AND a.testcd      = b.testcd
                AND a.instcd      = c.instcd
                AND a.orddeptcd   = c.deptcd
                AND a.prcpdd      BETWEEN c.valifromdd AND c.valitodd
                AND a.orddrid     = d.userid
                   ) acpt LEFT OUTER JOIN pam.pmihinpt inpt ON acpt.instcd         = inpt.instcd
											               AND acpt.pid            = inpt.pid
											               AND acpt.orddd          = inpt.indd
											               AND inpt.indschacptstat = 'A'
											               AND inpt.histstat       = 'Y'
											               AND inpt.mskind         = 'M'
	                      LEFT OUTER JOIN com.zsumusrb usrb ON inpt.atdoctid = usrb.userid
	                                                       AND acpt.orddd    BETWEEN usrb.userfromdd AND usrb.usertodd
	                      LEFT OUTER JOIN lis.lprmrslt rslt ON acpt.instcd      = rslt.instcd
	                                                       AND acpt.ptno        = rslt.ptno
	                                                       AND acpt.pid         = rslt.pid
	                                                       AND rslt.rsltrgsthistno = '1'
             WHERE NVL(acpt.instcd, '-') != '-'
             UNION ALL
            SELECT bfpa.instcd, bfpa.pid, bfpa.orddd, bfrt.pcnts AS testnm, bfpa.prcpdd,
                   bfpa.prcpgenrflag, bfpa.orddrnm, bfpa.orddeptnm, bfpa.orddeptnm AS deptengabbr, bfpa.atdoctnm,
                   TO_CHAR(SYSDATE, 'YYYYMMDD') AS reqdt,
                   '000000' AS acpttm, '000000' AS prcprgsttm, '000000' AS readtm , ''  AS rgstdd
              FROM lis.lprmbfpa bfpa, lis.lprmbfrt bfrt
             WHERE bfpa.instcd     = :arg_instcd
               AND bfpa.pid        = :arg_pid
               AND bfpa.ptno       = :arg_ptno
               AND bfpa.rgsthistno = 1
               AND bfpa.instcd     = bfrt.instcd
               AND bfpa.ptno       = bfrt.ptno
               AND bfpa.rgstdd     = bfrt.rgstdd
               AND bfpa.rgstno     = bfrt.rgstno
               AND bfrt.rgsthistno = 1
               

/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getdiaginfo 
param=[017, 17488137, M17003176]
*/
            SELECT diag.termengnm, diag.termhngnm,
                   diag.pid, diag.orddd, diag.cretno, diag.prcpgenrflag, diag.prcpdd
              FROM (
                    SELECT REPLACE(
                            REPLACE(
                              REPLACE(
                                REPLACE(
                                  SUBSTR(XMLAGG(XMLELEMENT( NAME "a", '��'||TRIM(b.termengnm))).extract('//text()'),2),  
                                 -- 20120118 ����ȯ (����) ' -> ` �� ��ȯ�ȵǴ� ����
                               '&'||'lt;', '<'),
                              '&'||'gt;', '>'),
                             '&'||'amp;', '&'),
                           '&'||'apos;', '`') AS termengnm,
                           REPLACE(
                            REPLACE(
                              REPLACE(
                                REPLACE(
                                  SUBSTR(XMLAGG(XMLELEMENT( NAME "a", '��'||TRIM(b.termhngnm))).extract('//text()'),2) ,
                                  -- 20120118 ����ȯ (����) ' -> ` �� ��ȯ�ȵǴ� ����
                               '&'||'lt;', '<'),
                              '&'||'gt;', '>'),
                             '&'||'amp;', '&'),
                           '&'||'apos;', '`') AS termhngnm,
                           MIN(z.pid)    AS pid,    MIN(z.orddd)  AS orddd, 
                           MIN(z.cretno) AS cretno, MIN(z.prcpdd) AS prcpdd,
                           MIN(z.prcpgenrflag) AS prcpgenrflag
                     FROM (
                              SELECT a.pid, b.orddd, b.cretno, a.orddeptcd, 
                                     a.instcd, a.prcpgenrflag, MIN(a.prcpdd) AS prcpdd
                                FROM lis.lpjmacpt a, 
        
                                     -- �Կ� ���
                                     emr.mmodexip b, emr.mmohiprc c
                                     -- �ܷ� ���
                                     emr.mmodexop b, emr.mmohoprc c
           
                               WHERE a.instcd         = :arg_instcd
                                 AND a.pid            = :arg_pid
                                 AND a.ptno           = :arg_ptno
                                 AND a.acptstatcd    IN ('0', '2', '3', '4', '9')
                                 AND a.instcd         = b.instcd
                                 AND a.pid            = b.pid
                                 AND a.prcpdd         = b.prcpdd
                                 AND a.execprcpuniqno = b.execprcpuniqno
                                 AND b.instcd         = c.instcd
                                 AND b.prcpdd         = c.prcpdd
                                 AND b.prcpno         = c.prcpno
                                 AND b.prcphistno     = c.prcphistno
                                 AND b.execprcphistcd = 'O'
                                 AND c.prcphistcd     = 'O'
                                 AND c.prcpclscd      = 'D2'
                                 AND c.tempprcpflag   = 'N'
                               GROUP BY a.pid,    b.orddd, b.cretno, a.orddeptcd, 
                                        a.instcd, a.prcpgenrflag
                           ) z, emr.mmohdiag a, emr.mrtmterm b
                     WHERE a.pid            = z.pid
                       AND a.orddd          = z.orddd
                       AND a.cretno         = z.cretno
                       AND a.orddeptcd      = z.orddeptcd
                       AND a.instcd         = z.instcd           --   AND a.genrflagcd     = z.prcpgenrflag
                       AND a.diaghistcd     = 'O'
                       AND a.diagtypecd     = 'D'
                       AND a.diagkindcdflag = 'M'
                       AND a.instcd         = b.instcd
                       AND a.diagcd         = b.termcd
                       AND b.termflag       = '0'    
                       AND a.diagdd BETWEEN b.termfromdd AND b.termtodd
                   ) diag
                   
/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getbfptnohorilist 
param=[017, 17488137, M17003176, 017]
*/
            SELECT pnis.instcd, pnis.ptno,pnis.dispptno
              FROM (
                    SELECT a.instcd, 
                           SUBSTR(XMLAGG(XMLELEMENT( NAME "a", '��'||TRIM(a.ptno))).extract('//text()'),2) AS ptno,
                           SUBSTR(XMLAGG(XMLELEMENT( NAME "a", '��'||TRIM(a.dispptno))).extract('//text()'),2) AS dispptno
                           -- SR20150707000122 20150904 hamtn_����_�����ȸ_ū���_��ȸ_����
                           --REPLACE(WM_CONCAT(a.ptno), ',', '��') as ptno,
                           --REPLACE(WM_CONCAT(a.dispptno), ',', '��') as dispptno
                      FROM (
                            SELECT a.instcd, a.pid, a.ptno,
                                  SUBSTR(a.ptno,1,LENGTH(a.ptno)-6)||'-'||
                                  SUBSTR(a.ptno, LENGTH(a.ptno)-5, 6) AS dispptno
                              FROM lis.lpcmpnis a
                            WHERE a.instcd     = :arg_instcd
                              AND a.pid        = :arg_pid
                              AND a.delflagcd  = '0'   
                              AND a.ptno      != :arg_ptno
                              AND a.ptnocd IN (SELECT test.ptnocd 
                                                 FROM lis.lpcmtest test
                                                WHERE test.instcd = :arg_
                                                  AND test.workflagcd in ('1002', '1001', '1003', '1007')
                                                GROUP BY test.ptnocd
                                              )
                            ORDER BY a.instcd, a.acptdd DESC, a.ptno DESC
                          ) a
                    GROUP BY a.instcd
                   ) pnis


/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getrelaptnotestnmlist 
param=[017, 17488137, M17003176, 017, 17488137, M17003176, M17003176]
*/
            SELECT pnis.instcd, pnis.pid, pnis.ptno, pnis.dispptno, pnis.workflagcd,
                   SUBSTR(pnis.testnmlist, 1, LENGTH(pnis.testnmlist)) AS testnmlist
                 , pnis.statnm
                 , pnis.acptstatcd
              FROM (
                    SELECT pnis.instcd, pnis.pid, pnis.ptno, MIN(pnis.dispptno) AS dispptno,
                           MIN(pnis.workflagcd) AS workflagcd,
                           REPLACE(REPLACE(REPLACE(SUBSTR(XMLAGG(XMLELEMENT( NAME "a", '��'||TRIM(test.testengnm))).extract('//text()'),2),
                                                         '<', '<'), '>', '>'), '&', '&') AS testnmlist
                         , CASE MAX(acpt.acptstatcd) WHEN '0' THEN '����'
                                                     WHEN '2' THEN '������'
                                                     WHEN '3' THEN '��������'
                                                     ELSE ''
                           END AS statnm
                         , MAX(acpt.acptstatcd) AS acptstatcd
                      FROM (
                            SELECT a.instcd, a.pid, a.ptno, ptnf.workflagcd,
                                   SUBSTR(a.ptno,1,LENGTH(a.ptno)-6)||'-'||
                                   SUBSTR(a.ptno, LENGTH(a.ptno)-5, 6) AS dispptno          
                              FROM lis.lpcmpnis a, lis.lpcmptnf ptnf
                             WHERE a.instcd    = :arg_instcd
                               AND a.pid       = :arg_pid
                               AND (a.relaptno = :arg_ptno OR 
                                    a.ptno IN (SELECT z.relaptno FROM lis.lpcmpnis z 
                                                WHERE z.instcd    = :arg_instcd
                                                  AND z.pid       = :arg_pid
                                                  AND z.ptno      = :arg_ptno
                                                  AND z.delflagcd = '0'
                                              )
                                   )
                               AND a.delflagcd = '0'
                               AND a.instcd    = ptnf.instcd
                               AND a.ptnocd    = ptnf.ptnocd
                           ) pnis, lis.lpjmacpt acpt, lis.lpcmtest test
                     WHERE pnis.instcd      = :arg_instcd
                       AND pnis.ptno       != :arg_ptno
                       AND pnis.instcd      = acpt.instcd
                       AND pnis.ptno        = acpt.ptno
                       AND pnis.pid         = acpt.pid
                       AND acpt.acptstatcd IN ('0', '2', '3', '4', '9')
                       AND acpt.instcd      = test.instcd
                       AND acpt.testcd      = test.testcd
                     GROUP BY pnis.instcd, pnis.pid, pnis.ptno
                   ) pnis
             ORDER BY pnis.instcd, pnis.pid, pnis.ptno DESC
        

/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getaddrsltlist 
param=[017, M17003176]*/
            SELECT addr.instcd,    addr.ptno,     addr.readaddbase, addr.readrslt, 
                   addr.cmtcnts,   addr.rgstdd,   addr.addreaddrnm,
                   addr.addrsltdg, addr.diagcnts, addr.etccnts,
                   addr.cnfmdd,    addr.cnfmtm,   addr.cnfmid,
                   TRIM(COALESCE(usrb.usernm, '-')) AS cnfmnm,
                   addr.specdrid,-- ����������
                   TRIM(COALESCE(usrb2.usernm, '-')) AS specdrnm, -- ���������� ��
                   addr.rsltkindflag,
                   addr.grosdesc,
                   addr.microscopicdesc                   
              FROM (
                    SELECT addr.instcd,    addr.ptno,     addr.readaddbase, addr.readrslt, 
                           addr.cmtcnts,   addr.rgstdd,   TRIM(usrb.usernm) AS addreaddrnm,
                           addr.addrsltdg, addr.diagcnts, addr.etccnts,
                           COALESCE(addr.cnfmdd, '-') AS cnfmdd, COALESCE(addr.cnfmtm, '-') AS cnfmtm,
                           COALESCE(addr.cnfmid, '-') AS cnfmid,
                           COALESCE(addr.specdrid, '-') AS specdrid,
                           addr.rsltkindflag,
                           addr.grosdesc,
                           addr.microscopicdesc 
                      FROM lis.lprmaddr addr, com.zsumusrb usrb
                     WHERE addr.instcd    = :arg_instcd
                       AND addr.ptno      = :arg_ptno 
                       AND addr.dghistno  = 1
                       AND addr.delflagcd = '0'
                       AND addr.readdrid  = usrb.userid
                       AND addr.rgstdd    BETWEEN usrb.userfromdd AND usrb.usertodd
                   ) addr 
                  LEFT OUTER JOIN com.zsumusrb usrb
                               ON addr.cnfmid  = usrb.userid
                              AND addr.cnfmdd  BETWEEN usrb.userfromdd AND usrb.usertodd
                              LEFT OUTER JOIN com.zsumusrb usrb2
                               ON addr.specdrid  = usrb2.userid
                              AND addr.rgstdd  BETWEEN usrb2.userfromdd AND usrb2.usertodd 
             ORDER BY addr.rsltkindflag, addr.instcd, addr.addrsltdg, addr.cnfmdd, addr.cnfmtm
        
/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getcnstrsltlist 
param=[017, M17003176, J1]
*/
            SELECT cnst.instcd,           cnst.ptno,
                   cnst.cnstdg,           cnst.cnstflagcd, resn.resncnts,
                   cnst.casedcuscd,
                   TO_DATE(cnst.trsmdd||cnst.trsmtm, 'YYYYMMDD HH24MISS') AS trsmdt, cnst.trsmid, usrb.usernm AS trsmnm,
                   COALESCE(cnst.diagdd,'-') AS diagdd, cnst.diagtm,
                   cnst.cnsthospcd,       cnst.cnstdrid,
                   cnst.hosoutcnsthospnm, cnst.hosoutcnstdrid, cnst.hosoutcnstdrnm,   
                   cnst.estmdiagcnts,     cnst.cnstdiagcnts
              FROM lis.lprmcnst cnst, lis.lpcmresn resn, com.zsumusrb usrb
             WHERE cnst.instcd     = :arg_instcd
               AND cnst.ptno       = :arg_ptno 
               AND cnst.cnsthistdg = 1
               AND cnst.delflagcd  = '0'
               AND cnst.instcd     = resn.instcd
               AND resn.baseflagcd = :arg_baseflagcd
               AND cnst.cnstflagcd = resn.resncd
               AND cnst.trsmid     = usrb.userid
               AND cnst.trsmdd     BETWEEN usrb.userfromdd AND usrb.usertodd     
               AND COALESCE(cnst.cnstdiagcnts,'-') != '-'   
        
/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getopnm 
param=[I, 20170724, 20170724, 20170724, 20170724, 17488137, 20170628, 1, 017]
*/
            SELECT oprd.trgtcd AS opcd, oper.opengnm AS opnm
              FROM emr.mmrmformrec mrec INNER JOIN emr.mmodoprd oprd 
                ON mrec.formrecseq = oprd.opreclnkno
               AND mrec.chosflag   = :arg_prcpgenrflag
               AND mrec.instcd     = oprd.instcd
               AND mrec.formcd     = '0000000676'
               AND oprd.cdflag     = '4'                           
                   INNER JOIN emr.mmbvoper oper 
                ON oprd.trgtcd        = oper.orgopcd
               AND oper.opattrfromdd <= :arg_prcpdd
               AND oper.opattrtodd   >= :arg_prcpdd
               AND oper.termfromdd   <= :arg_prcpdd
               AND oper.termtodd     >= :arg_prcpdd
               AND oprd.INSTCD        = oper.instcd
             WHERE mrec.pid    = :arg_pid
               AND mrec.orddd  = :arg_orddd
               AND mrec.cretno = CAST(:arg_cretno AS INTEGER)
               AND mrec.instcd = :arg_instcd
               AND ROWNUM < 2
             ORDER by mrec.lastupdtdt DESC
 

/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getcncllist        
param=[017, M17003176]
*/
            SELECT (SELECT resn.resncnts
                      FROM lis.lpcmresn resn
                     WHERE resn.instcd = rslt.instcd
                       AND resn.baseflagcd = rslt.cnclflagcd
                       AND resn.resncd     = rslt.cnclresncd)   AS cnclresn
              FROM lis.lprmrslt rslt
             WHERE rslt.instcd = :arg_instno
               AND rslt.ptno   = :arg_ptno
               AND rslt.cncldd IS NOT NULL
               AND rslt.cncldd <> '-'
               AND (rslt.rsltrgstdd, rslt.rsltrgstno) IN (SELECT rslt2.rsltrgstdd, rslt2.rsltrgstno 
                                                            FROM lis.lprmrslt rslt2
                                                           WHERE rslt2.instcd         = rslt.instcd
                                                             AND rslt2.ptno           = rslt.ptno
                                                             AND rslt2.rsltrgsthistno = 1
                                                             AND rslt2.delflagcd      = '0'
                                                          )
            ORDER BY cncldd DESC, cncltm DESC

/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getRtnSlide 
param=[017, 17488137, M17003176]
*/
    SELECT CASE WHEN rtnflag = 'Y' THEN rtndd ||'�Ͽ� �ݳ�.' END AS rtnyn
	  FROM lis.lprmchic chic
	 WHERE instcd = :arg_instcd
	   AND pid    = :arg_pid
	   AND ptno   = :arg_ptno
		   
		   
/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getworkrelaptnolist 
param=[017, 17488137, M17003176, 017, 17488137, M17003176, 017, 17488137, M17003176, 017, 17488137, M17003176]
*/
            SELECT a.instcd, a.pid, a.ptno, a.workflagcd,
                   CASE WHEN TRIM(COALESCE(rslt.readdd, '-')) = '-' THEN '-'
                        WHEN TRIM(COALESCE(rslt.readdd, '-')) = ''  THEN '-'
                        ELSE TRIM(COALESCE(rslt.readdd, '-'))
                   END AS readdd,
                   COALESCE(rslt.preprsltflagcd, '0') AS preprsltflagcd,
                   NVL(
                       (SELECT MIN(z.acptstatcd) FROM lis.lpjmacpt z
                         WHERE z.instcd = a.instcd
                           AND z.pid    = a.pid
                           AND z.ptno   = a.ptno
                           AND z.acptstatcd IN ('0', '2', '3', '4')
                       ), '0'
                   ) AS acptstatcd
                   , readtm
              FROM (
                    SELECT a.instcd, a.pid, a.ptno, 
                           CASE WHEN (SELECT COUNT(z.testcd) 
                                        FROM lis.lpjmacpt z, lis.lpcmtest x
                                       WHERE z.instcd      = a.instcd
                                         AND z.pid         = a.pid
                                         AND z.ptno        = a.ptno
                                         AND z.acptstatcd IN ('0', '2', '3', '4', '9')
                                         AND z.instcd      = x.instcd
                                         AND z.testcd      = x.testcd
                                         AND x.workflagcd  = '1002'
                                     ) > 0 THEN '1002' 
                                ELSE b.workflagcd
                           END AS workflagcd
                      FROM lis.lpcmpnis a, lis.lpcmptnf b
                     WHERE a.instcd = ?
                       AND a.pid    = ?
                       AND (a.ptno  = ?
                         OR a.ptno IN (SELECT z.ptno FROM lis.lpcmpnis z 
                                       WHERE z.instcd      = :arg_instcd
                                         AND z.pid         = :arg_pid
                                         AND z.ptno        = :arg_ptno
                                          AND z.delflagcd = '0'
                                      )
                         OR a.ptno IN (SELECT z.relaptno FROM lis.lpcmpnis z 
                                       WHERE z.instcd      = :arg_instcd
                                         AND z.pid         = :arg_pid
                                         AND z.ptno        = :arg_ptno
                                          AND z.delflagcd = '0'
                                      )
                           )
                       AND a.delflagcd = '0'
                       AND a.instcd    = b.instcd
                       AND a.ptnocd    = b.ptnocd 
                   ) a LEFT OUTER JOIN lis.lprmrslt rslt
                ON a.instcd = rslt.instcd
               AND a.ptno   = rslt.ptno
               AND a.pid    = rslt.pid
               AND rslt.rsltrgsthistno = 1
             UNION ALL
            SELECT bfpa.instcd, bfpa.pid, bfpa.ptno,
                   CASE WHEN SUBSTR(bfpa.ptno,1,1) = 'S' THEN '1001'
                        WHEN SUBSTR(bfpa.ptno,1,1) = 'C' THEN '1003'
                        WHEN SUBSTR(bfpa.ptno,1,1) = 'B' THEN '1012'
                        ELSE '��Ÿ'
                   END AS workflagcd,
                   bfpa.diagdd AS readdd, '0' AS preprsltflagcd,
                   '730' AS acptstatcd
                   , '' readtm 
              FROM lis.lprmbfpa bfpa
             WHERE bfpa.instcd      = :arg_instcd
               AND bfpa.pid         = :arg_pid
               AND bfpa.ptno        = :arg_ptno
               AND bfpa.rgsthistno = 1
             ORDER BY 1, 2, readdd, readtm, 3 DESC 
        

/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getmoletestlist 
param=[017, M17003176]
*/
            SELECT acpt.instcd, acpt.ptno, acpt.testcd, acpt.pid, 
                   SUBSTR(acpt.ptno,1,LENGTH(acpt.ptno)-6)||'-'||
                   SUBSTR(acpt.ptno,  LENGTH(acpt.ptno)-5, 6) AS dispptno,
                   acpt.testnm, acpt.rsltcnts1, acpt.rsltcnts3, acpt.rsltcnts4, acpt.rsltcnts5, acpt.cmtcnts, 
                   COALESCE(trlt.testrslt, ' ') AS testrslt
                 , NVL(acpt.readdd, '-') AS readdd
                 , NVL(acpt.readtm, '-') AS readtm
                 , acpt.readid
                 , acpt.keybloc
                 , acpt.workflagcd -- [SR20150721000117] 20150824  �������� �˻� �߰��� ���ں����� �����ϱ� ���� �÷��׷� WorkFlagCd �߰�
              FROM (
                    SELECT acpt.instcd, acpt.ptno, acpt.testcd, acpt.pid, 
                           COALESCE(test.testengabbr, test.testengnm) AS testnm,
                           cnts.rsltcnts1, cnts.rsltcnts3, cnts.rsltcnts4, cnts.rsltcnts5,cnts.cmtcnts
                         , rslt.readdd
                         , rslt.readtm
                         , rslt.readid
                         , rslt.keybloc
                         , ptnf.workflagcd
                      FROM lis.lpjmacpt acpt, lis.lprmrslt rslt, lis.lpcmtest test, 
                           lis.lprmcnts cnts
                           , lis.lpcmptnf ptnf
                     WHERE acpt.instcd         = :arg_instcd
                       AND acpt.ptno           = :arg_ptno
                       AND acpt.acptstatcd    IN ('0', '2', '3', '4', '9')
                       AND acpt.instcd         = rslt.instcd
                       AND acpt.ptno           = rslt.ptno
                       AND acpt.pid            = rslt.pid
                       AND rslt.rsltrgsthistno = 1
                       AND rslt.delflagcd      = '0'
                       AND acpt.instcd         = cnts.instcd
                       AND acpt.ptno           = cnts.ptno
                       AND acpt.pid            = cnts.pid
                       AND cnts.rsltrgsthistno = 1
                       AND cnts.delflagcd      = '0'
                       AND acpt.instcd         = test.instcd
                       AND acpt.testcd         = test.testcd
                       AND acpt.instcd		   = ptnf.instcd
                       AND acpt.ptnocd		   = ptnf.ptnocd
                   ) acpt LEFT OUTER JOIN lis.lprmtrlt trlt
                  ON acpt.instcd = trlt.instcd
                 AND acpt.ptno   = trlt.ptno
                 AND acpt.testcd = trlt.testcd
                 AND trlt.rsltrgsthistno = 1
                 AND trlt.delflagcd      = '0'


/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getaddrsltlist 
param=[017, M17003176]
*/
            SELECT addr.instcd,    addr.ptno,     addr.readaddbase, addr.readrslt, 
                   addr.cmtcnts,   addr.rgstdd,   addr.addreaddrnm,
                   addr.addrsltdg, addr.diagcnts, addr.etccnts,
                   addr.cnfmdd,    addr.cnfmtm,   addr.cnfmid,
                   TRIM(COALESCE(usrb.usernm, '-')) AS cnfmnm,
                   addr.specdrid,-- ����������
                   TRIM(COALESCE(usrb2.usernm, '-')) AS specdrnm, -- ���������� ��
                   addr.rsltkindflag,
                   addr.grosdesc,
                   addr.microscopicdesc                   
              FROM (
                    SELECT addr.instcd,    addr.ptno,     addr.readaddbase, addr.readrslt, 
                           addr.cmtcnts,   addr.rgstdd,   TRIM(usrb.usernm) AS addreaddrnm,
                           addr.addrsltdg, addr.diagcnts, addr.etccnts,
                           COALESCE(addr.cnfmdd, '-') AS cnfmdd, COALESCE(addr.cnfmtm, '-') AS cnfmtm,
                           COALESCE(addr.cnfmid, '-') AS cnfmid,
                           COALESCE(addr.specdrid, '-') AS specdrid,
                           addr.rsltkindflag,
                           addr.grosdesc,
                           addr.microscopicdesc 
                      FROM lis.lprmaddr addr, com.zsumusrb usrb
                     WHERE addr.instcd    = :arg_instcd
                       AND addr.ptno      = :arg_ptno  
                       AND addr.dghistno  = 1
                       AND addr.delflagcd = '0'
                       AND addr.readdrid  = usrb.userid
                       AND addr.rgstdd    BETWEEN usrb.userfromdd AND usrb.usertodd
                   ) addr 
                  LEFT OUTER JOIN com.zsumusrb usrb
                               ON addr.cnfmid  = usrb.userid
                              AND addr.cnfmdd  BETWEEN usrb.userfromdd AND usrb.usertodd
                              LEFT OUTER JOIN com.zsumusrb usrb2
                               ON addr.specdrid  = usrb2.userid
                              AND addr.rgstdd  BETWEEN usrb2.userfromdd AND usrb2.usertodd 
             ORDER BY addr.rsltkindflag, addr.instcd, addr.addrsltdg, addr.cnfmdd, addr.cnfmtm
        

/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getreaddoctlist 
param=[017, M17003176, 1, 1]
*/
            SELECT rddt.ptno, rddt.userid, NVL(rddt.usernm,'-') AS usernm,
                   COALESCE(usab.userabbr, rddt.usernm) AS userabbr
              FROM (
                    SELECT rddt.instcd, rddt.ptno, rddt.userid, TRIM(usrb.usernm) AS usernm
                      FROM lis.lprmrddt rddt, com.zsumusrb usrb
                     WHERE rddt.instcd     = :arg_instcd
                       AND rddt.ptno       = :arg_ptno
                       AND rddt.userflagcd = :arg_userflagcd
                       AND rddt.rgstdg     = CAST(:arg_rgstdg AS INTEGER)
                       AND rddt.delflagcd  = '0'
                       AND rddt.dispseqno  = 1
                       AND rddt.userid     = usrb.userid
                       AND rddt.userrgstdd BETWEEN usrb.userfromdd AND usrb.usertodd
                   ) rddt LEFT OUTER JOIN lis.lpcmusab usab
                ON rddt.instcd = usab.instcd
               AND rddt.userid = usab.userid
        
/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml gethospenvsetinfo 
param=[017]
*/	    
		    SELECT instcd,     lendrtnterm, recvqualmthdcd,
                   plgydeptcd, plgyteamcd,  doctjobgradcd, teamjobgradcd,
                   COALESCE(csteeqmtip, '-')      AS csteeqmtip,
                   COALESCE(csteeqmtport, '-')    AS csteeqmtport,
                   COALESCE(slideqmtip, '-')      AS slideqmtip,
                   COALESCE(slideqmtport, '-')    AS slideqmtport,
                   COALESCE(slidbceqmtip,    '-') AS slidbceqmtip,
                   COALESCE(slidbceqmtport,  '-') AS slidbceqmtport,
                   COALESCE(slidbceqmtip2,   '-') AS slidbceqmtip2,
                   COALESCE(slidbceqmtport2, '-') AS slidbceqmtport2,
                   COALESCE(repttitl,  '-')       AS repttitl,
                   COALESCE(repttitl2, '-')       AS repttitl2,
                   COALESCE(plgyrecroom1, '-')    AS plgyrecroom1,
                   COALESCE(plgyrecroom2, '-')    AS plgyrecroom2,
                   COALESCE(ptnoacptflag, '0')    AS ptnoacptflag
			  FROM lis.lpcmhpes
			 WHERE instcd = :arg_instcd
		

/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getspecdrnmlist 
param=[017, M17003176, 2, 1]
*/
	    SELECT CASE MAX(a.rgstdd) WHEN '-' THEN 'false' ELSE 'true' END AS choi,
                   a.instcd, a.userid, 
                   MIN(a.usernm)     AS usernm,
                   MAX(a.ptno)       AS ptno,
                   MAX(a.rgstdd)     AS rgstdd, 
                   MAX(a.rgstno)     AS rgstno,
                   MAX(a.userrgstdd) AS userrgstdd,
                   MAX(a.dispseqno)  AS dispseqno,
                   MAX(a.dispseqno)  AS orgdispseqno
              FROM (
                   SELECT a.instcd, a.ptno, a.userid,  
                          TRIM(usrb.usernm) AS usernm, 
                          a.userrgstdd, a.dispseqno, a.rgstdd, a.rgstno, 99999 AS sepodispseq
                     FROM lis.lprmrddt a, com.zsumusrb usrb
                    WHERE a.instcd            = ?
                      AND a.ptno              = ?
                      AND a.userflagcd        = ?
                      AND a.rgstdg            = TO_NUMBER(?)
                      AND a.delflagcd         = '0'
                      AND a.userid            = usrb.userid
                      AND a.userrgstdd        BETWEEN usrb.userfromdd AND usrb.usertodd
                   ) a
             GROUP BY a.instcd, a.userid
        

/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getreqfrmcnts     
param=[017, M17003176, 17488137, 017, M17003176, 17488137]
*/
            SELECT '1' AS rno, reqf.formrecseq,
                   NVL(spcinfo.reqcnt, ' ') AS reqcnt,
                   NVL(spcinfo.endono, '0') AS endono
              FROM (
                    SELECT acpt.instcd, acpt.pid, 
                          CAST(acpt.reqfrmno AS INTEGER) AS reqno
                      FROM lis.lpjmacpt acpt 
                    WHERE acpt.instcd      = :arg_instcd
                      AND acpt.ptno        = :arg_ptno
                      AND acpt.pid         = :arg_pid
                      AND acpt.acptstatcd IN ('0', '2', '3', '4')
                    GROUP BY acpt.instcd, acpt.pid, acpt.reqfrmno
                   ) acpt, lis.llchreqf reqf,
                  XMLTABLE(
                      '/reqfrminfo'
                      PASSING reqf.reqcnts
                      COLUMNS reqcnt VARCHAR2(1000)   PATH '/reqfrminfo/reqcnts',
                              endono VARCHAR2(1000)   PATH '/reqfrminfo/endono'
                  ) spcinfo
             WHERE acpt.instcd = reqf.instcd
               AND acpt.pid    = reqf.pid
               AND acpt.reqno  = reqf.reqno
               AND reqf.reqhistno = 1       
            UNION ALL 
            SELECT '2' AS rno, reqf.formrecseq,
                   spcinfo.resnnm AS reqcnt,    '0' AS endono
              FROM (
                    SELECT acpt.instcd, acpt.pid, 
                          CAST(acpt.reqfrmno AS INTEGER) AS reqno
                      FROM lis.lpjmacpt acpt 
                    WHERE acpt.instcd      = :arg_instcd
                      AND acpt.ptno        = :arg_ptno
                      AND acpt.pid         = :arg_pid
                    GROUP BY acpt.instcd, acpt.pid, acpt.reqfrmno
                   ) acpt, lis.llchreqf reqf,
                  XMLTABLE(
                      '/reqfrminfo/resncd/resncdlist'
                      PASSING reqf.reqcnts
                      COLUMNS choi   VARCHAR2(5)   PATH 'choi',
                              resnnm VARCHAR2(100) PATH 'resnnm'
                  ) spcinfo
             WHERE acpt.instcd = reqf.instcd
               AND acpt.pid    = reqf.pid
               AND acpt.reqno  = reqf.reqno
               AND reqf.reqhistno = 1       
               AND spcinfo.choi   = 'true'    
        
/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getrsltimglist 
param=[017, M, 17, 003176]*/
            SELECT rimg.instcd, rimg.pid, rimg.ptno, rimg.rgstdd, rimg.rgsttm,
                   rimg.rgstrid, rimg.rsltimg,       TRIM(ptbs.hngnm) AS patnm,
                   com.fn_zz_getsex(ptbs.rrgstno1, ptbs.rrgstno2, '2')||'/'||
                   com.fn_zz_getage(ptbs.rrgstno1, ptbs.rrgstno2, rimg.rgstdd, 'A', '-') AS sexage,
                   SUBSTR(rimg.ptno,1,LENGTH(rimg.ptno)-6)||'-'||
                   SUBSTR(rimg.ptno,  LENGTH(rimg.ptno)-5, 6) AS dispptno,                   
                   TRIM(usrb.usernm) AS rgstrnm
              FROM lis.lprmrimg rimg, pam.pmcmptbs ptbs, com.zsumusrb usrb
             WHERE rimg.instcd    = :arg_instcd
               AND rimg.ptno      = :arg_ptnocode||:arg_fromyear||:arg_fromptno
               AND rimg.imghistno = 1
               AND rimg.instcd    = ptbs.instcd
               AND rimg.pid       = ptbs.pid
               AND rimg.rgstrid   = usrb.userid
               AND rimg.rgstdd    BETWEEN usrb.userfromdd AND usrb.usertodd    
             ORDER BY rimg.instcd, rimg.ptno DESC
        
     
/* himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml setpathpacs 
param=[           *���ΰ�ü: ������ü other
           *�ӻ����� �� �䱸����: R/O PMO12040

            [MOLECULAR PATHOLOGY]   M17-003176       HPV genotyping real-time PCR
                                                         
              Sample adequacy
                 Adequate
              Result
                 [Methods]
                    Seegene HPV Real-time PCR (Anyplex II HPV 28 Detection Real-time PCR)
                 
                 [Result]
                 HPV High Risk Type : POSITIVE (18+, 68+, 31+++)
                 HPV Low  Risk Type : POSITIVE (70+, 61+) 
              Comments
                 3. Comment
                 �������� ���̷��� (Human papilloma virus)�� �ڱð�ξ��� �ֿ� �������ڷ� �˷��� �ֽ��ϴ�. �ϰ��� ���ü� 
                 ������ ���� �����豺 (high risk)�� �����豺 (low risk)�� ���еǸ�, �����豺 HPV�� �밳 �ð��� ������ ��
                 �ǵǰų� �縶�� ���� �缺������ ������ �Ǵ� �ݸ�, �����豺�� �ڱð�ξ��� ���߽�Ű�µ� �����մϴ� (N Eng
                 l J Med. 2003 348:518).
                 
                 �� �� ��ǰ�� 19���� �����豺 HPV (16, 18, 26, 31, 33, 25, 29, 45, 51, 52, 53, 56, 58, 59, 66, 68, 73, 82
                 ��)�� 9���� �����豺 HPV (6, 11, 40, 42, 43, 44, 54, 61, 70��), ���δ������� Ÿ���ٻ��� �����մϴ�. 
                 �� Viral load���� +++:10^5 copies/reaction, ++:10^5~10^2 copies/reaction, +:10^2 copies/reaction�� ��
                 �� �ؼ��� �� �ֽ��ϴ�. �� �� ��+���� �ſ� ���� �󵵷� ���� �ñ�, ��ü ä�� ���¿� ���� �ݺ� �˻� �� ����
                 ���� ���� �� �ֽ��ϴ�.
                 �� PCR �˻�� ��ü �� �ռ��� ���ų� �������� ��ü �Ǽ� �Ǵ� ���� ���������� �����ϴ� ��� �������� ���� 
                 �� �ֽ��ϴ�. ����, PCR �˻�� ������ ������ �˻��ϹǷ� �����հ� ����� ������ �ȵǾ� ���缺�� ���ɼ��� 
                 �ֽ��ϴ�. ��� �ؼ� �� �ӻ� ���� �������� �Ǵ��Ͻñ� �ٶ��ϴ�.
                 �� ��� �˻�� �˻� ���, �þ��� �������� �� �˻� ����� ������ �����ǿ� ���� Ȯ�εǾ����ϴ�.
                    (�˻� �����: �����)
, [C@5dada1b6
, 10602673
, M17003176
, 017 ]
*/  
            UPDATE lis.lpjmpacs
               SET reptdt     = NULL
                 , reptdrnm   = ''
                 , teststat = 'Y'
                 , maindiagnm = :arg_maindiagnm
                 , grosrslt1  = :arg_grosrslt1
                 , cnfmstat   = 'N'
                 , pacssyncdt = SYSDATE
                 , updtid     = :arg_userid
                 , updtdt     = SYSDATE
                 , updtip     = '-'
            WHERE ptno   = :arg_ptno
              AND instcd = :arg_instcd
        
     

/* -- ��������� ó�� ���� ��*/
/* -- ��������� ó�� ���� ��*/
/* -- ��������� ó�� ���� ��*/
/* -- ��������� ó�� ���� ��*/
/* -- ��������� ó�� ���� ��*/
/* -- ��������� ó�� ���� ��*/
/* -- ��������� ó�� ���� ��*/     