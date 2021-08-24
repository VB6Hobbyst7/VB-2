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
��Ʈ�� binding ����ó���� �κ� �����ּ���. nU�� ����Ŭ�� Full scan �ɰ�� DBA���� ���ܵ˴ϴ�. 
��Ʈ�� binding ����ó���� �κ� �����ּ���. nU�� ����Ŭ�� Full scan �ɰ�� DBA���� ���ܵ˴ϴ�.
��Ʈ�� binding ����ó���� �κ� �����ּ���. nU�� ����Ŭ�� Full scan �ɰ�� DBA���� ���ܵ˴ϴ�.
��Ʈ�� binding ����ó���� �κ� �����ּ���. nU�� ����Ŭ�� Full scan �ɰ�� DBA���� ���ܵ˴ϴ�.
��Ʈ�� binding ����ó���� �κ� �����ּ���. nU�� ����Ŭ�� Full scan �ɰ�� DBA���� ���ܵ˴ϴ�.
��Ʈ�� binding ����ó���� �κ� �����ּ���. nU�� ����Ŭ�� Full scan �ɰ�� DBA���� ���ܵ˴ϴ�.
��Ʈ�� binding ����ó���� �κ� �����ּ���. nU�� ����Ŭ�� Full scan �ɰ�� DBA���� ���ܵ˴ϴ�.
��Ʈ�� binding ����ó���� �κ� �����ּ���. nU�� ����Ŭ�� Full scan �ɰ�� DBA���� ���ܵ˴ϴ�.
==========*/ 


/* -- �������� ������ ������ȸ ������ ���ں��� ó�� PMO12040 HPV genotyping real-time PCR �� ���ѵ� Ưȭó�� */
/* -- �������� ������ ������ȸ ������ ���ں��� ó�� PMO12040 HPV genotyping real-time PCR �� ���ѵ� Ưȭó�� */
/* -- �������� ������ ������ȸ ������ ���ں��� ó�� PMO12040 HPV genotyping real-time PCR �� ���ѵ� Ưȭó�� */
/* -- �������� ������ ������ȸ ������ ���ں��� ó�� PMO12040 HPV genotyping real-time PCR �� ���ѵ� Ưȭó�� */
/* -- �������� ������ ������ȸ ������ ���ں��� ó�� PMO12040 HPV genotyping real-time PCR �� ���ѵ� Ưȭó�� */
/* param=[017, to_char(sysdate-1,'yyyymmdd'), to_char(sysdate,'yyyymmdd')]
������ �������� ���������� ������ ������ ��ȸ�˴ϴ�. ��ȸȭ�� ó�� �Ǵ� �Ʒ����� ó���� ��� �ʿ��� �����Դϴ�.
*/ 
select /*+ leading(acpt) */
    acpt.instcd,        --�����ȣ
    acpt.acptdd,        --��������
    acpt.acptno,        --������ȣ
    acpt.acptitemno,    --�����׸��ȣ
    acpt.ptno,          --������ȣ
    acpt.acptstatcd	    --���������ڵ�(0:����(440),1:���(100,240),2:������(710),3:�������(730), 4:�������(740))
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
-- ��� �������� ����üũ sql  ���� ����
-- and acpt.instcd         = :arg_instcd
-- and acpt.ptno           = :arg_ptno
-- and acpt.pid            = :arg_pid
-- and acpt.acptstatcd    in ('0','2')
and acpt.prcpdd between dept.valifromdd and dept.valitodd



/* acpt.acptstatcd -���������ڵ�(0:����(440),1:���(100,240),2:������(710),3:�������(730), 4:�������(740)) */
/* acpt.acptstatcd ����� �Ǽ��� ��� �������̽��� �Ǵ���... (0:����(440) ������ �űԷ� ����Է� , 2:������(710) ��� update �̻��¿����� �۵��� �Ǿ��մϴ�.*/
/* -- ��������� ó�� ���� ������ ���ں��� ó�� PMO12040 HPV genotyping real-time PCR �� ���ѵ� Ưȭó�� */
/* -- ��������� ó�� ���� ������ ���ں��� ó�� PMO12040 HPV genotyping real-time PCR �� ���ѵ� Ưȭó�� */
/* -- ��������� ó�� ���� ������ ���ں��� ó�� PMO12040 HPV genotyping real-time PCR �� ���ѵ� Ưȭó�� */
/* -- ��������� ó�� ���� ������ ���ں��� ó�� PMO12040 HPV genotyping real-time PCR �� ���ѵ� Ưȭó�� */
/* -- ��������� ó�� ���� ������ ���ں��� ó�� PMO12040 HPV genotyping real-time PCR �� ���ѵ� Ưȭó�� */
/* -- ��������� ó�� ���� ������ ���ں��� ó�� PMO12040 HPV genotyping real-time PCR �� ���ѵ� Ưȭó�� */
/* -- ��������� ó�� ���� ������ */

            SELECT TO_CHAR(SYSDATE,'YYYYMMDD') AS sysdd,
                   TO_CHAR(SYSDATE,'HH24MISS') AS systm
              FROM DUAL
              
    			prcpVO.set(iRowNo, "tretdd",         �������� yyyymmdd ���� 20170811 );
    			prcpVO.set(iRowNo, "trettm",         ����ú� hhs4miss ���� 090000 );
    			prcpVO.set(iRowNo, "readid",         �α��λ��);

/*  ������� �Ϸù�ȣ ä���� ���� ���� Row�� Lockó���Ѵ�. 
himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml setseqnolock 
arg_seqflagcd = '4'  �����߻������ڵ�(�˻���) �����Դϴ�.
param=[017, 4] */
update lis.lpcmseqn
       set lastgenrno = 1
 where instcd        = :arg_instcd   
   and seqgenryy = '1900'
   and seqflagcd  = :arg_seqflagcd
        
/* ������� ä���� �մϴ�. ���⼭ ��ȸ�� lastgenrno�� ���ʿ� rsltrgstno �� ó���˴ϴ�.
himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getlastseqno 
param=[017, 2017, 4]*/
select coalesce(lastgenrno+1, 1) as lastgenrno
  from lis.lpcmseqn
 where instcd     = :arg_instcd
   and seqgenryy  = :arg_seqgenryy 
   and seqflagcd  = :arg_seqflagcd

/* ������ ä���� �ߴµ� null �ϰ�� insert�ϰ� lastgenrno = 1�� ���� 1���� �����մϴ�. �⵵���� ���� �Է��� �˴ϴ�. */
/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml setseqnogenr 
param=[2017, 4, 017, 4, �α��λ��, �α��λ��]*/
            INSERT INTO lis.lpcmseqn (seqgenryy, seqflagcd, instcd, lastgenrno, 
                                                       fstrgstdt,      fstrgstrid,
                                                       lastupdtdt,   lastupdtrid)
                                         VALUES (:arg_seqgenryy, :arg_seqflagcd, :arg_instcd, 1, SYSDATE, :arg_userid,  SYSDATE, :arg_userid)
                                  
/* ������ ä���� �ߴµ� null �� �ƴ� ��� */                     
/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml setseqnogenr 
param=[32715, 10602673, 017, 2017, 4]*/
update lis.lpcmseqn
   set lastgenrno = :arg_lastgenrno ,
          lastupdtdt   = sysdate,
          lastupdtrid  = :arg_userid 
 where instcd        = :arg_instcd 
   and seqgenryy = :arg_seqgenryy
   and seqflagcd  = :arg_seqflagcd 
        
        
/*
		    			prcpVO.set(iRowNo, "newprcpstatcd",  "710");	// ������
		    			prcpVO.set(iRowNo, "newacptstatcd",  "2");	    // ������
		    			prcpVO.set(iRowNo, "bizflagcd",      "710");	// 
		    			prcpVO.set(iRowNo, "biztretflagcd",  "I");		//
		    			prcpVO.set(iRowNo, "truststatcd",  "4");		// CMC ����Ź����
*/

/* �˻���(Header) ���  �����մϴ�.
himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml instestrslt 
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
        
/* �˻���(Header) ���  �����մϴ� �ű������� �ƴҰ�� �̷��� ����ϴ�.*/
            INSERT INTO lis.lprmrslt (ptno,       rsltrgstdd,    rsltrgstno,     instcd,         rsltrgsthistno, 
                                      pid,        rsltrgsttm,    grostestrecdd,  grostestrectm,  grostestrecid,
                                      readdd,     readtm,        readid,         extrpartcnts,   extrmthdcnts,   
                                      diagcnts,   diagcd,
                                      spckeepflagcd, rslthideflagcd, conccaseflagcd, preprsltflagcd, ugcyalertflagcd, cnstcd,
                                      rsltrgstid, cnclflagcd,    cnclresncd,     cncldd,         cncltm,
                                      grospic,    keybloc,       tissbloct,      tissblocnt,     readgrade,
                                      cnclid,     delflagcd,     signno,
                                      fstrgstdt,  fstrgstrid,
                                      lastupdtdt, lastupdtrid, cncrjudgflagcd
                                     )
            SELECT ptno,       rsltrgstdd,    rsltrgstno,   instcd,  
                   (SELECT MAX(z.rsltrgsthistno)+1
                      FROM lis.lprmrslt z
                     WHERE instcd         = #instcd#
                       AND ptno           = #ptno#
                       AND pid            = #pid#
                   ),
                   pid,        rsltrgsttm,    grostestrecdd,  grostestrectm,  grostestrecid,
                   readdd,     readtm,        readid,         extrpartcnts,   extrmthdcnts,   
                   diagcnts,   diagcd,
                   spckeepflagcd, rslthideflagcd, conccaseflagcd, preprsltflagcd, ugcyalertflagcd,  cnstcd,
                   rsltrgstid, cnclflagcd,    cnclresncd,     cncldd,         cncltm,
                   grospic,    keybloc,       tissbloct,      tissblocnt,     readgrade,
                   cnclid,     '1',    signno,
                   fstrgstdt,  fstrgstrid,
                   lastupdtdt, lastupdtrid, cncrjudgflagcd
              FROM lis.lprmrslt
             WHERE instcd         = #instcd#
               AND ptno           = #ptno#
               AND pid            = #pid#
               AND rsltrgsthistno = 1
               AND delflagcd      = '0'        
               
/* �˻������� ��� �մϴ�. 
himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml instestrsltcnts 
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

/* �˻������� ��� �մϴ�.  �̷��� ����ϴ�.
*/       
            INSERT INTO lis.lprmcnts (ptno,       rsltrgstdd, rsltrgstno, rsltrgsthistno, instcd, pid,  
                                      rsltcnts1,  rsltcnts2,  rsltcnts3,    
                                      rsltcnts4,  rsltcnts5,  rsltcnts6,
                                      cmtcnts,    delflagcd,
                                      fstrgstdt,              fstrgstrid,
                                      lastupdtdt,             lastupdtrid)
            SELECT ptno,       rsltrgstdd, rsltrgstno, 
                   (SELECT MAX(z.rsltrgsthistno)+1
                      FROM lis.lprmcnts z
                     WHERE instcd         = #instcd#
                       AND ptno           = #ptno#
                       AND pid            = #pid#
                   ),
                   instcd,     pid,  
                   rsltcnts1,  rsltcnts2,  rsltcnts3,    
                   rsltcnts4,  rsltcnts5,  rsltcnts6,
                   cmtcnts,    '1',
                   fstrgstdt,              fstrgstrid,
                   lastupdtdt,             lastupdtrid            
              FROM lis.lprmcnts
             WHERE instcd         = #instcd#
               AND ptno           = #ptno#
               AND pid            = #pid#
               AND rsltrgsthistno = 1
               AND delflagcd      = '0'   

/* T/M/P ���� ����
  himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml updlastdiag 
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
        
/* ������ȣ �����̷� ���� 
himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml updlpcmpnis  
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

/* ������ �������� ����
himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml updexersltcomfirm 
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
          SELECT acpt.instcd, 
                 acpt.prcpdd, 
                 acpt.pid, 
                 acpt.prcpno, 
                 acpt.execprcpuniqno,
                 MIN(acpt.prcpgenrflag) AS prcpgenrflag,        -- �Կ��ܷ����� ��  
                 'I' AS biztretflagcd,  
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
                 :arg_scrno AS scrno, 
                 acpt.prcpgenrflag AS prcpgenrflagcd
              FROM lis.lpjmacpt acpt
                 , lis.lpcmpnis pnis
             WHERE acpt.instcd      = :arg_instcd
               AND acpt.ptno        = :arg_ptno
               AND acpt.pid         = :arg_pid
               AND acpt.acptstatcd IN ('0', '2', '3', '4', '9')
               AND acpt.instcd      = pnis.instcd
               AND acpt.ptno        = pnis.ptno
               AND pnis.delflagcd   = '0'
             GROUP BY acpt.instcd, acpt.prcpdd, acpt.pid, acpt.prcpno, acpt.execprcpuniqno, acpt.prcpgenrflag
        
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
--�Կ��ϰ�� prcpgenrflag = I, D, E
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

--�ܷ��ϰ�� prcpgenrflag = O, S 
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
--�Կ��ϰ�� prcpgenrflag = I, D, E
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

--�ܷ��ϰ�� prcpgenrflag = O, S 
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
        
/*  getDoPrcpTret  COUNT(prcpdd) < 0 �ϰ�쿡��   tretflagcd = 710  ó����� ������ (����Ȼ���)
himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml insDoPrcpTret 
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


/* -- ��������� ó�� ���� ��*/
/* -- ��������� ó�� ���� ��*/
/* -- ��������� ó�� ���� ��*/
/* -- ��������� ó�� ���� ��*/
/* -- ��������� ó�� ���� ��*/
/* -- ��������� ó�� ���� ��*/
/* -- ��������� ó�� ���� ��*/     


param=[M17003176, 20170724, 32715, 9, 017, 20170724, 33978, PMO12040, 1, [Methods]
   Seegene HPV Real-time PCR (Anyplex II HPV 28 Detection Real-time PCR)

[Result]
HPV High Risk Type : POSITIVE (18+, 68+, 31+++)
HPV Low  Risk Type : POSITIVE (70+, 61+) , null, null, 10602673, 10602673]|1 records| 
        
            INSERT INTO lis.lprmtrlt 
                   (ptno,          rsltrgstdd,  rsltrgstno,           rsltrgsthistno, 
                    riskflagcd,    instcd,
                    acptdd,        acptno,                            testcd,     
                    acptitemno,                      testrslt,        testrsltxml,  testrsltetc, delflagcd,
                    fstrgstdt,     fstrgstrid,
                    lastupdtdt,    lastupdtrid)
            VALUES (?,        ?, CAST(? AS DECIMAL(12,0)), 1, 
                    ?,  ?,
                    ?,      CAST(? AS DECIMAL(12,0)),   ?,     
                    CAST(? AS SMALLINT),  ?,     ?  , ?  ,   '0',
                    SYSDATE, ?,
                    SYSDATE, ?)
        
     /* himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml instestitemrslt */ |25 msec
52001286[node= user=10602673 ip=10.110.114.57][2017.07.24 14:26:13] [ INFO] [himed.his.lis.plgyrsltmngtmgr.testrsltrgstmgt.dao.TestRsltRgstDAOImpl] exeTestItemRslt() ends.(25 msecs}
52001286[node= user=10602673 ip=10.110.114.57][2017.07.24 14:26:13] [ INFO] [himed.his.lis.lib.plgycommonmgt.PlgyCommonMgtImpl] exeAcptStat() starts.
52001286[node= user=10602673 ip=10.110.114.57][2017.07.24 14:26:13] [ INFO] [himed.his.lis.lib.plgycommonmgt.dao.PlgyCommonDAOImpl] exeAcptStat() starts.
52001286[node= user=10602673 ip=10.110.114.57][2017.07.24 14:26:13] [ INFO] [himed.his.lis.lib.plgycommonmgt.dao.PlgyCommonDAOImpl] query time = 3 msec.
52001286[node= user=10602673 ip=10.110.114.57][2017.07.24 14:26:13] [ INFO] [himed.his.lis.lib.plgycommonmgt.dao.PlgyCommonDAOImpl] execute query |3 msec|

param=[017, M17003176, 20170724, 33978, 1, PMO12040, 17488137, 20170724, 1151787391]|1 records|
        
            SELECT acptstatcd 
              FROM lis.lpjmacpt 
             WHERE instcd         = ?
               AND ptno           = ?
               AND acptdd         = ?
               AND acptno         = ?
               AND acptitemno     = ?
               AND testcd         = ?
               AND pid            = ? 
               AND prcpdd         = ?
               AND execprcpuniqno = ?        
        
     /* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getacptstatref */|3 msec
52001286[node= user=10602673 ip=10.110.114.57][2017.07.24 14:26:13] [ INFO] [himed.his.lis.lib.plgycommonmgt.dao.PlgyCommonDAOImpl] execute update |3 msec|


param=[2, 10602673, 017, M17003176, 20170724, 33978, 1, PMO12040, 17488137, 20170724, 1151787391]|1 records| 
        
            UPDATE lis.lpjmacpt
               SET acptstatcd  = ?,
        
        
        
                   lastupdtdt  = SYSDATE,
                   lastupdtrid = ?
             WHERE instcd         = ?
               AND ptno           = ?
               AND acptdd         = ?
               AND acptno         = CAST(? AS DECIMAL(12,0))
               AND acptitemno     = CAST(? AS SMALLINT)
               AND testcd         = ?
               AND pid            = ? 
               AND prcpdd         = ?
               AND execprcpuniqno = CAST(? AS INTEGER)
        
     /* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml updacptstat */ |3 msec
52001286[node= user=10602673 ip=10.110.114.57][2017.07.24 14:26:13] [ INFO] [himed.his.lis.lib.plgycommonmgt.dao.PlgyCommonDAOImpl] query time = 2 msec.
52001286[node= user=10602673 ip=10.110.114.57][2017.07.24 14:26:13] [ INFO] [himed.his.lis.lib.plgycommonmgt.dao.PlgyCommonDAOImpl] execute query |2 msec|param=[017, 17488137, 20170724, 1151787391]|1 records|
        
            SELECT COUNT(distinct acptstatcd) AS acptststcnt , COUNT(distinct ptnocd) AS ptnocd 
              FROM lis.lpjmacpt 
             WHERE instcd          = ?
               AND pid             = ? 
               AND prcpdd          = ?
               AND execprcpuniqno  = CAST(? AS INTEGER)
               AND acptstatcd     IN ('0', '2', '3', '4', '9')
             GROUP BY instcd, pid, prcpdd, execprcpuniqno
        
     /* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getprcpacptstatref */|2 msec
52001286[node= user=10602673 ip=10.110.114.57][2017.07.24 14:26:13] [ INFO] [himed.his.lis.lib.plgycommonmgt.dao.PlgyCommonDAOImpl] query time = 1 msec.
52001286[node= user=10602673 ip=10.110.114.57][2017.07.24 14:26:13] [ INFO] [himed.his.lis.lib.plgycommonmgt.dao.PlgyCommonDAOImpl] execute query |1 msec|param=[017, 17488137, 20170724, 1151787391]|1 records|
            
            SELECT b.prcpstatcd
        
        
                
              FROM emr.mmodexip a, emr.mmohiprc b
            
         
            
             WHERE a.instcd         = ?
               AND a.pid            = ?
               AND a.prcpdd         = ?
               AND a.execprcpuniqno = ?
               AND a.execprcphistcd = 'O'
               AND a.instcd         = b.instcd
               AND a.prcpdd         = b.prcpdd
               AND a.prcpno         = b.prcpno
               AND a.prcphistno     = b.prcphistno
               AND b.prcphistcd     = 'O'
               AND b.prcpclscd      = 'D2'
               AND b.tempprcpflag   = 'N'
        
     /* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getprcpstat */|1 msec
52001286[node= user=10602673 ip=10.110.114.57][2017.07.24 14:26:13] [ INFO] [himed.his.lis.lib.plgycommonmgt.dao.PlgyCommonDAOImpl] execute update |1 msec|param=[710, 10602673, 017, 17488137, 20170724, 1151787391]|1 records| 
            
            UPDATE 
        
        
                
                   emr.mmohiprc
            
         
            
               SET prcpstatcd  = ?,
                   lastupdtdt  = SYSDATE,
                   lastupdtrid = ?
             WHERE (instcd, pid, prcpdd, prcpno, prcphistno) IN
                   (SELECT instcd, pid, prcpdd, prcpno, prcphistno
        
        
                
                      FROM emr.mmodexip
            
         
            
                     WHERE instcd         = ?
                       AND pid            = ?
                       AND prcpdd         = ?
                       AND execprcpuniqno = ?
                       AND execprcphistcd = 'O'
                   )
               AND prcphistcd   = 'O'
               AND prcpclscd    = 'D2'
               AND tempprcpflag = 'N'
        
     /* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml setprcpstat */ |1 msec
52001286[node= user=10602673 ip=10.110.114.57][2017.07.24 14:26:13] [ INFO] [himed.his.lis.lib.plgycommonmgt.dao.PlgyCommonDAOImpl] execute update |1 msec|param=[710, 10602673, 017, 17488137, 20170724, 1151787391]|1 records| 
            
            UPDATE 
        
        
                
                   emr.mmodexip a
            
         
            
               SET a.execprcpstatcd = ?,
        
        
            
                   a.lastupdtdt     = SYSDATE,
                   a.lastupdtrid    = ?
             WHERE a.instcd         = ?
               AND a.pid            = ?
               AND a.prcpdd         = ?
               AND a.execprcpuniqno = ?
               AND a.execprcphistcd = 'O'
        
     /* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml setexecprcpstat */ |1 msec
52001286[node= user=10602673 ip=10.110.114.57][2017.07.24 14:26:13] [ INFO] [himed.his.lis.lib.plgycommonmgt.dao.PlgyCommonDAOImpl] query time = 1 msec.
52001286[node= user=10602673 ip=10.110.114.57][2017.07.24 14:26:13] [ INFO] [himed.his.lis.lib.plgycommonmgt.dao.PlgyCommonDAOImpl] execute query |1 msec|param=[017, 1151787391, 20170724, 710]|1 records|
            
            SELECT COUNT(prcpdd) AS tretcnt
              FROM emr.mmodexpt
             WHERE instcd         = ?
               AND execprcpuniqno = ?
               AND prcpdd         = ?
               AND tretflagcd     = ?
        
     /* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getDoPrcpTret */|1 msec
52001286[node= user=10602673 ip=10.110.114.57][2017.07.24 14:26:13] [ INFO] [himed.his.lis.lib.plgycommonmgt.dao.PlgyCommonDAOImpl] execute update |9 msec|param=[20170724, 1151787391, 710, 017, 20170724, 142613, null, 10602673, null, 10602673, 10602673]|1 records| 
            
            INSERT INTO emr.mmodexpt (prcpdd,       execprcpuniqno, 
                                      tretflagcd,   instcd,            
                                      tretdd,       trettm,    tretpsnid, 
                                      fstrgstrid,   fstrgstdt, 
                                      lastupdtrid,  lastupdtdt)
                              VALUES (?,     CAST(? AS INTEGER), 
                                      ?,  ?,  
                                      ?,     ?,  CASE WHEN ? IS NULL THEN ? ELSE ? END,
                                      ?,    SYSDATE,
                                      ?,    SYSDATE) 
        
     /* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml insDoPrcpTret */ |9 msec