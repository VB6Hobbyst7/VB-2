/*    시스템 database 접속 정보와 병리과 사용자 확인 정보 전달 합니다.(패스워드는 암호화 제공이 되지 않습니다. 사원번호만으로 확인을 합니다.)
      성빈센트 ORACLE 로그인 아이디 / 패스워드 = lisu02 / lisu02 
1. Tnsname 접속 정보

# 성빈센트병원 운영(REAL)  이중화 1호기 실패시 2호기
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

# 성빈센트병원 교육(TEST)
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


/* -- 병리과 소속  사용자 로그인 정보 요청자 사번 20801950 / 이재왕
param=['20801950']
성빈센트병원 병리과 기준자료에 등록된 분들만 대상으로 조회됩니다.
*/ 
select usab.instcd, -- 기관기호 
       usab.userid, -- 사용자아이디
       usab.userabbr as usernm -- 사용자성명
  from lis.lpcmusab usab,  -- 병리과 사용자관리 
       com.zsumusrb usrb   -- 시스템 사용자정보
 where usab.instcd     = '017' 
   and usab.delflagcd  = '0'
   and usab.userid  = :arg_userid
   and usab.userid  = usrb.userid
   and to_char(sysdate,'yyyymmdd')    between usrb.userfromdd and usrb.usertodd

 

/* ==========
힌트와 binding 변수처리된 부분 지켜주세요. nU에 오라클에 Full scan 될경우 DBA에서 차단됩니다. 
힌트와 binding 변수처리된 부분 지켜주세요. nU에 오라클에 Full scan 될경우 DBA에서 차단됩니다.
힌트와 binding 변수처리된 부분 지켜주세요. nU에 오라클에 Full scan 될경우 DBA에서 차단됩니다.
힌트와 binding 변수처리된 부분 지켜주세요. nU에 오라클에 Full scan 될경우 DBA에서 차단됩니다.
힌트와 binding 변수처리된 부분 지켜주세요. nU에 오라클에 Full scan 될경우 DBA에서 차단됩니다.
힌트와 binding 변수처리된 부분 지켜주세요. nU에 오라클에 Full scan 될경우 DBA에서 차단됩니다.
힌트와 binding 변수처리된 부분 지켜주세요. nU에 오라클에 Full scan 될경우 DBA에서 차단됩니다.
힌트와 binding 변수처리된 부분 지켜주세요. nU에 오라클에 Full scan 될경우 DBA에서 차단됩니다.
==========*/ 


/* -- 병리과에 접수된 오더조회 병리과 분자병리 처방 PMO12040 HPV genotyping real-time PCR 에 제한된 특화처리 */
/* -- 병리과에 접수된 오더조회 병리과 분자병리 처방 PMO12040 HPV genotyping real-time PCR 에 제한된 특화처리 */
/* -- 병리과에 접수된 오더조회 병리과 분자병리 처방 PMO12040 HPV genotyping real-time PCR 에 제한된 특화처리 */
/* -- 병리과에 접수된 오더조회 병리과 분자병리 처방 PMO12040 HPV genotyping real-time PCR 에 제한된 특화처리 */
/* -- 병리과에 접수된 오더조회 병리과 분자병리 처방 PMO12040 HPV genotyping real-time PCR 에 제한된 특화처리 */
/* param=[017, to_char(sysdate-1,'yyyymmdd'), to_char(sysdate,'yyyymmdd')]
어제와 오늘일자 병리과에서 접수된 오더가 조회됩니다. 조회화면 처리 또는 아래내역 처리시 모두 필요한 값들입니다.
*/ 
select /*+ leading(acpt) */
    acpt.instcd,        --기관기호
    acpt.acptdd,        --접수일자
    acpt.acptno,        --접수번호
    acpt.acptitemno,    --접수항목번호
    acpt.ptno,          --병리번호
    acpt.acptstatcd	    --접수상태코드(0:접수(440),1:취소(100,240),2:예비결과(710),3:최종결과(730), 4:수정결과(740))
    acpt.pid,           --등록번호
    acpt.testcd,        --검사코드
    test.testengnm,     --영문검사명
    spcm.spcnm,         --검체명
    acpt.prcpgenrflag,  --입원/외래구분
    acpt.orddeptcd,     --진료과코드
    dept.deptengabbr,   --진료과명
    acpt.prcpdd,        --처방일자,
    acpt.execprcpuniqno, --실시처방유일번호
    acpt.prcpno,        --처방번호
    ptbs.hngnm,         --환자명
    ptbs.sex,           --성별
    ptbs.brthdd,        --생일
    com.fn_zz_getage(ptbs.rrgstno1, ptbs.rrgstno2, acpt.acptdd, 'A', ptbs.brthdd) as age  -- 접수일자기준 나이
from lis.lpjmacpt acpt, lis.lpcmtest test, lis.lpcmspcm spcm, pam.pmcmptbs ptbs, com.zsdddept dept
where acpt.instcd = :arg_instcd --'성빈센트병원(017)는 고정입니다.
and acpt.acptdd between  :arg_acptfromdd and :arg_acpttodd 
and acpt.testcd = 'PMO12040' -- 검사 처방코드 PMO12040 고정 GFX96 에서 다른처방 처리시 in 절로 변경
and acpt.acptstatcd = '0' --접수상태코드(0:접수,1:취소,2:예비결과,3:최종결과, 4:수정결과 고정
and acpt.instcd          = test.instcd
and acpt.testcd          = test.testcd
and acpt.instcd          = spcm.instcd
and acpt.spccd           = spcm.spccd
and acpt.instcd          = ptbs.instcd
and acpt.pid             = ptbs.pid
and acpt.instcd          = dept.instcd
and acpt.orddeptcd       = dept.deptcd
-- 결과 저장전에 상태체크 sql  사용시 변수
-- and acpt.instcd         = :arg_instcd
-- and acpt.ptno           = :arg_ptno
-- and acpt.pid            = :arg_pid
-- and acpt.acptstatcd    in ('0','2')
and acpt.prcpdd between dept.valifromdd and dept.valitodd



/* acpt.acptstatcd -접수상태코드(0:접수(440),1:취소(100,240),2:예비결과(710),3:최종결과(730), 4:수정결과(740)) */
/* acpt.acptstatcd 사용자 실수로 결과 인터페이스가 되더라도... (0:접수(440) 에서는 신규로 결과입력 , 2:예비결과(710) 결과 update 이상태에서만 작동이 되야합니다.*/
/* -- 병리과결과 처리 실행 병리과 분자병리 처방 PMO12040 HPV genotyping real-time PCR 에 제한된 특화처리 */
/* -- 병리과결과 처리 실행 병리과 분자병리 처방 PMO12040 HPV genotyping real-time PCR 에 제한된 특화처리 */
/* -- 병리과결과 처리 실행 병리과 분자병리 처방 PMO12040 HPV genotyping real-time PCR 에 제한된 특화처리 */
/* -- 병리과결과 처리 실행 병리과 분자병리 처방 PMO12040 HPV genotyping real-time PCR 에 제한된 특화처리 */
/* -- 병리과결과 처리 실행 병리과 분자병리 처방 PMO12040 HPV genotyping real-time PCR 에 제한된 특화처리 */
/* -- 병리과결과 처리 실행 병리과 분자병리 처방 PMO12040 HPV genotyping real-time PCR 에 제한된 특화처리 */
/* -- 병리과결과 처리 실행 끝까지 */

            SELECT TO_CHAR(SYSDATE,'YYYYMMDD') AS sysdd,
                   TO_CHAR(SYSDATE,'HH24MISS') AS systm
              FROM DUAL
              
    			prcpVO.set(iRowNo, "tretdd",         현재일자 yyyymmdd 예시 20170811 );
    			prcpVO.set(iRowNo, "trettm",         현재시분 hhs4miss 예시 090000 );
    			prcpVO.set(iRowNo, "readid",         로그인사번);

/*  결과순번 일련번호 채번을 위해 기준 Row를 Lock처리한다. 
himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml setseqnolock 
arg_seqflagcd = '4'  순번발생구분코드(검사결과) 고정입니다.
param=[017, 4] */
update lis.lpcmseqn
       set lastgenrno = 1
 where instcd        = :arg_instcd   
   and seqgenryy = '1900'
   and seqflagcd  = :arg_seqflagcd
        
/* 결과순번 채번을 합니다. 여기서 조회된 lastgenrno가 뒤쪽에 rsltrgstno 로 처리됩니다.
himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getlastseqno 
param=[017, 2017, 4]*/
select coalesce(lastgenrno+1, 1) as lastgenrno
  from lis.lpcmseqn
 where instcd     = :arg_instcd
   and seqgenryy  = :arg_seqgenryy 
   and seqflagcd  = :arg_seqflagcd

/* 위에서 채번을 했는데 null 일경우 insert하고 lastgenrno = 1로 리턴 1부터 시작합니다. 년도마다 새로 입력이 됩니다. */
/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml setseqnogenr 
param=[2017, 4, 017, 4, 로그인사번, 로그인사번]*/
            INSERT INTO lis.lpcmseqn (seqgenryy, seqflagcd, instcd, lastgenrno, 
                                                       fstrgstdt,      fstrgstrid,
                                                       lastupdtdt,   lastupdtrid)
                                         VALUES (:arg_seqgenryy, :arg_seqflagcd, :arg_instcd, 1, SYSDATE, :arg_userid,  SYSDATE, :arg_userid)
                                  
/* 위에서 채번을 했는데 null 일 아닐 경우 */                     
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
		    			prcpVO.set(iRowNo, "newprcpstatcd",  "710");	// 예비결과
		    			prcpVO.set(iRowNo, "newacptstatcd",  "2");	    // 예비결과
		    			prcpVO.set(iRowNo, "bizflagcd",      "710");	// 
		    			prcpVO.set(iRowNo, "biztretflagcd",  "I");		//
		    			prcpVO.set(iRowNo, "truststatcd",  "4");		// CMC 위수탁상태
*/

/* 검사결과(Header) 등록  저장합니다.
himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml instestrslt 
param=[M17003176, 20170724, 32715, 017, 
17488137, 142613, null, null, null, 
HPV High Risk Type : Positive (18+, 68+, 31+++)
HPV Low  Risk Type : Positive (70+, 61+) , 
null, 
null, null, 세포검체 other, HPV genotyping real-time PCR, null, 0, 
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
        
/* 검사결과(Header) 등록  저장합니다 신규저장이 아닐경우 이력을 남깁니다.*/
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
               
/* 검사결과내용 등록 합니다. 
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
인유두종 바이러스 (Human papilloma virus)는 자궁경부암의 주요 위험인자로 알려져 있습니다. 암과의 관련성 정도에 따라 고위험군 (high risk)과 저위험군 (low risk)로 구분되며, 저위험군 HPV는 대개 시간이 지나면 소실되거나 사마귀 등의 양성변변의 원인이 되는 반면, 고위험군은 자궁경부암을 유발시키는데 관여합니다 (N Engl J Med. 2003 348:518).

◆ 본 제품은 19종의 고위험군 HPV (16, 18, 26, 31, 33, 25, 29, 45, 51, 52, 53, 56, 58, 59, 66, 68, 73, 82형)과 9종의 저위험군 HPV (6, 11, 40, 42, 43, 44, 54, 61, 70형), 내부대조군의 타켓핵산을 검출합니다. 
◆ Viral load에서 +++:10^5 copies/reaction, ++:10^5~10^2 copies/reaction, +:10^2 copies/reaction의 농도로 해석될 수 있습니다. 이 중 “+”는 매우 낮은 농도로 감염 시기, 검체 채취 상태에 따라 반복 검사 시 재현되지 않을 수 있습니다.
◆ PCR 검사는 검체 내 균수가 적거나 부적절한 검체 의석 또는 증폭 억제물질이 존재하는 경우 위음성이 나올 수 있습니다. 또한, PCR 검사는 유전자 유무를 검사하므로 생존균과 사균의 구분이 안되어 위양성의 가능성이 있습니다. 결과 해석 시 임상 양상과 연관지어 판단하시기 바랍니다.
◆ 상기 검사는 검사 방법, 시약의 정도관리 및 검사 결과가 병리과 전문의에 의해 확인되었습니다.
   (검사 담당자: 이재왕), 
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

/* 검사결과내용 등록 합니다.  이력을 남깁니다.
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

/* T/M/P 진단 설정
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
        
/* 병리번호 발행이력 수정 
himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml updlpcmpnis  
param=[세포검체 other, 
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

/* 병리과 접수정보 수정
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
                 MIN(acpt.prcpgenrflag) AS prcpgenrflag,        -- 입원외래구분 값  
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
              FROM emr.mmodexip a, emr.mmohiprc b   -- 입원
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
              FROM emr.mmodexop a, emr.mmohoprc b   -- 외래
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
--입원일경우 prcpgenrflag = I, D, E
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

--외래일경우 prcpgenrflag = O, S 
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
--입원일경우 prcpgenrflag = I, D, E
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

--외래일경우 prcpgenrflag = O, S 
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
        
/*  getDoPrcpTret  COUNT(prcpdd) < 0 일경우에만   tretflagcd = 710  처방상태 예비결과 (저장된상태)
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


/* -- 병리과결과 처리 실행 끝*/
/* -- 병리과결과 처리 실행 끝*/
/* -- 병리과결과 처리 실행 끝*/
/* -- 병리과결과 처리 실행 끝*/
/* -- 병리과결과 처리 실행 끝*/
/* -- 병리과결과 처리 실행 끝*/
/* -- 병리과결과 처리 실행 끝*/     


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