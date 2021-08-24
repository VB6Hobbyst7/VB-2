/*  -------------------------------------------------------------
 *  Copyright (c) 2004  SAMSUNG SDS Co.,Ltd. All rights reserved.
 *  -------------------------------------------------------------
 *
 *  Service ID : SL_USERM_L1
 *
 *  Specification
 *    1. 사용자정보조회
 *    !
 *
 *  Related Tables
 *    1. cdb..ccusermt
 *    2. cdb..ccdeptct
 *    3. cdb..sp_cc_decrypt_password
 *
 *  Input Parameters
 *    1.  :          ( FML_BUF : DB_COLUMN )
 *    2.  :
 *
 *  Output Parameters
 *
 *    1.     :                 ( FML_BUF : DB_COLUMN )
 *    2.     :
 *
 *  Modification Log
 *      ==========================================================
 *    #   Date         Author                          EditLabel
 *      ----------------------------------------------------------
 *       Description
 *      ==========================================================
 *    0. 2004-12-06 13:58:25  Manwoong Park           C041206-
 *        Create
 *
 */


#include <global.h>           /* Mandatory */

exec sql include sqlca;

#define  MAXROWCNT 200

#define  DEBUG 0

SL_USERM_L1(TPSVCINFO *msg)
{
   FBFR *transf;        /* Fixed Variable */
   char message[MSGLEN];   /* Fixed Variable */

   exec sql begin declare section;
      CS_CHAR   in_flag  [10+1];
      CS_CHAR   in_id    [20+1];
      CS_CHAR   in_locate[10+1];
      CS_CHAR   sUserid  [MAXROWCNT][20+1];
      CS_CHAR   sUsername[MAXROWCNT][50+1];
      CS_CHAR   sPassword[MAXROWCNT][20+1];
      CS_CHAR   sPwd     [MAXROWCNT][20+1];
      CS_CHAR   sPubpwd  [MAXROWCNT][20+1];
      CS_CHAR   sPwd2    [MAXROWCNT][20+1];
      CS_CHAR   sDpcd    [MAXROWCNT][10+1];
      CS_CHAR   sUserstat[MAXROWCNT][5+1];
      CS_CHAR   sInptdept[MAXROWCNT][5+1];
      CS_CHAR   sJikmu   [MAXROWCNT][5+1];
      CS_CHAR   sLicno1  [MAXROWCNT][20+1];
      CS_CHAR   sLicno2  [MAXROWCNT][20+1];
      CS_CHAR   sLocate  [MAXROWCNT][10+1];
      CS_CHAR   sDeptname[MAXROWCNT][25+1];
      CS_CHAR   sJikjong [MAXROWCNT][5+1];
      CS_CHAR   sWkareacd[MAXROWCNT][10+1];
      CS_CHAR   sWkareanm[MAXROWCNT][25+1];

      int       retcode;
   exec sql end declare section;

   int      rowcnt=0,ii=0,ix=0,iy=0;
   int      SendCnt=0;

   applog(msg);            /* Mandatory:Write Svc Info To Trace File */


   /* initialize variable */
   memset(in_flag  , 0, sizeof(in_flag  ));
   memset(in_id    , 0, sizeof(in_id    ));
   memset(in_locate, 0, sizeof(in_locate));
   memset(sUserid  , 0, sizeof(sUserid  ));
   memset(sUsername, 0, sizeof(sUsername));
   memset(sPassword, 0, sizeof(sPassword));
   memset(sPwd     , 0, sizeof(sPwd     ));
   memset(sPubpwd  , 0, sizeof(sPubpwd  ));
   memset(sPwd2    , 0, sizeof(sPwd2    ));
   memset(sDpcd    , 0, sizeof(sDpcd    ));
   memset(sUserstat, 0, sizeof(sUserstat));
   memset(sInptdept, 0, sizeof(sInptdept));
   memset(sJikmu   , 0, sizeof(sJikmu   ));
   memset(sLicno1  , 0, sizeof(sLicno1  ));
   memset(sLicno2  , 0, sizeof(sLicno2  ));
   memset(sLocate  , 0, sizeof(sLocate  ));
   memset(sDeptname, 0, sizeof(sDeptname));
   memset(sJikjong , 0, sizeof(sJikjong ));
   memset(sWkareacd, 0, sizeof(sWkareacd));
   memset(sWkareacd, 0, sizeof(sWkareanm));
   transf = (FBFR *)msg->data;   /* get input buffer Pointer */

#ifdef DEBUG
   Fprint(transf);               /* Print Input Buffer to stdout File */
#endif

   /* get Data From FML Buffer to variable */
   GET (S_FLAG2  , 0, in_flag  );
   GET (S_IDNUM1 , 0, in_id    );
   GET (S_FLAG1  , 0, in_locate);

   /* Main SQL Statement */
   if ( strcmp(in_flag, "ALL")== 0 ) {
      exec sql
         select
                u.userid
              , u.username
           into
                :sUserid
              , :sUsername
           from
                cdb..ccusermt u
          where
                u.dpcd like :in_id
            and u.locate =  isnull(rtrim(:in_locate), u.locate)
            and (   u.deldate is null
                 or u.deldate >= getdate())
            and getdate() between u.startdt and u.enddt
          order by
                   u.userid
               ;
   /* 해당사용자가 특정화면에 대한 권한이 있는지 check */
   } else if ( strcmp(in_flag, "FORM")== 0 ) {
      exec sql
         select
                count(*)
           into
               :sUserstat
           from
                cdb..csusgrpt u
               (index csusgrpt_pk)
               ,cdb..cspgrpdt p
               (index cspgrpdt_idx1)
          where
                u.userid  = :in_id
            and p.groupid = u.groupid
            and p.progid  = :in_locate
               ;
   } else {
      exec sql
         select
                u.userid
               ,u.username
               ,isnull(u.password, '')
               ,isnull(u.dpcd    , '')
               ,isnull(u.userstat, '')
               ,isnull(u.inptdept, '')
               ,isnull(u.jikmu   , '')
               ,isnull(u.licno1  , '')
               ,isnull(u.licno2  , '')
               ,isnull(u.locate  , '')
               ,isnull(d.deptnm  , '')
               ,isnull(u.jikjong , '')
               ,isnull(u.pubpwd  , '')
               ,isnull(u.wkareacd, '')
               ,isnull((select
			       deptnm
                        from cdb..ccdeptct
                       where dpcd = u.wkareacd
                         and locate = u.locate),'')
           into
               :sUserid
              ,:sUsername
              ,:sPassword
              ,:sDpcd
              ,:sUserstat
              ,:sInptdept
              ,:sJikmu
              ,:sLicno1
              ,:sLicno2
              ,:sLocate
              ,:sDeptname
              ,:sJikjong
              ,:sPubpwd
              ,:sWkareacd
              ,:sWkareanm
           from
                cdb..ccusermt u
               ,cdb..ccdeptct d
          where
                u.userid =  :in_id
            and u.locate =  isnull(rtrim(:in_locate), u.locate)
            and (   u.deldate is null
                 or u.deldate >= getdate())
            and getdate() between u.startdt and u.enddt
            and d.locate =* u.locate
            and d.dpcd   =* u.dpcd
               ;
   }
	/* SQL Statement Error Check */
	if  (SQLCODE!=0 && SQLCODE!=NOTFOUND && SQLCODE!=TOOMANY) {
		MakeErrMsg(message,SVCNAME,DBSEL,SQLCODE,SQLMSG);	/* C041022-1 */
		TPRETURN_ERROR(message,-1);
	}

	if  (SQLCODE==TOOMANY) {				/* Array Overflow Check */
		MakeMsg(message,SF_B220,MAXROWCNT);	/* 검색자료가 %d 건을 초과 */
		TPRETURN_ERROR(message,0);
	}

	if  (SQLCODE==NOTFOUND) {
		MakeMsg(message,SF_B120);			/* 조회할 자료가 없습니다 */
		TPRETURN_ERROR(message,0);
	}

	rowcnt = SQLCNT;

	/* FML Buffer realloc for data output */
	if  ((transf=(FBFR *)tprealloc((char *)transf,MAXBUFSIZE))==(FBFR *)NULL) {
		MakeErrMsg(message,SVCNAME,TUXALLOC,TPCODE,TPMSG);	/* C041021-1 */
		TPRETURN_ERROR(message,-1);
	}

	/* Fml buffer clear */
	if  (Finit(transf,Fsizeof(transf)) == -1) {
		MakeErrMsg(message,SVCNAME,TUXINIT,TPFCODE,TPFMSG);	/* C041021-1 */
		TPRETURN_ERROR(message,-1);
	}

   for (iy=0;iy<rowcnt;iy++,SendCnt++) {
      exec sql
         exec :retcode = cdb..sp_cc_decrypt_password
                            :sPassword[iy]
                           ,:sPwd     [iy]     out
             ;
      if (retcode < 0) {
         MakeErrMsg(message,SVCNAME,message,SQLCODE,SQLMSG);
         TPRETURN_ERROR(message,-1);
      }

      exec sql
         exec :retcode = cdb..sp_cc_decrypt_password
                            :sPubpwd[iy]
                           ,:sPwd2  [iy]     out
             ;
      if (retcode < 0) {
         MakeErrMsg(message,SVCNAME,message,SQLCODE,SQLMSG);
         TPRETURN_ERROR(message,-1);
      }

      SPUT (S_IDNUM1, sUserid   [iy]);
      SPUT (S_NAME1 , sUsername [iy]);
      SPUT (S_TEXT1 , sPwd      [iy]);
      SPUT (S_CODE1 , sDpcd     [iy]);
      SPUT (S_STAT1 , sUserstat [iy]);
      SPUT (S_CODE2 , sInptdept [iy]);
      SPUT (S_CODE3 , sJikmu    [iy]);
      SPUT (S_NO1   , sLicno1   [iy]);
      SPUT (S_NO2   , sLicno2   [iy]);
      SPUT (S_FLAG1 , sLocate   [iy]);
      SPUT (S_NAME2 , sDeptname [iy]);
      SPUT (S_CODE4 , sJikjong  [iy]);
      SPUT (S_TEXT2 , sPwd2     [iy]);
      SPUT (S_CODE5 , sWkareacd [iy]);
      SPUT (S_NAME3 , sWkareanm [iy]);

   }

   MakeMsg(message,SF_A220,SendCnt);   /* %d건의 자료가 조회되었습니다 */

   TPRETURN(message,0);
}

/*
 * End of Service
 */

/*
 * End of Source
 */