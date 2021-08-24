/*	Copyright (c) 1994 Novell
	All rights reserved

	THIS IS UNPUBLISHED PROPRIETARY
	SOURCE CODE OF NOVELL
	The copyright notice above does not
	evidence any actual or intended
	publication of such source code.

#ident	"@(#) tuxedo/include/tpadm	$Revision: 1.1.8.1 $"
*/
/*
	Warning: This file should not be changed in any
	way, doing so will destroy the compatibility with TUXEDO programs
	and libraries.
*/
#ifndef _TPADM_H
#define _TPADM_H
#ifndef NOWHAT
static	char h_tpadm[] = "@(#) tuxedo/include/tpadm	$Revision: 1.1.8.1 $";
#endif
#ifndef TMENV_H
#include <tmenv.h>
#endif
#ifndef ATMI_H
#include <atmi.h>
#endif

#include <fml32.h>

/*	User level interface function prototypes */

#if defined(__cplusplus)
extern "C" {
#endif

extern int	tpadmcall _((FBFR32 *, FBFR32 **, long));

#if defined(__cplusplus)
}
#endif

/*	ADMIN FIELD ENTRIES */

/*
 * Flags values stored in TA_FLAGS, bit-wise flag map.
 *
 * Flag names beginning with MIB_ are generic to all component MIBs and
 * should not be reused or overlapped.  These values should also not be
 * changed as changes will affect release interoperability.
 *
 * Flag names beginning with other values (e.g., TMIB_ or QMIB_ are
 * specific to particular component MIBs and may be overlapped to conserve
 * the "name" space of the TA_FLAGS values.  Be sure when overlapping and
 * defining component MIB flag values to avoid overlapping with MIB_ values.
 */
#define MIB_PREIMAGE	0x00000001
#define MIB_LOCAL	0x00010000
#define MIB_SELF	0x00100000

/*#define TMIB_RSRVD1	0x00010000*/
#define TMIB_NOTIFY	0x00020000
#define TMIB_ADMONLY	0x00040000
#define TMIB_CONFIG	0x00080000
/*#define TMIB_RSRVD2	0x00100000*/
#define TMIB_APPONLY	0x00200000

/*#define QMIB_RSRVD1		0x00010000*/
#define QMIB_FORCECLOSE	0x00020000
#define QMIB_FORCEPURGE	0x00040000
#define QMIB_FORCEDELETE	0x00080000
/*#define QMIB_RSRVD2		0x00100000*/

#define MIB_ALLFLAGS	(MIB_PREIMAGE|MIB_LOCAL|MIB_SELF|TMIB_NOTIFY|\
			 TMIB_ADMONLY|TMIB_CONFIG|TMIB_APPONLY|\
			 QMIB_FORCECLOSE|QMIB_FORCEPURGE|QMIB_FORCEDELETE)\

/* TA_ATTFLAGS values, bit-wise flag map */
#define MIBATT_KEYFIELD	0x00000001
#define MIBATT_LOCAL		0x00000002
#define MIBATT_REGEXKEY	0x00000004
#define MIBATT_REQUIRED	0x00000008
#define MIBATT_SETKEY		0x00000010
#define MIBATT_NEWONLY		0x00000020
#define MIBATT_RUNTIME		0x00000040


/* Error Return codes set in TA_ERROR */

#define TAEAPP		-1	/* App component error during MIB processing */
#define TAECONFIG	-2	/* Operating system error */
#define TAEINVAL	-3	/* invalid argument */
#define TAEOS		-4	/* Operating system error */
#define TAEPERM	-5	/* permission error */
#define TAEPREIMAGE	-6	/* preimage does not match current image */
#define TAEPROTO	-7	/* MIB specific protocol error */
#define TAEREQUIRED	-8	/* field value required but not present */
#define TAESUPPORT	-9	/* Documented but unsupported feature */
#define TAESYSTEM	-10	/* Internal System/T error */
#define TAEUNIQ	-11	/* SET did not specify unique class instance */

/* Non-Error Return codes set in TA_ERROR */

#define TAOK		0	/* suceeded */
#define TAUPDATED	1	/* suceeded and updated a record */
#define TAPARTIAL	2	/* suceeded at master; failed elsewhere */

/*	fname	fldid            */
/*	-----	-----            */
#define	TA_ATTRIBUTE	((FLDID32)33560432)	/* number: 6000	 type: long */
#define	TA_BADFLD	((FLDID32)33560433)	/* number: 6001	 type: long */
#define	TA_CLASS	((FLDID32)167778162)	/* number: 6002	 type: string */
#define	TA_CLASSNAME	((FLDID32)167778163)	/* number: 6003	 type: string */
#define	TA_CURSOR	((FLDID32)167778164)	/* number: 6004	 type: string */
#define	TA_CURSORHOLD	((FLDID32)33560437)	/* number: 6005	 type: long */
#define	TA_ERROR	((FLDID32)33560438)	/* number: 6006	 type: long */
#define	TA_FILTER	((FLDID32)33560439)	/* number: 6007	 type: long */
#define	TA_FLAGS	((FLDID32)33560440)	/* number: 6008	 type: long */
#define	TA_MIBTIMEOUT	((FLDID32)33560441)	/* number: 6009	 type: long */
#define	TA_MORE	((FLDID32)33560442)	/* number: 6010	 type: long */
#define	TA_OCCURS	((FLDID32)33560443)	/* number: 6011	 type: long */
#define	TA_OPERATION	((FLDID32)167778172)	/* number: 6012	 type: string */
#define	TA_PERM	((FLDID32)33560445)	/* number: 6013	 type: long */
#define	TA_STATE	((FLDID32)167778174)	/* number: 6014	 type: string */
#define	TA_STATUS	((FLDID32)167778175)	/* number: 6015	 type: string */

/*
	The following fields are TM_MIB(5) specific fields.  Consult TM_MIB(5)
	for details on each field and its use.
*/

#ifndef REDUCE_CPP
/*
	The field numbers below should always begin at 0 and increase.
	Field numbers cannot be reused or changed from release to release or
	interoperability will be broken.  Fields 0-244 (base 6200) defined in
	release 5.0, fields defined in release 6.0 begin at field 245 and
	increase from there.

	Note that fields 7200-7299 are reserved for use by TUXEDO OEMs.
*/
#define	TA_ACCWORD	((FLDID32)33560632)	/* number: 6200	 type: long */
#define	TA_APPDIR	((FLDID32)167778361)	/* number: 6201	 type: string */
#define	TA_AUTHSVC	((FLDID32)167778362)	/* number: 6202	 type: string */
#define	TA_AUTOTRAN	((FLDID32)167778363)	/* number: 6203	 type: string */
#define	TA_BASESRVID	((FLDID32)33560636)	/* number: 6204	 type: long */
#define	TA_BBLQUERY	((FLDID32)33560637)	/* number: 6205	 type: long */
#define	TA_BLOCKTIME	((FLDID32)33560638)	/* number: 6206	 type: long */
#define	TA_BRIDGE	((FLDID32)167778367)	/* number: 6207	 type: string */
#define	TA_BUFTYPE	((FLDID32)167778368)	/* number: 6208	 type: string */
#define	TA_CFGDEVICE	((FLDID32)167778369)	/* number: 6209	 type: string */
#define	TA_CFGOFFSET	((FLDID32)33560642)	/* number: 6210	 type: long */
#define	TA_CLIENTID	((FLDID32)167778371)	/* number: 6211	 type: string */
#define	TA_CLOPT	((FLDID32)167778372)	/* number: 6212	 type: string */
#define	TA_CLOSEINFO	((FLDID32)167778373)	/* number: 6213	 type: string */
#define	TA_CLTLMID	((FLDID32)167778374)	/* number: 6214	 type: string */
#define	TA_CLTNAME	((FLDID32)167778375)	/* number: 6215	 type: string */
#define	TA_CLTPID	((FLDID32)33560648)	/* number: 6216	 type: long */
#define	TA_CLTREPLY	((FLDID32)167778377)	/* number: 6217	 type: string */
#define	TA_CMPLIMIT	((FLDID32)167778378)	/* number: 6218	 type: string */
#define	TA_CMTRET	((FLDID32)167778379)	/* number: 6219	 type: string */
#define	TA_CONNOGRPNO	((FLDID32)33560652)	/* number: 6220	 type: long */
#define	TA_CONNOLMID	((FLDID32)167778381)	/* number: 6221	 type: string */
#define	TA_CONNOPID	((FLDID32)33560654)	/* number: 6222	 type: long */
#define	TA_CONNOREGIDX	((FLDID32)33560655)	/* number: 6223	 type: long */
#define	TA_CONNOSNDCNT	((FLDID32)33560656)	/* number: 6224	 type: long */
#define	TA_CONNOSRVID	((FLDID32)33560657)	/* number: 6225	 type: long */
#define	TA_CONNSGRPNO	((FLDID32)33560658)	/* number: 6226	 type: long */
#define	TA_CONNSLMID	((FLDID32)167778387)	/* number: 6227	 type: string */
#define	TA_CONNSPID	((FLDID32)33560660)	/* number: 6228	 type: long */
#define	TA_CONNSSNDCNT	((FLDID32)33560661)	/* number: 6229	 type: long */
#define	TA_CONNSSRVID	((FLDID32)33560662)	/* number: 6230	 type: long */
#define	TA_CONTIME	((FLDID32)33560663)	/* number: 6231	 type: long */
#define	TA_CONV	((FLDID32)167778392)	/* number: 6232	 type: string */
#define	TA_COORDLMID	((FLDID32)167778393)	/* number: 6233	 type: string */
#define	TA_CURACCESSERS	((FLDID32)33560666)	/* number: 6234	 type: long */
#define	TA_CURCLIENTS	((FLDID32)33560667)	/* number: 6235	 type: long */
#define	TA_CURCONV	((FLDID32)33560668)	/* number: 6236	 type: long */
#define	TA_CURDRT	((FLDID32)33560669)	/* number: 6237	 type: long */
#define	TA_CURGROUPS	((FLDID32)33560670)	/* number: 6238	 type: long */
#define	TA_CURGTT	((FLDID32)33560671)	/* number: 6239	 type: long */
#define	TA_CURLMID	((FLDID32)167778400)	/* number: 6240	 type: string */
#define	TA_CURMACHINES	((FLDID32)33560673)	/* number: 6241	 type: long */
#define	TA_CURQUEUES	((FLDID32)33560674)	/* number: 6242	 type: long */
#define	TA_CURREQ	((FLDID32)33560675)	/* number: 6243	 type: long */
#define	TA_CURRFT	((FLDID32)33560676)	/* number: 6244	 type: long */
#define	TA_CURRLOAD	((FLDID32)33560677)	/* number: 6245	 type: long */
#define	TA_CURRSERVICE	((FLDID32)167778406)	/* number: 6246	 type: string */
#define	TA_CURRTDATA	((FLDID32)33560679)	/* number: 6247	 type: long */
#define	TA_CURSERVERS	((FLDID32)33560680)	/* number: 6248	 type: long */
#define	TA_CURSERVICES	((FLDID32)33560681)	/* number: 6249	 type: long */
#define	TA_CURSTYPE	((FLDID32)33560682)	/* number: 6250	 type: long */
#define	TA_CURTIME	((FLDID32)33560683)	/* number: 6251	 type: long */
#define	TA_CURTYPE	((FLDID32)33560684)	/* number: 6252	 type: long */
#define	TA_CURWSCLIENTS	((FLDID32)33560685)	/* number: 6253	 type: long */
#define	TA_DBBLWAIT	((FLDID32)33560686)	/* number: 6254	 type: long */
#define	TA_DEBUG	((FLDID32)167778415)	/* number: 6255	 type: string */
#define	TA_DEVICE	((FLDID32)167778416)	/* number: 6256	 type: string */
#define	TA_DEVINDEX	((FLDID32)33560689)	/* number: 6257	 type: long */
#define	TA_DEVOFFSET	((FLDID32)33560690)	/* number: 6258	 type: long */
#define	TA_DEVSIZE	((FLDID32)33560691)	/* number: 6259	 type: long */
#define	TA_DOMAINID	((FLDID32)167778420)	/* number: 6260	 type: string */
#define	TA_ENDTIME	((FLDID32)33560693)	/* number: 6261	 type: long */
#define	TA_ENVFILE	((FLDID32)167778422)	/* number: 6262	 type: string */
#define	TA_FIELD	((FLDID32)167778423)	/* number: 6263	 type: string */
#define	TA_FLOWCNT	((FLDID32)33560696)	/* number: 6264	 type: long */
#define	TA_FREEMAPAVAIL	((FLDID32)33560697)	/* number: 6265	 type: long */
#define	TA_FREEMAPCOUNT	((FLDID32)33560698)	/* number: 6266	 type: long */
#define	TA_FREEMAPINDEX	((FLDID32)33560699)	/* number: 6267	 type: long */
#define	TA_FREEMAPOFFSET	((FLDID32)33560700)	/* number: 6268	 type: long */
#define	TA_FREEMAPSIZE	((FLDID32)33560701)	/* number: 6269	 type: long */
#define	TA_GENERATION	((FLDID32)33560702)	/* number: 6270	 type: long */
#define	TA_GID	((FLDID32)33560703)	/* number: 6271	 type: long */
#define	TA_GRACE	((FLDID32)33560704)	/* number: 6272	 type: long */
#define	TA_GRPCOUNT	((FLDID32)33560705)	/* number: 6273	 type: long */
#define	TA_GRPINDEX	((FLDID32)33560706)	/* number: 6274	 type: long */
#define	TA_GRPNO	((FLDID32)33560707)	/* number: 6275	 type: long */
#define	TA_GSTATE	((FLDID32)167778436)	/* number: 6276	 type: string */
#define	TA_GTRID	((FLDID32)167778437)	/* number: 6277	 type: string */
#define	TA_HITICKET	((FLDID32)33560710)	/* number: 6278	 type: long */
#define	TA_HWACCESSERS	((FLDID32)33560711)	/* number: 6279	 type: long */
#define	TA_HWCLIENTS	((FLDID32)33560712)	/* number: 6280	 type: long */
#define	TA_HWCONV	((FLDID32)33560713)	/* number: 6281	 type: long */
#define	TA_HWDRT	((FLDID32)33560714)	/* number: 6282	 type: long */
#define	TA_HWGROUPS	((FLDID32)33560715)	/* number: 6283	 type: long */
#define	TA_HWGTT	((FLDID32)33560716)	/* number: 6284	 type: long */
#define	TA_HWMACHINES	((FLDID32)33560717)	/* number: 6285	 type: long */
#define	TA_HWQUEUES	((FLDID32)33560718)	/* number: 6286	 type: long */
#define	TA_HWRFT	((FLDID32)33560719)	/* number: 6287	 type: long */
#define	TA_HWRTDATA	((FLDID32)33560720)	/* number: 6288	 type: long */
#define	TA_HWSERVERS	((FLDID32)33560721)	/* number: 6289	 type: long */
#define	TA_HWSERVICES	((FLDID32)33560722)	/* number: 6290	 type: long */
#define	TA_HWWSCLIENTS	((FLDID32)33560723)	/* number: 6291	 type: long */
#define	TA_IPCKEY	((FLDID32)33560724)	/* number: 6292	 type: long */
#define	TA_ITERATION	((FLDID32)33560725)	/* number: 6293	 type: long */
#define	TA_LASTGRP	((FLDID32)33560726)	/* number: 6294	 type: long */
#define	TA_LDBAL	((FLDID32)167778455)	/* number: 6295	 type: string */
#define	TA_LICEXPIRE	((FLDID32)167778456)	/* number: 6296	 type: string */
#define	TA_LICMAXUSERS	((FLDID32)33560729)	/* number: 6297	 type: long */
#define	TA_LICSERIAL	((FLDID32)167778458)	/* number: 6298	 type: string */
#define	TA_LOAD	((FLDID32)33560731)	/* number: 6299	 type: long */
#define	TA_MASTER	((FLDID32)167778460)	/* number: 6300	 type: string */
#define	TA_MAX	((FLDID32)33560733)	/* number: 6301	 type: long */
#define	TA_MAXACCESSERS	((FLDID32)33560734)	/* number: 6302	 type: long */
#define	TA_MAXBUFSTYPE	((FLDID32)33560735)	/* number: 6303	 type: long */
#define	TA_MAXBUFTYPE	((FLDID32)33560736)	/* number: 6304	 type: long */
#define	TA_MAXCONV	((FLDID32)33560737)	/* number: 6305	 type: long */
#define	TA_MAXDRT	((FLDID32)33560738)	/* number: 6306	 type: long */
#define	TA_MAXGEN	((FLDID32)33560739)	/* number: 6307	 type: long */
#define	TA_MAXGROUPS	((FLDID32)33560740)	/* number: 6308	 type: long */
#define	TA_MAXGTT	((FLDID32)33560741)	/* number: 6309	 type: long */
#define	TA_MAXMACHINES	((FLDID32)33560742)	/* number: 6310	 type: long */
#define	TA_MAXMTYPE	((FLDID32)33560743)	/* number: 6311	 type: long */
#define	TA_MAXQUEUES	((FLDID32)33560744)	/* number: 6312	 type: long */
#define	TA_MAXRFT	((FLDID32)33560745)	/* number: 6313	 type: long */
#define	TA_MAXRTDATA	((FLDID32)33560746)	/* number: 6314	 type: long */
#define	TA_MAXSERVERS	((FLDID32)33560747)	/* number: 6315	 type: long */
#define	TA_MAXSERVICES	((FLDID32)33560748)	/* number: 6316	 type: long */
#define	TA_MAXWSCLIENTS	((FLDID32)33560749)	/* number: 6317	 type: long */
#define	TA_MIBMASK	((FLDID32)33560750)	/* number: 6318	 type: long */
#define	TA_MIN	((FLDID32)33560751)	/* number: 6319	 type: long */
#define	TA_MINOR	((FLDID32)33560752)	/* number: 6320	 type: long */
#define	TA_MMDDYY	((FLDID32)33560753)	/* number: 6321	 type: long */
#define	TA_MODEL	((FLDID32)167778482)	/* number: 6322	 type: string */
#ifndef REDUCE_CPP_NOIPC
#define	TA_MSGID	((FLDID32)33560755)	/* number: 6323	 type: long */
#define	TA_MSG_CBYTES	((FLDID32)33560756)	/* number: 6324	 type: long */
#define	TA_MSG_CTIME	((FLDID32)33560757)	/* number: 6325	 type: long */
#define	TA_MSG_LRPID	((FLDID32)33560758)	/* number: 6326	 type: long */
#define	TA_MSG_LSPID	((FLDID32)33560759)	/* number: 6327	 type: long */
#define	TA_MSG_QBYTES	((FLDID32)33560760)	/* number: 6328	 type: long */
#define	TA_MSG_QNUM	((FLDID32)33560761)	/* number: 6329	 type: long */
#define	TA_MSG_RTIME	((FLDID32)33560762)	/* number: 6330	 type: long */
#define	TA_MSG_STIME	((FLDID32)33560763)	/* number: 6331	 type: long */
#endif
#define	TA_NADDR	((FLDID32)167778492)	/* number: 6332	 type: string */
#define	TA_NCOMPLETED	((FLDID32)33560765)	/* number: 6333	 type: long */
#define	TA_NLSADDR	((FLDID32)167778494)	/* number: 6334	 type: string */
#define	TA_NOTIFY	((FLDID32)167778495)	/* number: 6335	 type: string */
#define	TA_NQUEUED	((FLDID32)33560768)	/* number: 6336	 type: long */
#define	TA_NUMCONV	((FLDID32)33560769)	/* number: 6337	 type: long */
#define	TA_NUMDEQUEUE	((FLDID32)33560770)	/* number: 6338	 type: long */
#define	TA_NUMENQUEUE	((FLDID32)33560771)	/* number: 6339	 type: long */
#define	TA_NUMPOST	((FLDID32)33560772)	/* number: 6340	 type: long */
#define	TA_NUMREQ	((FLDID32)33560773)	/* number: 6341	 type: long */
/* TA_NUMSEMWKUP was used previously by locking code - do no use offset 142 */
#define	TA_NUMSERVED	((FLDID32)33560775)	/* number: 6343	 type: long */
#define	TA_NUMSUBSCRIBE	((FLDID32)33560776)	/* number: 6344	 type: long */
#define	TA_NUMTRAN	((FLDID32)33560777)	/* number: 6345	 type: long */
#define	TA_NUMTRANABT	((FLDID32)33560778)	/* number: 6346	 type: long */
#define	TA_NUMTRANCMT	((FLDID32)33560779)	/* number: 6347	 type: long */
#define	TA_NUMUNSOL	((FLDID32)33560780)	/* number: 6348	 type: long */
#define	TA_OPENINFO	((FLDID32)167778509)	/* number: 6349	 type: string */
#define	TA_OPTIONS	((FLDID32)167778510)	/* number: 6350	 type: string */
#define	TA_PAGESIZE	((FLDID32)33560783)	/* number: 6351	 type: long */
#define	TA_PID	((FLDID32)33560784)	/* number: 6352	 type: long */
#define	TA_PMID	((FLDID32)167778513)	/* number: 6353	 type: string */
#define	TA_PRIO	((FLDID32)33560786)	/* number: 6354	 type: long */
#define	TA_RANGES	((FLDID32)201332947)	/* number: 6355	 type: carray */
#define	TA_RCMD	((FLDID32)167778516)	/* number: 6356	 type: string */
#define	TA_RCVDBYT	((FLDID32)33560789)	/* number: 6357	 type: long */
#define	TA_RCVDNUM	((FLDID32)33560790)	/* number: 6358	 type: long */
#define	TA_RELEASE	((FLDID32)33560791)	/* number: 6359	 type: long */
#define	TA_REPLYQ	((FLDID32)167778520)	/* number: 6360	 type: string */
#define	TA_RESTART	((FLDID32)167778521)	/* number: 6361	 type: string */
#define	TA_ROLE	((FLDID32)167778522)	/* number: 6362	 type: string */
#define	TA_ROUTINGNAME	((FLDID32)167778523)	/* number: 6363	 type: string */
#define	TA_RPID	((FLDID32)33560796)	/* number: 6364	 type: long */
#define	TA_RPPERM	((FLDID32)33560797)	/* number: 6365	 type: long */
#define	TA_RQADDR	((FLDID32)167778526)	/* number: 6366	 type: string */
#define	TA_RQID	((FLDID32)33560799)	/* number: 6367	 type: long */
#define	TA_RQPERM	((FLDID32)33560800)	/* number: 6368	 type: long */
#define	TA_SANITYSCAN	((FLDID32)33560801)	/* number: 6369	 type: long */
#define	TA_SCANUNIT	((FLDID32)33560802)	/* number: 6370	 type: long */
#define	TA_SECURITY	((FLDID32)167778531)	/* number: 6371	 type: string */
#ifndef REDUCE_CPP_NOIPC
#define	TA_SEMID	((FLDID32)33560804)	/* number: 6372	 type: long */
#define	TA_SEMTICKET	((FLDID32)33560805)	/* number: 6373	 type: long */
#define	TA_SEM_OTIME	((FLDID32)33560806)	/* number: 6374	 type: long */
#endif
#define	TA_SENTBYT	((FLDID32)33560807)	/* number: 6375	 type: long */
#define	TA_SENTNUM	((FLDID32)33560808)	/* number: 6376	 type: long */
#define	TA_SEQUENCE	((FLDID32)33560809)	/* number: 6377	 type: long */
#define	TA_SERVERCNT	((FLDID32)33560810)	/* number: 6378	 type: long */
#define	TA_SERVERNAME	((FLDID32)167778539)	/* number: 6379	 type: string */
#define	TA_SERVICENAME	((FLDID32)167778540)	/* number: 6380	 type: string */
#define	TA_SEVERITY	((FLDID32)167778541)	/* number: 6381	 type: string */
#ifndef REDUCE_CPP_NOIPC
#define	TA_SHMID	((FLDID32)33560814)	/* number: 6382	 type: long */
#define	TA_SHMKEY	((FLDID32)33560815)	/* number: 6383	 type: long */
#define	TA_SHMSZ	((FLDID32)33560816)	/* number: 6384	 type: long */
#define	TA_SHM_ATIME	((FLDID32)33560817)	/* number: 6385	 type: long */
#define	TA_SHM_CTIME	((FLDID32)33560818)	/* number: 6386	 type: long */
#define	TA_SHM_DTIME	((FLDID32)33560819)	/* number: 6387	 type: long */
#define	TA_SHM_NATTCH	((FLDID32)33560820)	/* number: 6388	 type: long */
#endif
#define	TA_SOURCE	((FLDID32)167778549)	/* number: 6389	 type: string */
#define	TA_SPINCOUNT	((FLDID32)33560822)	/* number: 6390	 type: long */
#define	TA_SUSPTIME	((FLDID32)33560823)	/* number: 6391	 type: long */
#define	TA_SVCRNAM	((FLDID32)167778552)	/* number: 6392	 type: string */
#define	TA_SVCTIMEOUT	((FLDID32)33560825)	/* number: 6393	 type: long */
#define	TA_SWRELEASE	((FLDID32)167778554)	/* number: 6394	 type: string */
#define	TA_SYSTEM_ACCESS	((FLDID32)167778555)	/* number: 6395	 type: string */
#define	TA_TIMELEFT	((FLDID32)33560828)	/* number: 6396	 type: long */
#define	TA_TIMEOUT	((FLDID32)33560829)	/* number: 6397	 type: long */
#define	TA_TIMERESTART	((FLDID32)33560830)	/* number: 6398	 type: long */
#define	TA_TIMESTART	((FLDID32)33560831)	/* number: 6399	 type: long */
#define	TA_TLOGCOUNT	((FLDID32)33560832)	/* number: 6400	 type: long */
#define	TA_TLOGDATA	((FLDID32)167778561)	/* number: 6401	 type: string */
#define	TA_TLOGDEVICE	((FLDID32)167778562)	/* number: 6402	 type: string */
#define	TA_TLOGINDEX	((FLDID32)33560835)	/* number: 6403	 type: long */
#define	TA_TLOGNAME	((FLDID32)167778564)	/* number: 6404	 type: string */
#define	TA_TLOGOFFSET	((FLDID32)33560837)	/* number: 6405	 type: long */
#define	TA_TLOGSIZE	((FLDID32)33560838)	/* number: 6406	 type: long */
#define	TA_TMDEBUG	((FLDID32)167778567)	/* number: 6407	 type: string */
#define	TA_TMNETLOAD	((FLDID32)33560840)	/* number: 6408	 type: long */
#define	TA_TMSCOUNT	((FLDID32)33560841)	/* number: 6409	 type: long */
#define	TA_TMSNAME	((FLDID32)167778570)	/* number: 6410	 type: string */
#define	TA_TOTNP	((FLDID32)33560843)	/* number: 6411	 type: long */
#define	TA_TOTNQUEUED	((FLDID32)33560844)	/* number: 6412	 type: long */
#define	TA_TOTNV	((FLDID32)33560845)	/* number: 6413	 type: long */
#define	TA_TOTREQC	((FLDID32)33560846)	/* number: 6414	 type: long */
#define	TA_TOTWANTERS	((FLDID32)33560847)	/* number: 6415	 type: long */
#define	TA_TOTWKQUEUED	((FLDID32)33560848)	/* number: 6416	 type: long */
#define	TA_TOTWKUPRCV	((FLDID32)33560849)	/* number: 6417	 type: long */
#define	TA_TOTWKUPSENT	((FLDID32)33560850)	/* number: 6418	 type: long */
#define	TA_TOTWORKL	((FLDID32)33560851)	/* number: 6419	 type: long */
#define	TA_TPTRANID	((FLDID32)167778580)	/* number: 6420	 type: string */
#define	TA_TRANLEV	((FLDID32)33560853)	/* number: 6421	 type: long */
#define	TA_TRANTIME	((FLDID32)33560854)	/* number: 6422	 type: long */
#define	TA_TUXCONFIG	((FLDID32)167778583)	/* number: 6423	 type: string */
#define	TA_TUXDIR	((FLDID32)167778584)	/* number: 6424	 type: string */
#define	TA_TUXOFFSET	((FLDID32)33560857)	/* number: 6425	 type: long */
#define	TA_TYPE	((FLDID32)167778586)	/* number: 6426	 type: string */
#define	TA_UID	((FLDID32)33560859)	/* number: 6427	 type: long */
#define	TA_ULOGCAT	((FLDID32)167778588)	/* number: 6428	 type: string */
#define	TA_ULOGLINE	((FLDID32)33560861)	/* number: 6429	 type: long */
#define	TA_ULOGMSG	((FLDID32)167778590)	/* number: 6430	 type: string */
#define	TA_ULOGMSGNUM	((FLDID32)33560863)	/* number: 6431	 type: long */
#define	TA_ULOGPFX	((FLDID32)167778592)	/* number: 6432	 type: string */
#define	TA_ULOGPROCNM	((FLDID32)167778593)	/* number: 6433	 type: string */
#define	TA_ULOGTIME	((FLDID32)33560866)	/* number: 6434	 type: long */
#define	TA_USIGNAL	((FLDID32)167778595)	/* number: 6435	 type: string */
#define	TA_USRNAME	((FLDID32)167778596)	/* number: 6436	 type: string */
#define	TA_VALIDATION	((FLDID32)167778597)	/* number: 6437	 type: string */
#define	TA_WAITS	((FLDID32)167778598)	/* number: 6438	 type: string */
#define	TA_WKCOMPLETED	((FLDID32)33560871)	/* number: 6439	 type: long */
#define	TA_WKINITIATED	((FLDID32)33560872)	/* number: 6440	 type: long */
#define	TA_WKQUEUED	((FLDID32)33560873)	/* number: 6441	 type: long */
#define	TA_WSC	((FLDID32)167778602)	/* number: 6442	 type: string */
#define	TA_WSH	((FLDID32)167778603)	/* number: 6443	 type: string */
#define	TA_XID	((FLDID32)167778604)	/* number: 6444	 type: string */
/*
	The field numbers below have been added to the system with release 6.0.
	They begin at 245 (base 6200) and increase from there.
	Field numbers cannot be reused or changed from release to release or
	interoperability will be broken.  Fields 0-244 (base 6200) defined in
	release 5.0, fields defined in release 6.0 begin at field 245 and
	increase from there.

	Note that fields 7200-7299 are reserved for use buy TUXEDO OEMs.
*/
#define	TA_ACLCACHEACCESS	((FLDID32)33560877)	/* number: 6445	 type: long */
#define	TA_ACLCACHEHITS	((FLDID32)33560878)	/* number: 6446	 type: long */
#define	TA_ACLFAIL	((FLDID32)33560879)	/* number: 6447	 type: long */
#define	TA_ACLGROUPIDS	((FLDID32)167778608)	/* number: 6448	 type: string */
#define	TA_ACLNAME	((FLDID32)167778609)	/* number: 6449	 type: string */
#define	TA_ACLTYPE	((FLDID32)167778610)	/* number: 6450	 type: string */
#define	TA_ACTIVE	((FLDID32)167778611)	/* number: 6451	 type: string */
#define	TA_ATTFLAGS	((FLDID32)33560884)	/* number: 6452	 type: long */
#define	TA_CURHANDLERS	((FLDID32)33560885)	/* number: 6453	 type: long */
#define	TA_CURWORK	((FLDID32)33560886)	/* number: 6454	 type: long */
#define	TA_DEFAULT	((FLDID32)167778615)	/* number: 6455	 type: string */
#define	TA_FACTPERM	((FLDID32)33560888)	/* number: 6456	 type: long */
#define	TA_GETSTATES	((FLDID32)167778617)	/* number: 6457	 type: string */
#define	TA_GROUPID	((FLDID32)33560890)	/* number: 6458	 type: long */
#define	TA_GROUPNAME	((FLDID32)167778619)	/* number: 6459	 type: string */
#define	TA_HWACLCACHE	((FLDID32)33560892)	/* number: 6460	 type: long */
#define	TA_HWHANDLERS	((FLDID32)33560893)	/* number: 6461	 type: long */
#define	TA_IDLETIME	((FLDID32)33560894)	/* number: 6462	 type: long */
#define	TA_INASTATES	((FLDID32)167778623)	/* number: 6463	 type: string */
#define	TA_MAXACLCACHE	((FLDID32)33560896)	/* number: 6464	 type: long */
#define	TA_MAXACLGROUPS	((FLDID32)33560897)	/* number: 6465	 type: long */
#define	TA_MAXHANDLERS	((FLDID32)33560898)	/* number: 6466	 type: long */
#define	TA_MAXIDLETIME	((FLDID32)33560899)	/* number: 6467	 type: long */
#define	TA_MAXINITTIME	((FLDID32)33560900)	/* number: 6468	 type: long */
#define	TA_MAXPERM	((FLDID32)33560901)	/* number: 6469	 type: long */
#define	TA_MINHANDLERS	((FLDID32)33560902)	/* number: 6470	 type: long */
#define	TA_MULTIPLEX	((FLDID32)33560903)	/* number: 6471	 type: long */
#define	TA_NUMBLOCKQ	((FLDID32)33560904)	/* number: 6472	 type: long */
#define	TA_PREFERENCES	((FLDID32)167778633)	/* number: 6473	 type: string */
#define	TA_PRINCLTNAME	((FLDID32)167778634)	/* number: 6474	 type: string */
#define	TA_PRINGRP	((FLDID32)33560907)	/* number: 6475	 type: long */
#define	TA_PRINID	((FLDID32)33560908)	/* number: 6476	 type: long */
#define	TA_PRINNAME	((FLDID32)167778637)	/* number: 6477	 type: string */
#define	TA_PRINPASSWD	((FLDID32)167778638)	/* number: 6478	 type: string */
#define	TA_SETSTATES	((FLDID32)167778639)	/* number: 6479	 type: string */
#define	TA_SUSPENDED	((FLDID32)167778640)	/* number: 6480	 type: string */
#define	TA_SVCTYPE	((FLDID32)167778641)	/* number: 6481	 type: string */
#define	TA_TOTACTTIME	((FLDID32)33560914)	/* number: 6482	 type: long */
#define	TA_TOTIDLTIME	((FLDID32)33560915)	/* number: 6483	 type: long */
#define	TA_VIEWREFRESH	((FLDID32)167778644)	/* number: 6484	 type: string */
#define	TA_WSHCLIENTID	((FLDID32)167778645)	/* number: 6485	 type: string */
#define	TA_WSHNAME	((FLDID32)167778646)	/* number: 6486	 type: string */
#define	TA_WSPROTO	((FLDID32)33560919)	/* number: 6487	 type: long */
/*
	The fields below have been added to TMIB beginning release 6.3
*/
#define	TA_COORDGRPNO	((FLDID32)33560920)	/* number: 6488	 type: long */
#define	TA_COORDSRVGRP	((FLDID32)167778649)	/* number: 6489	 type: string */
#define	TA_MINWSHPORT	((FLDID32)33560922)	/* number: 6490	 type: long */
#define	TA_MAXWSHPORT	((FLDID32)33560923)	/* number: 6491	 type: long */
#define	TA_MINENCRYPTBITS	((FLDID32)167778652)	/* number: 6492	 type: string */
#define	TA_MAXENCRYPTBITS	((FLDID32)167778653)	/* number: 6493	 type: string */
#define	TA_CURENCRYPTBITS	((FLDID32)167778654)	/* number: 6494	 type: string */
#define	TA_EXT_NADDR	((FLDID32)167778655)	/* number: 6495	 type: string */
#define	TA_COMPONENTS	((FLDID32)167778656)	/* number: 6496	 type: string */
#define	TA_OLDENCRYPT	((FLDID32)167778657)	/* number: 6497	 type: string */
#define	TA_OLDCMPLIMIT	((FLDID32)167778658)	/* number: 6498	 type: string */
/*
	The fields below have been added to TMIB beginning release 6.4
*/
#define	TA_MAXPENDINGBYTES	((FLDID32)33560931)	/* number: 6499	 type: long */
#define	TA_NETGROUP	((FLDID32)167778660)	/* number: 6500	 type: string */
#define	TA_NETGRPNO	((FLDID32)33560933)	/* number: 6501	 type: long */
#define	TA_NETPRIO	((FLDID32)33560934)	/* number: 6502	 type: long */
#define	TA_MAXNETGROUPS	((FLDID32)33560935)	/* number: 6503	 type: long */
#define	TA_KEEPALIVE	((FLDID32)167778664)	/* number: 6504	 type: string */
#define	TA_NETTIMEOUT	((FLDID32)33560937)	/* number: 6505	 type: long */

#ifndef REDUCE_CPP_NOQMIB
/*
	The fields below belong to the /Q MIB.
*/
#define	TA_APPQMSGID	((FLDID32)167778760)	/* number: 6600	 type: string */
#define	TA_APPQNAME	((FLDID32)167778761)	/* number: 6601	 type: string */
#define	TA_APPQORDER	((FLDID32)167778762)	/* number: 6602	 type: string */
#define	TA_APPQSPACENAME	((FLDID32)167778763)	/* number: 6603	 type: string */
#define	TA_APPQSPACERM	((FLDID32)167778764)	/* number: 6604	 type: string */
#define	TA_BLOCKING	((FLDID32)33561037)	/* number: 6605	 type: long */
#define	TA_CMD	((FLDID32)167778766)	/* number: 6606	 type: string */
#define	TA_CMDHW	((FLDID32)167778767)	/* number: 6607	 type: string */
#define	TA_CMDLW	((FLDID32)167778768)	/* number: 6608	 type: string */
#define	TA_CORRID	((FLDID32)167778769)	/* number: 6609	 type: string */
#define	TA_CURBLOCKS	((FLDID32)33561042)	/* number: 6610	 type: long */
#define	TA_CUREXTENT	((FLDID32)33561043)	/* number: 6611	 type: long */
#define	TA_CURMSG	((FLDID32)33561044)	/* number: 6612	 type: long */
#define	TA_CURPROC	((FLDID32)33561045)	/* number: 6613	 type: long */
#define	TA_CURRETRIES	((FLDID32)33561046)	/* number: 6614	 type: long */
#define	TA_CURTRANS	((FLDID32)33561047)	/* number: 6615	 type: long */
#define	TA_ERRORQNAME	((FLDID32)167778776)	/* number: 6616	 type: string */
#define	TA_FORCEINIT	((FLDID32)167778777)	/* number: 6617	 type: string */
#define	TA_HIGHPRIORITY	((FLDID32)33561050)	/* number: 6618	 type: long */
#define	TA_HWMSG	((FLDID32)33561051)	/* number: 6619	 type: long */
#define	TA_HWPROC	((FLDID32)33561052)	/* number: 6620	 type: long */
#define	TA_HWTRANS	((FLDID32)33561053)	/* number: 6621	 type: long */
#define	TA_LOWPRIORITY	((FLDID32)33561054)	/* number: 6622	 type: long */
#define	TA_LSTATE	((FLDID32)167778783)	/* number: 6623	 type: string */
#define	TA_MAXMSG	((FLDID32)33561056)	/* number: 6624	 type: long */
#define	TA_MAXPAGES	((FLDID32)33561057)	/* number: 6625	 type: long */
#define	TA_MAXPROC	((FLDID32)33561058)	/* number: 6626	 type: long */
#define	TA_MAXRETRIES	((FLDID32)33561059)	/* number: 6627	 type: long */
#define	TA_MAXTRANS	((FLDID32)33561060)	/* number: 6628	 type: long */
#define	TA_MSGENDTIME	((FLDID32)167778789)	/* number: 6629	 type: string */
#define	TA_MSGSIZE	((FLDID32)33561062)	/* number: 6630	 type: long */
#define	TA_MSGSTARTTIME	((FLDID32)167778791)	/* number: 6631	 type: string */
#define	TA_NEWAPPQNAME	((FLDID32)167778792)	/* number: 6632	 type: string */
#define	TA_OUTOFORDER	((FLDID32)167778793)	/* number: 6633	 type: string */
#define	TA_PERCENTINIT	((FLDID32)33561066)	/* number: 6634	 type: long */
#define	TA_PRIORITY	((FLDID32)33561067)	/* number: 6635	 type: long */
#define	TA_QMCONFIG	((FLDID32)167778796)	/* number: 6636	 type: string */
#define	TA_RETRYDELAY	((FLDID32)33561069)	/* number: 6637	 type: long */
#define	TA_STARTTIME	((FLDID32)167778798)	/* number: 6638	 type: string */
#define	TA_TIME	((FLDID32)167778799)	/* number: 6639	 type: string */
#endif
#endif
/*
	The field numbers below should always begin at 500 and increase.
	Field numbers cannot be reused or changed from release to release or
	interoperability will be broken.

	Note that fields 7200-7299 are reserved for use buy TUXEDO OEMs.
*/
#define	TA_LMID	((FLDID32)167778860)	/* number: 6700	 type: string */
#define	TA_PASSWORD	((FLDID32)167778861)	/* number: 6701	 type: string */
#define	TA_SRVGRP	((FLDID32)167778862)	/* number: 6702	 type: string */
#define	TA_SRVID	((FLDID32)33561135)	/* number: 6703	 type: long */
#define	TA_TMTRACE	((FLDID32)167778864)	/* number: 6704	 type: string */


/* 
	The following fields have been defined for use by the Domain MIB.
	The field numbers reserved for Domain MIB are 6750 - 6950.
*/

#ifndef REDUCE_CPP

 /* DM_SNADOM    */
#define	TA_DMSNADOM	((FLDID32)167778920)	/* number: 6760	 type: string */
#define	TA_DMLUNAME	((FLDID32)167778921)	/* number: 6761	 type: string */
#define	TA_DMMODENAME	((FLDID32)167778922)	/* number: 6762	 type: string */
#define	TA_DMNETID	((FLDID32)167778923)	/* number: 6763	 type: string */
#define	TA_DMSECTYPE	((FLDID32)167778924)	/* number: 6764	 type: string */
#define	TA_DMSYMDESTNAME	((FLDID32)167778925)	/* number: 6765	 type: string */
#define	TA_DMMAXSNASESS	((FLDID32)6766)	/* number: 6766	 type: short */
#define	TA_DMLCONV	((FLDID32)167778927)	/* number: 6767	 type: string */

 /* DM_TDOMAIN */
#define	TA_DMTDOM	((FLDID32)167778935)	/* number: 6775	 type: string */
#define	TA_DMNWADDR	((FLDID32)167778936)	/* number: 6776	 type: string */
#define	TA_NWADDRLEN	((FLDID32)6777)	/* number: 6777	 type: short */
#define	TA_DMNWDEVICE	((FLDID32)167778938)	/* number: 6778	 type: string */
#define	TA_DMNWIDLETIME	((FLDID32)33561211)	/* number: 6779	 type: long */

 /* DM_OSITP */
#define	TA_DMOSITP	((FLDID32)167778940)	/* number: 6780	 type: string */
#define	TA_DMAPT	((FLDID32)167778941)	/* number: 6781	 type: string */
#define	TA_APTLEN	((FLDID32)6782)	/* number: 6782	 type: short */
#define	TA_DMAEQ	((FLDID32)167778943)	/* number: 6783	 type: string */
#define	TA_AEQLEN	((FLDID32)6784)	/* number: 6784	 type: short */
#define	TA_DMAET	((FLDID32)167778945)	/* number: 6785	 type: string */
#define	TA_DMACN	((FLDID32)167778946)	/* number: 6786	 type: string */
#define	TA_ACN2	((FLDID32)6787)	/* number: 6787	 type: short */
#define	TA_DMAPID	((FLDID32)6788)	/* number: 6788	 type: short */
#define	TA_DMAEID	((FLDID32)6789)	/* number: 6789	 type: short */
#define	TA_DMPROFILE	((FLDID32)167778950)	/* number: 6790	 type: string */
#define	TA_PROFILE2	((FLDID32)6791)	/* number: 6791	 type: short */
#define	TA_DMURCH	((FLDID32)167778952)	/* number: 6792	 type: string */

 /* DM_LOCAL_DOMAIN */
#define	TA_LDOMAIN	((FLDID32)167778953)	/* number: 6793	 type: string */
#define	TA_DMSRVGROUP	((FLDID32)167778954)	/* number: 6794	 type: string */
#define	TA_DMTYPE	((FLDID32)167778955)	/* number: 6795	 type: string */
#define	TA_DOMAINAME	((FLDID32)167778956)	/* number: 6796	 type: string */
#define	TA_DMTLOGDEV	((FLDID32)167778957)	/* number: 6797	 type: string */
#define	TA_DMAUDITLOG	((FLDID32)167778958)	/* number: 6798	 type: string */
#define	TA_DMBLOCKTIME	((FLDID32)6799)	/* number: 6799	 type: short */
#define	TA_DMTLOGNAME	((FLDID32)167778960)	/* number: 6800	 type: string */
#define	TA_DMTLOGSIZE	((FLDID32)33561233)	/* number: 6801	 type: long */
#define	TA_MAXDATALEN	((FLDID32)33561234)	/* number: 6802	 type: long */
#define	TA_DMMAXRDOM	((FLDID32)33561235)	/* number: 6803	 type: long */
#define	TA_DMMAXRDTRAN	((FLDID32)6804)	/* number: 6804	 type: short */
#define	TA_DMMAXTRAN	((FLDID32)6805)	/* number: 6805	 type: short */
#define	TA_MAXSENDLEN	((FLDID32)33561238)	/* number: 6806	 type: long */
#define	TA_DMSECURITY	((FLDID32)167778967)	/* number: 6807	 type: string */
#define	TA_SECURITY2	((FLDID32)6808)	/* number: 6808	 type: short */
#define	TA_DMTUXCONFIG	((FLDID32)167778969)	/* number: 6809	 type: string */
#define	TA_DMTUXOFFSET	((FLDID32)33561242)	/* number: 6810	 type: long */
#define	TA_BLOB_SHM_SIZE	((FLDID32)33561243)	/* number: 6811	 type: long */

 /* DM_REMOTE_DOMAIN */
#define	TA_RDOMAIN	((FLDID32)167778972)	/* number: 6812	 type: string */
#define	TA_LOOPBACK	((FLDID32)6813)	/* number: 6813	 type: short */

 /* DM_ACCESS_CONTROL */
#define	TA_DMACLNAME	((FLDID32)167778974)	/* number: 6814	 type: string */
#define	TA_NRDOM	((FLDID32)6815)	/* number: 6815	 type: short */
#define	TA_DMRDOMLIST	((FLDID32)167778976)	/* number: 6816	 type: string */

 /* DM_LOCAL_SERVICES */
#define	TA_DMBUFTYPE	((FLDID32)167778977)	/* number: 6817	 type: string */
#define	TA_DMBUFSTYPE	((FLDID32)167778978)	/* number: 6818	 type: string */
#define	TA_DMINBUFTYPE	((FLDID32)167778979)	/* number: 6819	 type: string */
#define	TA_DMOBUFTYPE	((FLDID32)167778980)	/* number: 6820	 type: string */
#define	TA_DMOBUFSTYPE	((FLDID32)167778981)	/* number: 6821	 type: string */
#define	TA_DMOUTBUFTYPE	((FLDID32)167778982)	/* number: 6822	 type: string */
#define	TA_DMREMOTENAME	((FLDID32)167778983)	/* number: 6823	 type: string */
#define	TA_DMSERVICENAME	((FLDID32)167778984)	/* number: 6824	 type: string */

 /* DM_ROUTING */
#define	TA_DMROUTINGNAME	((FLDID32)167778985)	/* number: 6825	 type: string */
#define	TA_DMFIELD	((FLDID32)167778986)	/* number: 6826	 type: string */
#define	TA_DMRANGES	((FLDID32)167778987)	/* number: 6827	 type: string */

 /* DM_REMOTE_SERVICES */
#define	TA_DMAUTOTRAN	((FLDID32)167778988)	/* number: 6828	 type: string */
#define	TA_AUTOTRAN2	((FLDID32)6829)	/* number: 6829	 type: short */
#define	TA_DMCONV	((FLDID32)167778990)	/* number: 6830	 type: string */
#define	TA_CONV2	((FLDID32)6831)	/* number: 6831	 type: short */
#define	TA_DMLOAD	((FLDID32)6832)	/* number: 6832	 type: short */
#define	TA_DMPRIO	((FLDID32)6833)	/* number: 6833	 type: short */
#define	TA_DMTRANTIME	((FLDID32)33561266)	/* number: 6834	 type: long */
#define	TA_DUMMYSRVGRP	((FLDID32)167778995)	/* number: 6835	 type: string */
#define	TA_DUMMYSRVID	((FLDID32)6836)	/* number: 6836	 type: short */
#define	TA_DUMMYSTATE	((FLDID32)6837)	/* number: 6837	 type: short */
#define	TA_DUMMYACTIVITY	((FLDID32)6838)	/* number: 6838	 type: short */
#define	TA_DUMMYALLSTATS	((FLDID32)6839)	/* number: 6839	 type: short */
#define	TA_DUMMYDOMSTATS	((FLDID32)6840)	/* number: 6840	 type: short */
#define	TA_DUMMYOPTION	((FLDID32)6841)	/* number: 6841	 type: short */
#define	TA_ON	((FLDID32)6842)	/* number: 6842	 type: short */
#define	TA_OFF	((FLDID32)6843)	/* number: 6843	 type: short */
#define	TA_RESET	((FLDID32)6844)	/* number: 6844	 type: short */
#define	TA_DUMMYTOGGLE	((FLDID32)6845)	/* number: 6845	 type: short */
#define	TA_DMSTATISTICS	((FLDID32)6846)	/* number: 6846	 type: short */
#define	TA_DMAUDIT	((FLDID32)6847)	/* number: 6847	 type: short */
#define	TA_TRAN_STATE	((FLDID32)33561280)	/* number: 6848	 type: long */
/*TA_DUMMYMTYPE	100	string	-	Machine Type in LDOM/RDOM */
#define	TA_DMNUMREQLSVC	((FLDID32)33561282)	/* number: 6850	 type: long */
#define	TA_DMNUMREQRSVC	((FLDID32)33561283)	/* number: 6851	 type: long */
#define	TA_DMNUMREPLSVC	((FLDID32)33561284)	/* number: 6852	 type: long */
#define	TA_DMNUMREPRSVC	((FLDID32)33561285)	/* number: 6853	 type: long */
#define	TA_CURACTRQ	((FLDID32)33561286)	/* number: 6854	 type: long */
#define	TA_DMNUMREQCOMP	((FLDID32)33561287)	/* number: 6855	 type: long */
#define	TA_DMNUMREQFAIL	((FLDID32)33561288)	/* number: 6856	 type: long */
#define	TA_DMNUMCONVACT	((FLDID32)33561289)	/* number: 6857	 type: long */
#define	TA_DMNUMCONVLOC	((FLDID32)33561290)	/* number: 6858	 type: long */
#define	TA_DMNUMCONVLSND	((FLDID32)33561291)	/* number: 6859	 type: long */
#define	TA_DMNUMCONVREM	((FLDID32)33561292)	/* number: 6860	 type: long */
#define	TA_DMNUMCONVRSND	((FLDID32)33561293)	/* number: 6861	 type: long */
#define	TA_CURACTEV	((FLDID32)33561294)	/* number: 6862	 type: long */
#define	TA_CURSUSPTEV	((FLDID32)33561295)	/* number: 6863	 type: long */
#define	TA_CURSUSPNEV	((FLDID32)33561296)	/* number: 6864	 type: long */
#define	TA_DMNUMTXBEGUN	((FLDID32)33561297)	/* number: 6865	 type: long */
#define	TA_DMNUMTXCOMMIT	((FLDID32)33561298)	/* number: 6866	 type: long */
#define	TA_DMNUMTXHCOMMIT	((FLDID32)33561299)	/* number: 6867	 type: long */
#define	TA_DMNUMTXHRLBCK	((FLDID32)33561300)	/* number: 6868	 type: long */
#define	TA_DMNUMTXRLBCK	((FLDID32)33561301)	/* number: 6869	 type: long */
#define	TA_DMSTATRESETIME	((FLDID32)33561302)	/* number: 6870	 type: long */
#define	TA_SHM_CONTENTIONS	((FLDID32)33561303)	/* number: 6871	 type: long */
#define	TA_LOG_CONTENTIONS	((FLDID32)33561304)	/* number: 6872	 type: long */
#define	TA_PASSWD	((FLDID32)167779035)	/* number: 6875	 type: string */
#define	TA_ENCPASSWD	((FLDID32)167779036)	/* number: 6876	 type: string */
#define	TA_DMLPWD	((FLDID32)167779037)	/* number: 6877	 type: string */
#define	TA_DMRPWD	((FLDID32)167779038)	/* number: 6878	 type: string */
#define	TA_ENC2_LPWD	((FLDID32)167779039)	/* number: 6879	 type: string */
#define	TA_ENC2_RPWD	((FLDID32)167779040)	/* number: 6880	 type: string */
#define	TA_REENCRYPT_PWD	((FLDID32)6881)	/* number: 6881	 type: short */
#define	TA_DMGWNUM	((FLDID32)33561314)	/* number: 6882	 type: long */
#define	TA_DMRDOMNUM	((FLDID32)33561315)	/* number: 6883	 type: long */
#define	TA_DMTXPARENT	((FLDID32)167779044)	/* number: 6884	 type: string */
#define	TA_DMTXID	((FLDID32)167779045)	/* number: 6885	 type: string */
#define	TA_DMPRINNAME	((FLDID32)167779053)	/* number: 6893	 type: string */
#define	TA_DMRPRINNAME	((FLDID32)167779054)	/* number: 6894	 type: string */
#define	TA_DMRPRINPASSWD	((FLDID32)167779055)	/* number: 6895	 type: string */
#define	TA_DMDIRECTION	((FLDID32)167779056)	/* number: 6896	 type: string */
#define	TA_DMRDOMSEC	((FLDID32)167779057)	/* number: 6897	 type: string */
#define	TA_DMRDOMUSR	((FLDID32)167779058)	/* number: 6898	 type: string */
#define	TA_DMCMPLIMIT	((FLDID32)33561332)	/* number: 6900	 type: long */
#define	TA_DMMINENCRYPTBITS	((FLDID32)167779061)	/* number: 6901	 type: string */
#define	TA_DMMAXENCRYPTBITS	((FLDID32)167779062)	/* number: 6902	 type: string */
/* SNA2 Begin */
#define	TA_DMSNALINK	((FLDID32)167779063)	/* number: 6903	 type: string */
#define	TA_DMLSYSID	((FLDID32)167779064)	/* number: 6904	 type: string */
#define	TA_DMRSYSID	((FLDID32)167779065)	/* number: 6905	 type: string */
#define	TA_DMMINWIN	((FLDID32)6906)	/* number: 6906	 type: short */
#define	TA_DMMAXSYNCLVL	((FLDID32)6907)	/* number: 6907	 type: short */
#define	TA_DMSNASTACK	((FLDID32)167779068)	/* number: 6908	 type: string */
#define	TA_DMSNACRM	((FLDID32)167779069)	/* number: 6909	 type: string */
#define	TA_DMSTACKTYPE	((FLDID32)167779070)	/* number: 6910	 type: string */
#define	TA_DMTPNAME	((FLDID32)167779071)	/* number: 6911	 type: string */
#define	TA_DMSTACKPARMS	((FLDID32)167779072)	/* number: 6912	 type: string */
#define	TA_DMAPI	((FLDID32)167779073)	/* number: 6913	 type: string */
#define	TA_DMFUNCTION	((FLDID32)167779074)	/* number: 6914	 type: string */
#define	TA_DMCODEPAGE	((FLDID32)167779075)	/* number: 6915	 type: string */
#define	TA_DMSTARTTYPE	((FLDID32)167779076)	/* number: 6916	 type: string */
/* SNA2 End */
#define	TA_DMCONNECTION_POLICY	((FLDID32)167779077)	/* number: 6917	 type: string */
#define	TA_DMRETRY_INTERVAL	((FLDID32)33561350)	/* number: 6918	 type: long */
#define	TA_DMMAXRETRY	((FLDID32)33561351)	/* number: 6919	 type: long */
/* end REDUCE_CPP */
#endif
/* 
	The following fields have been defined for use by the Event MIB.
	The field numbers reserved for Event MIB are 6950 - 7000.
*/

#ifndef REDUCE_CPP
#define	TA_EVENT_EXPR	((FLDID32)167779111)	/* number: 6951	 type: string */
#define	TA_EVENT_FILTER	((FLDID32)167779112)	/* number: 6952	 type: string */
#define	TA_EVENT_FILTER_BINARY	((FLDID32)201333545)	/* number: 6953	 type: carray */
#define	TA_QSPACE	((FLDID32)167779114)	/* number: 6954	 type: string */
#define	TA_QNAME	((FLDID32)167779115)	/* number: 6955	 type: string */
#define	TA_QCTL_QTIME_ABS	((FLDID32)6956)	/* number: 6956	 type: short */
#define	TA_QCTL_QTIME_REL	((FLDID32)6957)	/* number: 6957	 type: short */
#define	TA_QCTL_QTOP	((FLDID32)6958)	/* number: 6958	 type: short */
#define	TA_QCTL_BEFOREMSGID	((FLDID32)6959)	/* number: 6959	 type: short */
#define	TA_QCTL_DEQ_TIME	((FLDID32)33561392)	/* number: 6960	 type: long */
#define	TA_QCTL_PRIORITY	((FLDID32)33561393)	/* number: 6961	 type: long */
#define	TA_QCTL_MSGID	((FLDID32)167779122)	/* number: 6962	 type: string */
#define	TA_QCTL_CORRID	((FLDID32)167779123)	/* number: 6963	 type: string */
#define	TA_QCTL_REPLYQUEUE	((FLDID32)167779124)	/* number: 6964	 type: string */
#define	TA_QCTL_FAILUREQUEUE	((FLDID32)167779125)	/* number: 6965	 type: string */
#define	TA_EVENT_PERSIST	((FLDID32)6966)	/* number: 6966	 type: short */
#define	TA_EVENT_TRAN	((FLDID32)6967)	/* number: 6967	 type: short */
#define	TA_USERLOG	((FLDID32)167779128)	/* number: 6968	 type: string */
#define	TA_COMMAND	((FLDID32)167779129)	/* number: 6969	 type: string */
#define	TA_EVENT_SERVER	((FLDID32)167779130)	/* number: 6970	 type: string */
#endif
#endif
