/*	Copyright (c) 1998 BEA Systems, Inc.
	All rights reserved

	THIS IS UNPUBLISHED PROPRIETARY
	SOURCE CODE OF BEA Systems, Inc.
	The copyright notice above does not
	evidence any actual or intended
	publication of such source code.
*/

/*	Copyright 1996 BEA Systems, Inc.	*/
/*	THIS IS UNPUBLISHED PROPRIETARY SOURCE CODE OF     	*/
/*	BEA Systems, Inc.                     	*/
/*	The copyright notice above does not evidence any   	*/
/*	actual or intended publication of such source code.	*/

/*	Copyright 1996 BEA Systems, Inc.	*/
/*	THIS IS UNPUBLISHED PROPRIETARY SOURCE CODE OF     	*/
/*	BEA Systems, Inc.                     	*/
/*	The copyright notice above does not evidence any   	*/
/*	actual or intended publication of such source code.	*/

/*
 * Copyright * 1996 BEA Systems, Inc.		
 *		
 * Portions of this software Copyright * 1995 Novell, Inc.
 * All rights reserved
 *
 * THIS IS UNPUBLISHED PROPRIETARY
 * SOURCE CODE OF BEA Systems, Inc.
 * The copyright notice above does not
 * evidence any actual or intended
 * publication of such source code.
 *
 * #ident	"@(#) tuxedo/include/atmi.h	$Revision: 1.1.6.1 $"
 */

#ifndef ATMI_H
#define ATMI_H

#ifndef TMENV_H
#include <tmenv.h>
#endif
#ifndef NOWHAT
static	char	h_atmi[] = "@(#) tuxedo/include/atmi.h	$Revision: 1.1.6.1 $";
#endif

/*
 *	DEFINITIONS NEEDED BY USER APPLICATION PROGRAMS.
 *
 *	Warning: This header file should not be changed in any
 *	way, doing so will destroy the compatibility with TUXEDO programs
 *	and libraries.
 */

/* Flags to service routines */

#define TPNOBLOCK	0x00000001	/* non-blocking send/rcv */
#define TPSIGRSTRT	0x00000002	/* restart rcv on interrupt */
#define TPNOREPLY	0x00000004	/* no reply expected */
#define TPNOTRAN	0x00000008	/* not sent in transaction mode */
#define TPTRAN		0x00000010	/* sent in transaction mode */
#define TPNOTIME	0x00000020	/* no timeout */
#define TPABSOLUTE	0x00000040	/* absolute value on tmsetprio */
#define TPGETANY	0x00000080	/* get any valid reply */
#define TPNOCHANGE	0x00000100	/* force incoming buffer to match */
#define RESERVED_BIT1	0x00000200	/* reserved for future use */
#define TPCONV		0x00000400	/* conversational service */
#define TPSENDONLY	0x00000800	/* send-only mode */
#define TPRECVONLY	0x00001000	/* recv-only mode */
#define TPACK		0x00002000	/* */

/* Flags to tpreturn() */
#define TPFAIL		0x00000001	/* service FAILure for tpreturn */
#define TPSUCCESS	0x00000002	/* service SUCCESS for tpreturn */
#define TPEXIT		0x08000000	/* service failue with server exit */

/* Flags to tpscmt() - Valid TP_COMMIT_CONTROL characteristic values */
#define TP_CMT_LOGGED	0x01		/* return after commit decision is logged */
#define TP_CMT_COMPLETE	0x02		/* return after commit has completed */

/* Flags to tpinit() */
#define TPU_MASK	0x00000007	/* unsolicited notification mask */
#define TPU_SIG		0x00000001	/* signal based notification */
#define TPU_DIP		0x00000002	/* dip-in based notification */
#define TPU_IGN		0x00000004	/* ignore unsolicited messages */

#define TPSA_FASTPATH	0x00000008	/* System access == fastpath */
#define TPSA_PROTECTED	0x00000010	/* System access == protected */

#define TPMULTICONTEXTS	0x00000020	/* Enable MULTI context */

/* Flags to tpconvert() */
#define TPTOSTRING	0x40000000	/* Convert structure to string */
#define TPCONVCLTID	0x00000001	/* Convert CLIENTID */
#define TPCONVTRANID	0x00000002	/* Convert TRANID */
#define TPCONVXID	0x00000004	/* Convert XID */

#define TPCONVMAXSTR	256		/* Maximum string size */

/* Return values to tpchkauth() */
#define TPNOAUTH	0		/* no authentication */
#define TPSYSAUTH	1		/* system authentication */
#define TPAPPAUTH	2		/* system and application authentication */

#ifndef MAXTIDENT
#define MAXTIDENT	30		/* max len of a /T identifier */
#endif

/* client identifier structure */
struct clientid_t {
	long	clientdata[4];		/* reserved for internal use */
};
typedef struct clientid_t CLIENTID;

/* interface to service routines */
struct tpsvcinfo {
#define XATMI_SERVICE_NAME_LENGTH  32
	char	name[XATMI_SERVICE_NAME_LENGTH];/* service name invoked */
	long	flags;		/* describes service attributes */
	char	*data;		/* pointer to data */
	long	len;		/* request data length */
	int	cd;		/* reserved for future use */
	long	appkey;		/* application authentication client key */
	CLIENTID cltid;		/* client identifier for originating client */
};
typedef struct tpsvcinfo TPSVCINFO;

/* X/Open buffer types */
#define X_OCTET "X_OCTET"
#define X_C_TYPE "X_C_TYPE"
#define X_COMMON "X_COMMON"

/* tpinit(3) interface structure */
struct	tpinfo_t {
	char	usrname[MAXTIDENT+2];	/* client user name */
	char	cltname[MAXTIDENT+2];	/* application client name */
	char	passwd[MAXTIDENT+2];	/* application password */
	char	grpname[MAXTIDENT+2];	/* client group name */
	long	flags;			/* initialization flags */
	long	datalen;		/* length of app specific data */
	long	data;			/* placeholder for app data */
};
typedef	struct	tpinfo_t TPINIT;
#ifndef lint
#define TPINITNEED(u)	(((u) > sizeof(long)) \
				? (sizeof(TPINIT) - sizeof(long) + (u)) \
				: (sizeof(TPINIT)))
#else
extern long	TPINITNEED _((long));
#endif

/* TPTRANID structure for tpsuspend(3) and tpresume(3) */
struct tp_tranid_t {
	long info[6];
};

typedef struct tp_tranid_t TPTRANID;

#if defined(__cplusplus)
extern "C" {
#endif

#if (defined(_TM_WIN) || defined(_TM_OS2)) && !defined(_TMDLL)

extern int _TM_FAR * _TMDLLENTRY _tmget_tperrno_addr(void);
extern long _TM_FAR * _TMDLLENTRY _tmget_tpurcode_addr(void);
extern int _TMDLLENTRY gettperrno(void);
extern long _TMDLLENTRY gettpurcode(void);
#define tperrno	(*_tmget_tperrno_addr())
#define tpurcode (*_tmget_tpurcode_addr())

#else

#ifdef _TMSTHREADS
extern int _TM_FAR * _TMDLLENTRY _tmget_tperrno_addr(void);
extern long _TM_FAR * _TMDLLENTRY _tmget_tpurcode_addr(void);
#define tperrno	(*_tmget_tperrno_addr())
#define tpurcode (*_tmget_tpurcode_addr())
#else
_TMITUXWSC extern _TM_THREADVAR int tperrno;
_TMITUXWSC extern _TM_THREADVAR long tpurcode;
#endif

#endif

#if defined(__cplusplus)
}
#endif

/*
 * tperrno values - error codes
 * The man pages explain the context in which the following error codes
 * can return.
 */

#define TPMINVAL	0	/* minimum error message */
#define TPEABORT	1
#define TPEBADDESC	2
#define TPEBLOCK	3
#define TPEINVAL	4
#define TPELIMIT	5
#define TPENOENT	6
#define TPEOS		7
#define TPEPERM		8
#define TPEPROTO	9
#define TPESVCERR	10
#define TPESVCFAIL	11
#define TPESYSTEM	12
#define TPETIME		13
#define TPETRAN		14
#define TPGOTSIG	15
#define TPERMERR	16
#define TPEITYPE	17
#define TPEOTYPE	18
#define TPERELEASE	19
#define TPEHAZARD	20
#define TPEHEURISTIC	21
#define TPEEVENT	22
#define	TPEMATCH	23
#define TPEDIAGNOSTIC	24
#define TPEMIB		25
#define TPMAXVAL	26	/* maximum error message */
/*
 *  WARNING:  when adding new error messages above, remember to:
 *	- increase TPMAXVAL
 *	- add a string for the message to LIBTUX.text
 *	- add an array entry in _tmemsgs[]
 */

/*
 * tperrordetail values - error detail codes
 * The man pages explain the context in which the following error detail codes
 * can return.
 */

#define TPED_MINVAL		0	/* minimum error message */
#define TPED_SVCTIMEOUT		1
#define TPED_TERM		2
#define TPED_NOUNSOLHANDLER	3
#define TPED_NOCLIENT		4
#define TPED_DOMAINUNREACHABLE	5
#define TPED_CLIENTDISCONNECTED 6
#define TPED_MAXVAL		7	/* maximum error message */
/*
 *  WARNING:  when adding new error messages above, remember to:
 *	- increase TPED_MAXVAL
 *	- add a string for the message to LIBTUX.text
 *	- add a string for the message to LIBWSC.text
 *	- add an array entry in _tmedmsgs[]
 */

#ifdef _as400_
extern void _tmunsolerrhdlr _((char *, long, long));
#define TPUNSOLERR	_tmunsolerrhdlr
#else
#define TPUNSOLERR	((void (_TMDLLENTRY *) _((char _TM_FAR *, long, long))) -1)
#endif

/* conversations - events */
#define TPEV_DISCONIMM	0x0001
#define TPEV_SVCERR	0x0002
#define TPEV_SVCFAIL	0x0004
#define TPEV_SVCSUCC	0x0008
#define TPEV_SENDONLY	0x0020


#if defined(__cplusplus)
extern "C" {
#endif

extern char	_TM_FAR * _TMDLLENTRY tpalloc _((char _TM_FAR *, char _TM_FAR *, long));
extern char	_TM_FAR * _TMDLLENTRY tprealloc _((char _TM_FAR *, long));
extern int	_TMDLLENTRY tpcall _((char _TM_FAR *, char _TM_FAR *, long, char _TM_FAR * _TM_FAR *, long _TM_FAR *, long));
extern int	_TMDLLENTRY tpacall _((char _TM_FAR *, char _TM_FAR *, long, long));
extern int	_TMDLLENTRY tpgetrply _((int _TM_FAR *, char _TM_FAR * _TM_FAR *, long _TM_FAR *, long));
extern int	_TMDLLENTRY tpcancel _((int));
extern int	_TMDLLENTRY tpscmt _((long));
extern int	_TMDLLENTRY tpabort _((long));
extern int	_TMDLLENTRY tpbegin _((unsigned long, long));
extern int	_TMDLLENTRY tpcommit _((long));
extern int	_TMDLLENTRY tpconvert _((char _TM_FAR *, char _TM_FAR *, long));
extern int	_TMDLLENTRY tpsuspend _((TPTRANID _TM_FAR *, long));
extern int	_TMDLLENTRY tpresume _((TPTRANID _TM_FAR *, long));
extern int	tpsvrinit _((int, char **));
extern int	_TMDLLENTRY tpinit _((TPINIT _TM_FAR *));
extern int	_TMDLLENTRY tpterm _((void));
extern int	_TMDLLENTRY tpsprio _((int, long));
extern int	_TMDLLENTRY tpgprio _((void));
extern int	_TMDLLENTRY tpopen _((void));
extern int	_TMDLLENTRY tpclose _((void));
extern int	_TMDLLENTRY tpgetlev _((void));
extern long	_TMDLLENTRY tptypes _((char _TM_FAR *, char _TM_FAR *, char _TM_FAR *));
extern void	_TMDLLENTRY tpfree _((char _TM_FAR *));
extern void	_TMDLLENTRY tpforward _((char *, char *, long, long));
extern void	_TMDLLENTRY tpreturn _((int, long, char *, long, long));
extern void	tpsvrdone _((void));
extern int	_TMDLLENTRY tpchkauth _((void));
extern int	_TMDLLENTRY tpbroadcast _((char _TM_FAR *, char _TM_FAR *, char _TM_FAR *, char _TM_FAR *, long, long));
extern int	_TMDLLENTRY tpnotify _((CLIENTID _TM_FAR *, char _TM_FAR *, long, long));
extern void	(_TMDLLENTRY * _TMDLLENTRY tpsetunsol _((void (_TMDLLENTRY *)(char _TM_FAR *, long, long)))) _((char _TM_FAR *, long, long));
extern int	_TMDLLENTRY tpchkunsol _((void));
extern int	_TMDLLENTRY tpadvertise _((char *, void (*)(TPSVCINFO *)));
extern int	_TMDLLENTRY tpunadvertise _((char *));
extern char 	_TM_FAR * _TMDLLENTRY tpstrerror _((int));
extern long 	_TMDLLENTRY tperrordetail _((long));
extern char 	_TM_FAR * _TMDLLENTRY tpstrerrordetail _((long, long));


/* conversations */
extern int	_TMDLLENTRY tpsend _((int, char _TM_FAR *, long, long, long _TM_FAR *));
extern int	_TMDLLENTRY tprecv _((int, char _TM_FAR * _TM_FAR *, long _TM_FAR *, long, long _TM_FAR *));
extern int	_TMDLLENTRY tpconnect _((char _TM_FAR *, char _TM_FAR *, long, long));
extern int	_TMDLLENTRY tpdiscon _((int));

/* /T Addition */
extern int	_TMDLLENTRY bq _((char _TM_FAR *));

/* /WS additions */

#if defined(_TM_WIN) || defined(_TM_OS2) || defined(WIN32)
typedef int (_TMDLLENTRY * _TM_FARPROC)(void);
extern int _TMDLLENTRY AEWisblocked _((void));
_TM_FARPROC _TMDLLENTRY AEWsetblockinghook _((_TM_FARPROC));
extern int _TMDLLENTRY AEPisblocked _((void));
_TM_FARPROC _TMDLLENTRY AEPsetblockinghook _((_TM_FARPROC));
extern int _TMDLLENTRY AEWsetunsol _((unsigned int, unsigned int));

#endif

extern char _TM_FAR * _TMDLLENTRY tuxgetenv _((char _TM_FAR *));
extern int _TMDLLENTRY tuxputenv _((char _TM_FAR *));
extern int _TMDLLENTRY tuxreadenv _((char _TM_FAR *, char _TM_FAR *));

#if defined(__cplusplus)
}
#endif


#ifndef _QADDON
#define _QADDON

/* START QUEUED MESSAGES ADD-ON */

#define TMQNAMELEN	15
#define TMMSGIDLEN	32
#define TMCORRIDLEN	32

struct tpqctl_t {		/* control parameters to queue primitives */
	long flags;		/* indicates which of the values are set */
	long deq_time;		/* absolute/relative  time for dequeuing */
	long priority;		/* enqueue priority */
	long diagnostic;	/* indicates reason for failure */
	char msgid[TMMSGIDLEN];	/* id of message before which to queue */
	char corrid[TMCORRIDLEN];/* correlation id used to identify message */
	char replyqueue[TMQNAMELEN+1];	/* queue name for reply message */
	char failurequeue[TMQNAMELEN+1];/* queue name for failure message */
	CLIENTID cltid;		/* client identifier for originating client */
	long urcode;		/* application user-return code */
	long appkey;		/* application authentication client key */
};
typedef struct tpqctl_t TPQCTL;

/* structure elements that are valid - set in flags */
#ifndef TPNOFLAGS
#define TPNOFLAGS	0x00000
#endif
#define	TPQCORRID	0x00001		/* set/get correlation id */
#define	TPQFAILUREQ	0x00002		/* set/get failure queue */
#define	TPQBEFOREMSGID	0x00004		/* enqueue before message id */
#define	TPQGETBYMSGID	0x00008		/* dequeue by msgid */
#define	TPQMSGID	0x00010		/* get msgid of enq/deq message */
#define	TPQPRIORITY	0x00020		/* set/get message priority */
#define	TPQTOP		0x00040		/* enqueue at queue top */
#define	TPQWAIT		0x00080		/* wait for dequeuing */
#define	TPQREPLYQ	0x00100		/* set/get reply queue */
#define	TPQTIME_ABS	0x00200		/* set absolute time */
#define	TPQTIME_REL	0x00400		/* set absolute time */
#define	TPQGETBYCORRID	0x00800		/* dequeue by corrid */
#define	TPQPEEK		0x01000		/* peek */

#ifndef _TMDLLENTRY
#define _TMDLLENTRY
#endif
#ifndef _TM_FAR
#define _TM_FAR
#endif

#if defined(__cplusplus)
extern "C" {
#endif

extern int _TMDLLENTRY tpenqueue _((char _TM_FAR *qspace, char _TM_FAR *qname, TPQCTL _TM_FAR *ctl, char _TM_FAR *data, long len, long flags));
extern int _TMDLLENTRY tpdequeue _((char _TM_FAR *qspace, char _TM_FAR *qname, TPQCTL _TM_FAR *ctl, char _TM_FAR * _TM_FAR *data, long _TM_FAR *len, long flags));
#if defined(_TMPROTOTYPES) && !defined(_H_SYS_TIME) && !defined(_SYS_TIME_INCLUDED)
struct tm;
#endif
extern long _TMDLLENTRY gp_mktime _((struct tm _TM_FAR *));

#if defined(__cplusplus)
}
#endif


/* THESE MUST MATCH THE DEFINITIONS IN qm.h */
#define QMEINVAL	-1
#define QMEBADRMID	-2
#define QMENOTOPEN	-3
#define QMETRAN		-4
#define QMEBADMSGID	-5
#define QMESYSTEM	-6
#define QMEOS		-7
#define QMEABORTED	-8
#define QMENOTA		QMEABORTED
#define QMEPROTO	-9
#define QMEBADQUEUE	-10
#define QMENOMSG	-11
#define QMEINUSE	-12
#define QMENOSPACE	-13

/* END QUEUED MESSAGES ADD-ON */
#endif

/* START EVENT BROKER MESSAGES */
#define TPEVSERVICE	0x00000001
#define TPEVQUEUE	0x00000002
#define TPEVTRAN	0x00000004
#define TPEVPERSIST	0x00000008

/* Subscription Control structure */
struct tpevctl_t {
	long flags;
	char name1[XATMI_SERVICE_NAME_LENGTH];
	char name2[XATMI_SERVICE_NAME_LENGTH];
	TPQCTL qctl;
};
typedef struct tpevctl_t TPEVCTL;

/* Function prototypes */
#if defined(__cplusplus)
extern "C" {
#endif

extern long	_TMDLLENTRY tpsubscribe _((char *eventexpr, char *filter, TPEVCTL *ctl, long flags));
extern int	_TMDLLENTRY tpunsubscribe _((long subscription, long flags));
extern int	_TMDLLENTRY tppost _((char *eventname, char *data, long len, long flags));

#if defined(__cplusplus)
}
#endif

/* END EVENT BROKER MESSAGES */

/* 
 * BEGIN buildserver section
 * 
 * WARNING: Modification or use of these structures in any way, may
 *          cause system failures. DO NOT USE!
 */
struct tmdsptchtbl_t {
	char  *svcname;
	char  *funcname;
	void (*svcfunc) _((TPSVCINFO *));
	TM32I index;
	char  flag;
};

#define TMSRVRFLAG_COBOL 0x00000001
struct tmsvrargs_t {
	struct xa_switch_t	 *xa_switch;
	struct tmdsptchtbl_t *tmdsptchtbl;		/* Created by buildserver				*/
	TM32U flags;							/* Set by buildserver					*/
	int  (*initfunc)  _((int, char **));	/* Consult your Tuxedo documentation	*/
	void (*donefunc)  _((void));			/* BEFORE modifying these values...		*/
	int  (*runsrvr)   _((int));				/* reserved for system use - DO NOT USE */
	void (*reserved1) _(());				/* reserved for system use - DO NOT USE */
	void (*reserved2) _(());				/* reserved for system use - DO NOT USE */
	void (*reserved3) _(());				/* reserved for system use - DO NOT USE */
	void (*reserved4) _(());				/* reserved for system use - DO NOT USE */
};

#if defined(__cplusplus)
extern "C" {
#endif
extern void _TMDLLENTRY _tmsetup _(( int *argcp, char **argv, struct tmsvrargs_t *tmsvrargs ));
extern int _TMDLLENTRY _tmstartserver _(( int argc, char **argv, struct tmsvrargs_t* tmsvrargs ));
extern struct tmsvrargs_t *_tmgetsvrargs _((void));
#if defined(__cplusplus)
}
#endif
/* END buildserver section */

#endif
