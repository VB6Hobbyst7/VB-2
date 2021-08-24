
/* ------------------------ usrinc/tmaxapi.h ------------------ */
/*								*/
/*              Copyright (c) 2000 - 2004 Tmax Soft Co., Ltd	*/
/*                   All Rights Reserved  			*/
/*								*/
/* ------------------------------------------------------------ */

#ifndef _TMAXAPI_H
#define _TMAXAPI_H

#ifndef _CE_MODULE
#include <sys/types.h>
#endif
#include <usrinc/atmi.h>
#ifdef _WIN32
#ifdef _CE_MODULE
#include <winsock.h>
#else
#include <winsock2.h>
#endif  /* _CE_MODULE */
#include <usrinc/svct.h>
#include <usrinc/sdl.h>
#else
#ifndef ORA_PROC
#include <sys/socket.h>
#endif
#define __cdecl
#endif

/* client logout type */
#define CLIENT_CLOSE_NORMAL	0
#define CLIENT_CLOSE_ABNORMAL	1
#define CLIENT_PRUNED		2

/* RQ Sub-queue type */
#define TMAX_ANY_QUEUE		0
#define TMAX_FAIL_QUEUE		1
#define TMAX_REQ_QUEUE		2
#define TMAX_RPLY_QUEUE		3
#define TMAX_MAX_QUEUE          4

extern char _rq_sub_queue_name[TMAX_MAX_QUEUE][XATMI_SERVICE_NAME_LENGTH];

/* RQ related macros */
#define RQ_NAME_LENGTH		16

/* unsolicited msg type */
#define UNSOL_TPPOST		1
#define UNSOL_TPBROADCAST	2
#define UNSOL_TPNOTIFY		3
#define UNSOL_TPSENDTOCLI	4
#define UNSOL_ANY		5

/* Check SVCINFO cmds */
#define ISSVC_FORWARDED	0x00000001
#define ISSVC_NOREPLY	0x00000002

/* TPEVCTL ctl_flags */
#define	TPEV_SVC	0x00000001
#define	TPEV_PROC	0x00000002

struct tpevctl {
    long ctl_flags;
    long post_flags;
    char svc[XATMI_SERVICE_NAME_LENGTH];
    char qname[RQ_NAME_LENGTH];
};

typedef struct tpevctl TPEVCTL;
typedef void __cdecl Unsolfunc(char *, long, long);
#define TPUNSOLERR      ((Unsolfunc *) -1)

/* Multicast call related structures */
struct svglist {
    int	ns_entry;	/* number of entries of s_list */
    int	nf_entry;	/* number of entries of f_list */
    int *s_list;	/* list of server group numbers */
    int *f_list;	/* list of server group numbers */
};

/* My svrinfo */
#ifndef TMAX_NAME_SIZE
#define TMAX_NAME_SIZE          16
#endif

typedef struct {
    int nodeno;	/* node index */
    int svgi;	/* server group index; unique in the node */
    int svri;	/* server index; unique in the node */
    int spri;	/* server process index; unique in the node */
    int spr_seqno;	/* server process seqno ; unique in the server */
    int min, max;	/* min/max server process number */
    int clhi;	/* for RDP only, corresponding CLH id */
    char nodename[TMAX_NAME_SIZE];
    char svgname[TMAX_NAME_SIZE];
    char svrname[TMAX_NAME_SIZE];
    char reserved_char[TMAX_NAME_SIZE];
    /* for more detail use tmadmin API */
} TMAXSVRINFO;

#ifdef _WIN32
typedef int (__cdecl *WinTmaxCallback)(WPARAM, LPARAM);
#endif

/* Macro functions */
#define tpalivechk()	tmax_chk_conn(0)

#if defined (__cplusplus)
extern "C" {
#endif

/* ----- unsolicited messaging API ----- */
long __cdecl tpsubscribe(char *eventexpr, char *filter, TPEVCTL *ctl, long flags);
long __cdecl tpsubscribe2(char *eventexpr, char *svcname, long flags);
int __cdecl tpunsubscribe(long sd, long flags);
int __cdecl tppost(char *eventname, char *data, long len, long flags);
int __cdecl tpbroadcast(char *lnid, char *usrname, char *cltname, char *data,
	    long len, long flags);
Unsolfunc *__cdecl tpsetunsol(Unsolfunc *func);
int __cdecl tpsetunsol_flag(int flag);
int __cdecl tpgetunsol(int type, char **data, long *len, long flags);
int __cdecl tpclearunsol(void);
int __cdecl tpchkunsol(void);

/* ----- RQS API -------- */
int __cdecl tpenq(char *qname, char *svc, char *data, long len, long flags);
int __cdecl tpdeq(char *qname, char *svc, char **data, long *len, long flags);
int __cdecl tpqstat(char *qname, long type);
int __cdecl tpqsvcstat(char *qname, char *svc, long type);
int __cdecl tpextsvcname(char *data, char *svc);
int __cdecl tpextsvcinfo(char *data, char *svc, int *type, int *errcode);
int __cdecl tpreissue(char *qname, char *filter, long flags);
char *__cdecl tpsubqname(int type);

/* ----- server API -------- */
int __cdecl tpgetminsvr(void);
int __cdecl tpgetmaxsvr(void);
int __cdecl tpgetmaxuser(void);
int __cdecl tpgetsvrseqno(void);
int __cdecl tpgetmysvrid(void);
int __cdecl tpgetmysvrno(void);
int __cdecl tpgetmaxuser(void);
int __cdecl tpsendtocli(int clid, char *data, long len, long flags);
int __cdecl tpgetclid(void);
int __cdecl tpgetpeer_ipaddr(struct sockaddr *name, int *namelen);
int __cdecl tpchkclid(int clid);
int __cdecl tmax_clh_maxuser(void);
int __cdecl tmax_my_svrinfo(TMAXSVRINFO*);
int __cdecl tmax_cind2clid(int cind);
char *__cdecl tpgetmynode(int *nodeno);
char *__cdecl tpgetmysvg(void);

/* ----- etc API ----------- */
int __cdecl tp_sleep(int sec);
int __cdecl tp_usleep(int usec);
int __cdecl tpset_timeout(int sec);
int __cdecl tpget_timeout(void);
int __cdecl tmaxreadenv(char *file, char *label);
char *__cdecl tpgetenv(char* str);
int __cdecl tpputenv(char* str);
int __cdecl tpgetsockname(struct sockaddr *name, int *namelen);
int __cdecl tpgetpeername(struct sockaddr *name, int *namelen);
int __cdecl tpgetactivesvr(char *nodename, char **outbufp);
int __cdecl tperrordetail(int i);
int __cdecl tpreset(void);
int __cdecl tptobackup(void);
struct svglist *__cdecl 
    tpmcall(char *qname, char *svc, char *data, long len, long flags);
struct svglist *__cdecl tpgetsvglist(char *svc, long flags);
int __cdecl tpsvgcall(int svgno, char *qname, 
	char *svc, char *data, long len, long flags);
int __cdecl tpflush(void);
char *__cdecl tmaxlastsvc(char *idata, char *odata, long flags);
int __cdecl tpgetorgnode(int clid);
int __cdecl tpgetorgclh(int clid);
char *__cdecl tpgetnodename(int nodeno);
int __cdecl tpgetnodeno(char *nodename);
int __cdecl tpgetasize(char *data);
int __cdecl tpgettype(char *data);
char * __cdecl tpgetsubtype(char *data);
int __cdecl tpgetcliaddr(int clid, int *ip, int *port, long flags);
int __cdecl tmax_chk_conn(int timeout);


#if defined(_WIN32)
int __cdecl WinTmaxAcall(TPSTART_T *sinfo, HANDLE wHandle, unsigned int msgType,
	char *svc, char *sndbuf, int len, int flags);
int __cdecl WinTmaxAcall2(TPSTART_T *sinfo, WinTmaxCallback fn,
	char *svc, char *sndbuf, int len, int flags);
#endif

#if !defined(_TMAX_KERNEL) && !defined(_TMAX_RCA_H)
/* ------- User supplied routines ---------- */
int __cdecl tpsvrinit(int argc, char *argv[]);
int __cdecl tpsvrdone(void);
void __cdecl tpsvctimeout(TPSVCINFO *msg);
#endif

/* 
   Internal functions: ONLY BE CALLED FROM AUTOMATICALLY 
   GENERATED STUB FILES. DO NOT DIRECTLY CALL THESE FUNCTIONS.
 */
int __cdecl get_clhfd(void);
int __cdecl tmax_chk_svcinfo(int cmd);
int __cdecl _tmax_main(int argc, char *argv[]);
int __cdecl _tmax_cob_main(int argc, char *argv[]);
#if defined(_WIN32)
int __cdecl _tmax_regfn(void *initFn, void *doneFn, void *timeoutFn, void *userMainFn);
int __cdecl _tmax_regtab(int svcTabSz, _svc_t *svcTab, int funcTabSz, void *funcTab);
int __cdecl _tmax_regsdl(int _sdl_table_size2, struct _sdl_struct_s *_sdl_table2,
	int _sdl_field_table_size2, struct _sdl_field_s *_sdl_field_table2);
#endif
int __cdecl _double_encode(char *in, char *out);
int __cdecl _double_decode(char *in, char *out);
/* --- power builder interface API --- */
int __cdecl _make_struct_from_pbindata(char *subtype, char *tpidata, int ilen, char *indata);
int __cdecl _make_field_from_pbindata(char **tpidata, char *indata);
int __cdecl _make_pbodata_from_struct(char *subtype, char *odata, int olen, char *tpodata);
int __cdecl _make_pbodata_from_field(char *form, char *odata, char *tpodata);
int __cdecl _make_pbfodata_from_field(char *fform, char *fodata, char *tpodata);
int __cdecl _get_value_from_pbsdata(char *cur, char *vald);
int __cdecl _get_name_value_from_pbfdata(char *cur, int *n2);
int __cdecl _get_name_from_form(char *cur);
int __cdecl _insert_value_to_pbodata(int type, char *out, char *in, int asize, int asize2);


#if defined (__cplusplus)
}
#endif

#endif       /* end of _TMAXAPI_H  */
