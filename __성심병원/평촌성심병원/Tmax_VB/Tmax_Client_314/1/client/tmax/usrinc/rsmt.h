
/* ------------------------ usrinc/rsmt.h --------------------- */
/*								*/
/*           Copyright (c) 2002 - 2004 Tmax Soft Co., Ltd	*/
/*                   All Rights Reserved  			*/
/*								*/
/* ------------------------------------------------------------ */

#ifndef _TMAX_RSMT_H
#define _TMAX_RSMT_H
#ifndef _TMAX_MTLIB
#define _TMAX_MTLIB	1
#endif

#ifndef _WIN32
#include <sys/time.h>
#define __cdecl
#endif

typedef struct {
	long	urcode;
	int	errcode;
	int	msgtype;
	int	cd;
	int	len;
	char	*data;
} UCSMSGINFO;

typedef int (__cdecl *UcsCallback)(UCSMSGINFO*);

#if defined (__cplusplus)
extern "C" {
#endif

#ifndef _TMAX_KERNEL
int __cdecl usermain(int argc, char *argv[]);
#endif
int __cdecl tpschedule(int sec);
int __cdecl tpuschedule(int usec);

/* register and unregister monitoring fds */
int __cdecl tpsetfd(int fd);
int __cdecl tpissetfd(int fd);
int __cdecl tpclrfd(int fd);

/* register and unregister callback function */
int __cdecl tpregcb(UcsCallback);
int __cdecl tpunregcb();

/* thread related function */
int __cdecl tmax_thr_create(void *(*func)(void *), void *argp, int flags);
int __cdecl tmax_thr_terminate(int thrid, int flags);

#if defined (__cplusplus)
}
#endif


#endif	/* _TMAX_RSMT_H */
