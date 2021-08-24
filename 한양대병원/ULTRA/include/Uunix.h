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

/*	Copyright (c) 1993 USL
	All rights reserved
*/
/* #ident	"@(#) gp/libgp/Uunix.h	$Revision: 1.1 $" */

#ifndef UUNIX_H
#define UUNIX_H

#ifndef TMENV_H
#include <tmenv.h>
#endif

#ifndef NOWHAT
static	char	h_Uunix[] = "@(#) gp/libgp/Uunix.h	$Revision: 1.1 $";
#endif

/*
 *	DEFINITIONS NEEDED BY USER APPLICATION PROGRAMS.
 *	
 *	Warning: This header file should not be changed in any way.
 *	Doing so will destroy the compatibility with TUXEDO programs
 *	and libraries.
 */

#define UUNIXMIN 0
#define UCLOSE	1
#define UCREAT	2
#define UEXEC	3
#define UFCTNL	4
#define UFORK	5
#define ULSEEK	6
#define UMSGCTL	7
#define UMSGGET	8
#define UMSGSND	9
#define UMSGRCV	10
#define UOPEN	11
#define UPLOCK	12
#define UREAD	13
#define USEMCTL	14
#define USEMGET	15
#define USEMOP	16
#define USHMCTL	17
#define USHMGET	18
#define USHMAT	19
#define USHMDT	20
#define USTAT	21
#define UWRITE	22
#define USBRK	23
#define USYSMUL 24
#define UWAIT	25
#define UKILL	26
#define UTIME	27
#define UMKDIR	28
#define ULINK	29
#define UUNLINK	30
#define UUNAME  31
#define UNLIST  32
#define UUNIXMAX 33


#if defined(__cplusplus)
extern "C" {
#endif
#ifndef Uunixerr
_TMIGP extern	_TM_THREADVAR int	Uunixerr;
#endif
extern	void	Uunix_err _((char *));
_TMIGP extern	char _TMCONST *_TMCONST Uunixmsg[];

#if defined(__cplusplus)
}
#endif


#if defined(_TMPROTOTYPES)
#include <stddef.h>
#include <stdlib.h>
#include <sys/types.h>
#include <fcntl.h>
#include <string.h>
#include <time.h>
#include <signal.h>
#if !defined(_TMDOWN) && !defined(THINK_C) && !defined(applec) && !defined(NeXT)
#include <sys/ipc.h>
#include <sys/sem.h>
#include <sys/shm.h>
#undef IN
#include <sys/msg.h>
#undef msg
#if !defined(_TM_NETWARE) && !defined(WIN32)
#include <unistd.h>
#endif
#if !defined(_TM_NETWARE) && !defined(_as400_)
#include <sys/times.h>
#endif
#endif

#ifdef __cplusplus
#define entry ____entry__
#include <search.h>
#undef entry
#ifdef _TMEDG
#include <wait.h>
#else
#if !defined(_TMDOWN) && !defined(__alpha)
#include <osfcn.h>
#endif
#endif

extern "C" {
extern char *strdup(const char *);
#if defined(__cpp_stdc)
extern  int   setpgid(pid_t, pid_t);
#endif
#if !defined(_TMDOWN)
extern char *tempnam(const char *, const char *);
#endif
}
#if defined(__cpp_stdc) && !defined(_TMNOCRYPTHDR)
#include <crypt.h>
#else
extern char *crypt _((const char* pw, const char* salt));
#endif

#else
/* not c++ */
#if !defined(_TMDOWN) && !defined(THINK_C) && !defined(applec) && !defined(_TM_NETWARE)
#include <sys/wait.h>
#endif
#if !defined(_TMNOCRYPTHDR)
#include <crypt.h>
#else
extern char *crypt _((const char* pw, const char* salt));
#endif
#ifdef _TMNOSTRDUP
extern char *strdup _((const char *));
#endif
#endif

#if defined(__cplusplus)
extern "C" {
#endif
extern char *Ustrerror(int);
#ifndef WIN32
extern size_t strftime(char *, size_t, const char *, const struct tm *);
#undef strerror
extern char *strerror(int);
#endif
#if defined(__cplusplus)
}
#endif

#else
/* classic c */
#include <string.h>
#if !defined(_TMDOWN) && !defined(NeXT)
#include <memory.h>
#endif
#if defined(sun41) || defined(r3000) || defined(r4000)
#include <sys/wait.h>
#endif
#if defined(_IBMR2) || defined(__osf__) || defined(NeXT) || defined(hpux) || defined(sun41)
extern  void     abort();
#else
extern  int     abort();
extern  int     abs();
#endif
extern	int	access();
extern  unsigned int    alarm();
extern  int     atoi();
extern  long	atol();
extern  double	atof();
extern	double	strtod();
#if defined(_IBMR2) || defined(__osf__) || defined(NeXT) || defined(hpux) || defined(sun41)
extern	void	*bsearch();
#else
extern	char	*bsearch();
#endif

#if defined(sun41) || defined(_SEQUENT_)
#include <malloc.h>
#else
#if defined(ultrix) || defined(hpux) || defined(_IBMR2) || defined(__osf__) || defined(NeXT)
extern	void	*malloc();
extern	void	*calloc();
extern	void	*realloc();
extern	void	free();
#else
extern	char	*malloc();
extern	char	*calloc();
extern	char	*realloc();
extern	void	free();
#endif
#endif

extern  int     chdir();
extern  int     close();
extern  int     creat();
extern  char    *crypt();
extern  int     dup();
#if !defined(_as400_)
#if !defined(_IBMR2)
extern	int	execl();
extern  int     execle();
extern  int     execlp();
#endif
extern	int	execv();
extern  int     execve();
extern  int     execvp();
#endif
extern  void    exit();
#ifdef sun
extern	char	*getpass();
#endif
extern	void	_exit();
#if !defined(_as400_)
extern  pid_t   fork();
#endif
extern  char    *gcvt();
extern	char	*getcwd();
extern  char    *getenv();
extern  gid_t  getegid();
extern  uid_t  geteuid();
extern  uid_t  getuid();
extern  gid_t  getgid();
extern  int     getopt();
extern  pid_t	getpid();
extern	int	ioctl();
#if !defined(_as400_)
extern  int     isatty();
#endif
extern	int	kill();
extern  off_t   lseek();
extern	int	msgget();
extern  int     msgsnd();
extern  int     msgrcv();
#if !defined(_as400_)
extern  int     nice();
#endif
#if !defined(_IBMR2)
extern  int     open();
#endif
extern  void    perror();
extern  int     pipe();
extern  int     putenv();
extern  void    qsort();
extern	int	rand();
extern  int     read();
extern  int     semget();
extern  int     setgid();
extern  pid_t   setpgrp();
extern	void	srand();
extern  int     setuid();
extern  char    *shmat();
extern	int	shmget();
extern  unsigned int    sleep();
extern  char    *strdup();
extern  long    strtol();
extern	char	*tmpnam();
extern	char	*mktemp();
#if !defined(__osf__)
extern  long    time();
#else
extern	time_t	time();
#endif
#if !defined(hpux) && !defined(__osf__)
extern  long    times();
#else
extern	clock_t	times();
#endif
#if !defined(_IBMR2)
extern  long    ulimit();
#endif
extern  mode_t  umask();
extern  int     unlink();
extern  pid_t   wait();
extern  int     write();
extern	char	*strdup();

#define const

extern char *Ustrerror();
#undef strerror
extern char *strerror();
/* don't define strftime because of problems with typedef for size_t */

#endif

#ifdef __cplusplus
extern	void	__cplusinit _((char *));
#else
#define	__cplusinit(a)
#endif

#if defined(__cplusplus)
extern "C" {
#endif
#if defined (_TM_NETWARE)
char ** _gp_get_optarg_addr(void);
int (*_gp_get_optind_addr(void));
int (*_gp_get_opterr_addr(void));
#define optarg	(*_gp_get_optarg_addr())
#define opterr	(*_gp_get_opterr_addr())
#define optind	(*_gp_get_optind_addr())
#else
/* global variables */
_TMIGP extern  char    *optarg;
_TMIGP extern  int     opterr;
_TMIGP extern  int     optind;
#endif
#if defined(__cplusplus)
}
#endif

#if defined(_TMPROTOTYPES) || defined(__cpp_stdc)
#include <limits.h>
#include <float.h>
/* define the values that used to appear in <values.h> */
#ifndef BITSPERBYTE
#define BITSPERBYTE CHAR_BIT
#endif
#ifndef MAXLONG
#define MAXLONG LONG_MAX
#endif
#ifndef MINLONG
#define MINLONG LONG_MIN
#endif
#ifndef MAXSHORT
#define MAXSHORT SHRT_MAX
#endif
#ifndef MINSHORT
#define MINSHORT SHRT_MIN
#endif
#ifndef	MAXINT
#define MAXINT		INT_MAX
#endif
#ifndef MININT
#define MININT		INT_MIN
#endif
#undef MAXFLOAT
#define MAXFLOAT FLT_MAX
#ifndef MAXDOUBLE
#define MAXDOUBLE DBL_MAX
#endif
#ifndef MINFLOAT
#define MINFLOAT FLT_MIN
#endif


#else

/* classic C   or  C++ */

#ifndef WIN32
#include <values.h>
#endif
#ifndef MINSHORT
#define MINSHORT	(-MAXSHORT-1)
#endif
#ifndef MININT
#define MININT		(-MAXINT-1)
#endif
#ifndef MINLONG
#define MINLONG		(-MAXLONG-1)
#endif

#ifndef _POSIX_NAME_MAX
#define _POSIX_NAME_MAX 14
#endif
#ifndef _POSIX_PATH_MAX
#define _POSIX_PATH_MAX 255
#endif

/* following kludge needed for a current bug in the 3b20 compiler */
#if u3b||u3b20d
#undef MAXSHORT
#define MAXSHORT 32767
#endif

/* don't need these - don't waste defining */
#undef DMAXEXP
#undef DMAXPOWTWO
#undef DMINEXP
#undef DSIGNIF
#undef FMAXEXP
#undef FMAXPOWTWO
#undef FMINEXP
#undef FSIGNIF
#undef M_LN2
#undef M_PI
#undef M_SQRT2
#undef X_EPS
#undef X_PLOSS
#undef X_TLOSS
#undef _DEXPLEN
#undef _EXPBASE
#undef _FEXPLEN
#undef _HIDDENBIT
#undef _IEEE
#undef _LENBASE
#endif

#define BITSPERLONG	(sizeof(long) * BITSPERBYTE)
#define BITSPERLONG32   (sizeof(TM32I) * BITSPERBYTE)
#ifndef MAXLONG32
#ifdef _TMLONG64
#define MAXLONG32 MAXINT
#define MINLONG32 MININT
#define xdr_TM32I xdr_int
#define xdr_TM32U xdr_u_int
#else
#define MAXLONG32 MAXLONG
#define MINLONG32 MINLONG
#define xdr_TM32I xdr_long
#define xdr_TM32U xdr_u_long
#endif
#endif

/* at this point, malloc is defined */
#ifndef _MALLOC_H
#define _MALLOC_H
#endif

#ifdef _as400_
#define GP_ROUNDUP_P(ptr,sz)	(ptr)
#else
#define GP_ROUNDUP_P(ptr,sz)	((((long)(ptr) + (sz) - 1) / (sz)) * (sz))
#endif
#define GP_ROUNDUP(amt,sz)	((((amt) + (sz) - 1) / (sz)) * (sz))
#define GP_ROUNDUP4(amt)	(((long)(amt) + 3L) & ~3L)
#define GP_ROUNDUP8(amt)	(((long)(amt) + 7L) & ~7L)

#ifdef _TMLONG64
#define ALIGN_LONG(x)   ((x)+(sizeof(long)-((x)%sizeof(long))))
#else
#define ALIGN_LONG(x)   (x)
#endif

#if defined(__cplusplus)
extern "C" {
#endif

extern char _TM_FAR * _TMDLLENTRY tuxgetenv _((char _TM_FAR *));
extern int _TMDLLENTRY tuxputenv _((char _TM_FAR *));
extern int _TMDLLENTRY tuxreadenv _((char _TM_FAR *, char _TM_FAR *));

#if defined(__cplusplus)
}
#endif

#endif
