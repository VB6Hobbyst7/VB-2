
/* ------------------------- usrinc/userlog.h ----------------- */
/*								*/
/*              Copyright (c) 2000 - 2004 Tmax Soft Co., Ltd	*/
/*                   All Rights Reserved  			*/
/*								*/
/* ------------------------------------------------------------ */

#ifndef _TMAX_USERLOG_H
#define _TMAX_USERLOG_H

#ifndef _CE_MODULE
#include <time.h>
#endif

#ifndef _WIN32
#define __cdecl
#endif

#if defined (__cplusplus)
extern "C" {
#endif

int __cdecl userlog(const char *fmt, ...);
int __cdecl UserLog(const char *fmt, ...);
#ifndef _CE_MODULE 
int __cdecl userlog2(char *data, int len);
#endif
int __cdecl ulogsync();

#if defined (__cplusplus)
}
#endif

#endif
