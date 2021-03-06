
/* ---------------------- usrinc/WinTmax.h -------------------- */
/*								*/
/*              Copyright (c) 2000 - 2004 Tmax Soft Co., Ltd	*/
/*                   All Rights Reserved  			*/
/*								*/
/* ------------------------------------------------------------ */

#ifndef _TMAX_WINTMAX_H
#define _TMAX_WINTMAX_H

#include        <process.h>
#include        <winsock2.h>
#include        <windows.h>

#if defined(__cplusplus)
extern "C" {
#endif
int __cdecl WinTmaxSend(int recvContext, 
	char *svc, char *data, long len, long flags);
int __cdecl WinTmaxSetContext(HANDLE winhandle, 
	unsigned int msgtype, int slot);
int __cdecl WinTmaxStart(TPSTART_T *tpinfo);
int __cdecl WinTmaxEnd(void);
#if defined(__cplusplus)
}
#endif

#endif
