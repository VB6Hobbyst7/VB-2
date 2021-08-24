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

/*	Copyright (c) 1989 AT&T; 1991 USL
	All rights reserved

	THIS IS UNPUBLISHED PROPRIETARY
	SOURCE CODE OF USL
	The copyright notice above does not
	evidence any actual or intended
	publication of such source code.
*/
/* #ident	"@(#) gp/libgp/userlog.h	$Revision: 1.1 $" */
#ifndef USERLOG_H
#define USERLOG_H

#ifndef TMENV_H
#include <tmenv.h>
#endif

#ifndef NOWHAT
static	char	h_userlog[] = "@(#) gp/libgp/userlog.h	$Revision: 1.1 $";
#endif

/*
 *	Warning: This TUXEDO header file should not be changed in any way;
 *	doing so will destroy its compatibility with TUXEDO programs and
 *	libraries.
 */


#if defined(__cplusplus)
extern "C" {
#endif

#if (defined(_TM_WIN) || defined(_TM_OS2)) && !defined(_TMDLL)
extern char _TM_FAR * _TM_FAR * _TMDLLENTRY gp_get_proc_name_addr(void);
#define proc_name	(*gp_get_proc_name_addr())
#else

#if defined(_TM_NETWARE)

extern char **_gp_get_proc_name_addr _((void));
#define proc_name	(*_gp_get_proc_name_addr())

#define GET_PROC_NAME()		(GP->_GP_proc_name)
#define SET_PROC_NAME(v)	GP->_GP_proc_name = (v)

#else /* NOT _TM_NETWARE */

_TMIGP extern _TM_THREADVAR char *proc_name;
#define GET_PROC_NAME()		(GP->_GP_proc_name = proc_name)
#define SET_PROC_NAME(v)	GP->_GP_proc_name = proc_name = (v)

#endif /* _TM_NETWARE */

#endif

/* PRINTFLIKE1 */
#ifdef WIN32
extern int _TM_FAR __stdcall _gp_write_eventlog _(( char *, va_list ));
extern int _TM_FAR __stdcall userlog _((char _TM_FAR *, ...));
#else
extern int _TM_FAR _TM_CDECL userlog _((char _TM_FAR *, ...));
#endif

#if defined(__cplusplus)
}
#endif

#ifdef WIN32
/* used by Windows NT event log routine */
#define WRITE_EVENT_LOG_ONLY 0x0
#define WRITE_ULOG_ONLY		 0x1
#endif

#endif
