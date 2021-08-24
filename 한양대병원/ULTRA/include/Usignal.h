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

/*	Copyright (c) 1984 AT&T; 1991 USL
	All rights reserved
*/
#ifndef USIGNAL_H
#define USIGNAL_H

/* #ident	"@(#) gp/libgp/Usignal.h	$Revision: 1.1 $" */
#ifndef TMENV_H
#include <tmenv.h>
#endif

#ifndef NOWHAT
static	char	h_Usignal[] = "@(#) gp/libgp/Usignal.h	$Revision: 1.1 $";
#endif

/*
 *	DEFINITIONS NEEDED BY USER APPLICATION PROGRAMS.
 *
 *	Warning: This header file should not be changed in any way.
 *	Doing so will destroy the compatibility with TUXEDO programs
 *	and libraries.
 */

/*
 *	macros for signal deferral/reinstatement
 */


#ifndef USIGTYP
#if !defined(__cplusplus) && defined(uts)
#define USIGTYP int
#else
#define USIGTYP void
#endif
#endif

#if defined(__cplusplus)
extern "C" {
#endif

extern void Usiginit _((void));
extern USIGTYP (*Usignal _((int, USIGTYP (*)(int)))) _((int));

extern void _gp_dosigs _((void));
extern UGDEFERLEVEL _((void));

#if defined(__cplusplus)
}
#endif


#if defined(_TM_NETWARE)

extern int *_gp_get_gp_sigdefer_addr _((void));
extern int *_gp_get_gp_sigspending_addr _((void));
#define _gp_sigdefer		(*_gp_get_gp_sigdefer_addr())
#define _gp_sigspending		(*_gp_get_gp_sigspending_addr())

#define GET_SIGDEFER()		(GP->_GP__gp_sigdefer)
#define SET_SIGDEFER(v)		GP->_GP__gp_sigdefer = (v)
#define GET_SIGSPENDING()	(GP->_GP__gp_sigspending)
#define SET_SIGSPENDING(v)	GP->_GP__gp_sigspending = (v)

#else /* NOT _TM_NETWARE */

_TMIGP extern int _gp_sigdefer;
_TMIGP extern int _gp_sigspending;

#define GET_SIGDEFER()		(GP->_GP__gp_sigdefer = _gp_sigdefer)
#define SET_SIGDEFER(v)		GP->_GP__gp_sigdefer = _gp_sigdefer = (v)
#define GET_SIGSPENDING()	(GP->_GP__gp_sigspending = _gp_sigspending)
#define SET_SIGSPENDING(v)	GP->_GP__gp_sigspending = _gp_sigspending = (v)

#endif /* _TM_NETWARE */

	/* macro to defer signals ... note that deferrals stack */
#define UDEFERSIGS() {_gp_sigdefer++;}

/* macro to reinstate signals ... note that reinstatement really unstacks */
#define URESUMESIGS() {if(_gp_sigdefer>0)_gp_sigdefer--; if((_gp_sigdefer<=0)&&_gp_sigspending) _gp_dosigs();}

	/* macro to pop deferral stack completely ...
		ensures that sigs will be processed */
#define UENSURESIGS() {_gp_sigdefer = 0; if(_gp_sigspending) _gp_dosigs();}


	/* the following macro is useful for setjmp/longjmp situations */

		/* macro to set defer level */
#define USDEFERLEVEL(newlev) {_gp_sigdefer=(newlev<0)?0:newlev; if ((_gp_sigdefer<=0)&&_gp_sigspending) _gp_dosigs();}

#endif
