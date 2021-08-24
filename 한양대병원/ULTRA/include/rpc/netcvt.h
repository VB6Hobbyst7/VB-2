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

/* #ident	"@(#) gp/libgp/rpc/netcvt.h	$Revision: 1.1 $" */

#ifndef _RPC_NETCVT_H
#define _RPC_NETCVT_H

#ifndef TMENV_H
#include <tmenv.h>
#endif

#ifndef NOWHAT
static char	h_rpc_netcvt[] = "@(#) gp/libgp/rpc/netcvt.h	$Revision: 1.1 $";
#endif

/*    @(#) gp/libgp/rpc/netcvt.h	$Revision: 1.1 $"   */

/*
 *  		PROPRIETARY NOTICE (Combined)
 *  
 *  This source code is unpublished proprietary information
 *  constituting, or derived under license from AT&T's Unix(r) System V.
 *  In addition, portions of such source code were derived from Berkeley
 *  4.3 BSD under license from the Regents of the University of
 *  California.
 *  
 *  
 *  
 *  		Copyright Notice 
 *  
 *  Notice of copyright on this source code product does not indicate 
 *  publication.
 *  
 *  	(c) 1986,1987,1988,1989  Sun Microsystems, Inc.
 *  	(c) 1983,1984,1985,1986,1987,1988,1989  AT&T.
 *  	          All rights reserved.
 */

/*
 * netcvt.h
 * Includes the network conversion macros
 */
#if !defined(_TML_ENDIAN) && !defined(ntohl) && !defined(lint)
/*
 * Macros for number representation conversion.
 */
#define	ntohl(x)	(x)
#define	ntohs(x)	(x)
#define	htonl(x)	(x)
#define	htons(x)	(x)
#endif


#if defined(__cplusplus)
extern "C" {
#endif

#if !defined(ntohl) && (defined(_TML_ENDIAN) || defined(lint))

#if (defined(_TMNOVELL_TCP) || defined(_TMWGONG_TCP) || defined(_TM_WINSOCKAPI)) && (defined(_TM_WIN) || defined(WIN32))
#if !defined(_IN)
#ifndef _TMWGONG_TCP
#ifdef WIN32
extern unsigned short __stdcall htons (unsigned short);
extern unsigned short __stdcall ntohs (unsigned short);
#else
extern unsigned short far pascal htons (unsigned short);
extern unsigned short far pascal ntohs (unsigned short);
#endif
#else
/* They are defined as such by PATHWAY in sockets.h */
#ifdef WIN32
extern unsigned __stdcall htons( unsigned );
extern unsigned __stdcall ntohs( unsigned );
#else
extern unsigned far pascal htons( unsigned );
extern unsigned far pascal ntohs( unsigned );
#endif
#endif /* _TMWGONG_TCP */
#ifdef WIN32
extern unsigned long  __stdcall ntohl (unsigned long);
extern unsigned long  __stdcall htonl (unsigned long);
#else
extern unsigned long  far pascal ntohl (unsigned long);
extern unsigned long  far pascal htonl (unsigned long);
#endif
#endif
#else

#ifdef _TMPROTOTYPES
extern u_short	ntohs(u_short), htons(u_short);
#if defined(__osf__)
extern u_int	ntohl(u_int), htonl(u_int);
#else
extern u_long	ntohl(u_long), htonl(u_long);
#endif
#else
extern u_short	ntohs(), htons();
extern u_long	ntohl(), htonl();
#endif
#endif
#endif

#if defined(__cplusplus)
}
#endif


#endif	/* _RPC_NETCVT_H */
