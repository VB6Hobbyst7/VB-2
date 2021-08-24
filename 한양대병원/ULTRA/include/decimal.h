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

/*	Copyright (c) 1990 Unix System Laboratories, Inc.
	All rights reserved

	THIS IS UNPUBLISHED PROPRIETARY
	SOURCE CODE OF Unix System Laboratories, Inc.
	The copyright notice above does not
	evidence any actual or intended
	publication of such source code.
*/
#ifndef DECIMAL_H
#define DECIMAL_H
/* #ident	"@(#) gp/libgp/decimal.h	$Revision: 1.1 $" */
#ifndef TMENV_H
#include <tmenv.h>
#endif

#ifndef NOWHAT
static char	h_decimal[] = "@(#) gp/libgp/decimal.h	$Revision: 1.1 $";
#endif

/*
 *	DEFINITIONS NEEDED BY USER APPLICATION PROGRAMS.
 *
 *	Warning:  This header file should not be changed in any way.
 *	Doing so will destroy the compatibility with TUXEDO programs
 *	and libraries.
 */

/***************************************************************************
 *
 *                    RELATIONAL DATABASE SYSTEMS, INC.
 *
 *  COPYRIGHT (c) 1981-1986 RELATIONAL DATABASE SYSTEMS, INC., MENLO PARK, 
 *  CALIFORNIA.  All rights reserved.  No part of this work covered by the 
 *  copyright hereon may be reproduced or used in any form or by any means 
 *  -- graphic, electronic, or mechanical, including photocopying, 
 *  recording, taping, or information storage and retrieval systems --
 *  without permission of RDS.
 *
 *  Title:	decimal.h
 *  Sccsid:	@(#) gp/libgp/decimal.h	$Revision: 1.1 $"
 *  Description:
 *		Header file for decimal data type.
 *
 ***************************************************************************
 */


/*
 * Unpacked Format (format for program usage)
 *
 *    Signed exponent "dec_exp" ranging from  -64 to +63
 *    Separate sign of mantissa "dec_pos"
 *    Base 100 digits (range 0 - 99) with decimal point
 *	immediately to the left of first digit.
 */

#define DECSIZE 16
#define DECUNKNOWN -2

struct decimal
    {
    short dec_exp;		/* exponent base 100		*/
    short dec_pos;		/* sign: 1=pos, 0=neg, -1=null	*/
    short dec_ndgts;		/* number of significant digits	*/
    char  dec_dgts[DECSIZE];	/* actual digits base 100	*/
    };
typedef struct decimal dec_t;

/*
 *  A decimal null will be represented internally by setting dec_pos
 *  equal to DECPOSNULL
 */

#define DECPOSNULL	(-1)

/*
 * DECLEN calculates minumum number of bytes
 * necessary to hold a decimal(m,n)
 * where m = total # significant digits and
 *	 n = significant digits to right of decimal
 */

#define DECLEN(m,n)	(((m)+((n)&1)+3)/2)
#define DECLENGTH(len)	DECLEN(PRECTOT(len),PRECDEC(len))

/*
 * DECPREC calculates a default precision given
 * number of bytes used to store number
 */

#define DECPREC(size)	(((size-1)<<9)+2)

/* macros to look at and make encoded decimal precision
 *
 *  PRECTOT(x)		return total precision (digits total)
 *  PRECDEC(x) 		return decimal precision (digits to right)
 *  PRECMAKE(x,y)	make precision from total and decimal
 */

#define PRECTOT(x)	(((x)>>8) & 0xff)
#define PRECDEC(x)	((x) & 0xff)
#define PRECMAKE(x,y)	(((x)<<8) + (y))

/*
 * Packed Format  (format in records in files)
 *
 *    First byte =
 *	  top 1 bit = sign 0=neg, 1=pos
 *	  low 7 bits = Exponent in excess 64 format
 *    Rest of bytes = base 100 digits in 100 complement format
 *    Notes --	This format sorts numerically with just a
 *		simple byte by byte unsigned comparison.
 *		Zero is represented as 80,00,00,... (hex).
 *		Negative numbers have the exponent complemented
 *		and the base 100 digits in 100's complement
 */

/* Internal functions */
extern int	_gp_bycmpr _((char *, char *, int));
extern void	_gp_bycopy _((char *, char *, int));
extern void	_gp_byfill _((char *, int, int));

/*
 * Define decimal functions to internal names to prevent conflict with
 * other libraries that provide the same functions
 */
#define decadd(A,B,C)		_gp_decadd(A,B,C)
#define deccmp(A,B)		_gp_deccmp(A,B)
#define deccopy(A,B)		_gp_deccopy(A,B)
#define deccvasc(A,B,C)		_gp_deccvasc(A,B,C)
#define deccvdbl(A,B)		_gp_deccvdbl(A,B)
#define deccvflt(A,B)		_gp_deccvflt(A,B)
#define deccvint(A,B)		_gp_deccvint(A,B)
#define deccvlong(A,B)		_gp_deccvlong(A,B)
#define decdiv(A,B,C)		_gp_decdiv(A,B,C)
#define dececvt(A,B,C,D)	_gp_dececvt(A,B,C,D)
#define decfcvt(A,B,C,D)	_gp_decfcvt(A,B,C,D)
#define decload(A,B,C,D,E)	_gp_decload(A,B,C,D,E)
#define decmul(A,B,C)		_gp_decmul(A,B,C)
#define decprec(A)		_gp_decprec(A)
#define decround(A,B)		_gp_decround(A,B)
#define decsub(A,B,C)		_gp_decsub(A,B,C)
#define dectoasc(A,B,C,D)	_gp_dectoasc(A,B,C,D)
#define dectodbl(A,B)		_gp_dectodbl(A,B)
#define dectoflt(A,B)		_gp_dectoflt(A,B)
#define dectoint(A,B)		_gp_dectoint(A,B)
#define dectolong(A,B)		_gp_dectolong(A,B)
#define dectrunc(A,B)		_gp_dectrunc(A,B)
#define lddecimal(A,B,C)	_gp_lddecimal(A,B,C)
#define stdecimal(A,B,C)	_gp_stdecimal(A,B,C)

/* External functions */

#if defined(__cplusplus)
extern "C" {
#endif

extern int	_TMDLLENTRY _gp_decadd _((dec_t _TM_FAR *, dec_t _TM_FAR *, dec_t _TM_FAR *));
extern int	_TMDLLENTRY _gp_deccmp _((dec_t _TM_FAR *, dec_t _TM_FAR *));
extern void	_TMDLLENTRY _gp_deccopy _((dec_t _TM_FAR *, dec_t _TM_FAR *));
extern int	_TMDLLENTRY _gp_deccvasc _((char _TM_FAR *, int, dec_t _TM_FAR *));
extern int	_TMDLLENTRY _gp_deccvdbl _((double, dec_t _TM_FAR *));
extern int	_TMDLLENTRY _gp_deccvflt _((double, dec_t _TM_FAR *));
extern int	_TMDLLENTRY _gp_deccvint _((int, dec_t _TM_FAR *));
extern int	_TMDLLENTRY _gp_deccvlong _((long, dec_t _TM_FAR *));
extern int	_TMDLLENTRY _gp_decdiv _((dec_t _TM_FAR *, dec_t _TM_FAR *, dec_t _TM_FAR *));
extern char	_TM_FAR * _TMDLLENTRY _gp_dececvt _((dec_t _TM_FAR *, int, int _TM_FAR *, int _TM_FAR *));
extern char	_TM_FAR * _TMDLLENTRY _gp_decfcvt _((dec_t _TM_FAR *, int, int _TM_FAR *, int _TM_FAR *));
extern int	_TMDLLENTRY _gp_decload _((dec_t _TM_FAR *, int, int, char _TM_FAR *, int));
extern int	_TMDLLENTRY _gp_decmul _((dec_t _TM_FAR *, dec_t _TM_FAR *, dec_t _TM_FAR *));
extern int	_TMDLLENTRY _gp_decprec _((dec_t _TM_FAR *));
extern void	_TMDLLENTRY _gp_decround _((dec_t _TM_FAR *, int));
extern int	_TMDLLENTRY _gp_decsub _((dec_t _TM_FAR *, dec_t _TM_FAR *, dec_t _TM_FAR *));
extern int	_TMDLLENTRY _gp_dectoasc _((dec_t _TM_FAR *, char _TM_FAR *, int, int));
extern int	_TMDLLENTRY _gp_dectodbl _((dec_t _TM_FAR *, double _TM_FAR *));
extern int	_TMDLLENTRY _gp_dectoflt _((dec_t _TM_FAR *, float _TM_FAR *));
extern int	_TMDLLENTRY _gp_dectoint _((dec_t _TM_FAR *, int _TM_FAR *));
extern int	_TMDLLENTRY _gp_dectolong _((dec_t _TM_FAR *, long _TM_FAR *));
extern void	_TMDLLENTRY _gp_dectrunc _((dec_t _TM_FAR *, int));
extern int	_TMDLLENTRY _gp_lddecimal _((char _TM_FAR *, int, dec_t _TM_FAR *));
extern void	_TMDLLENTRY _gp_stdecimal _((dec_t _TM_FAR *, char _TM_FAR *, int));

#if defined(__cplusplus)
}
#endif


#endif
