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

/*      Copyright (c) 1990 Unix System Laboratories, Inc.
        All rights reserved

        THIS IS UNPUBLISHED PROPRIETARY
        SOURCE CODE OF Unix System Laboratories, Inc.
        The copyright notice above does not
        evidence any actual or intended
        publication of such source code.
*/
#ifndef TMTYPES_H
#define TMTYPES_H
/* #ident	"@(#) tuxedo/include/tmtypes.h	$Revision: 1.1 $" */

#ifndef TMENV_H
#include <tmenv.h>
#endif
#ifndef NOWHAT
static	char	h_tmtypes[] = "@(#) tuxedo/include/tmtypes.h	$Revision: 1.1 $";
#endif

/*
 *	DEFINITIONS NEEDED BY INTERNAL TUXEDO PROGRAMS.
 *
 *	Warning: This is an internal TUXEDO header file and
 *	should not be included in any part of any user program.
 *	The definitions contained in this file MAY CHANGE from
 *	one release to the next.
 */

#define TMTYPELEN	8
#define TMSTYPELEN	16

#define TM_MAX_ITYPES	16		/* maximum # of internal types */
#define TMENCODE	0x00000001	/* message is encoded */
#define TMDECODE	0x00000002	/* message is not encoded */
#define TMCMPRS		0x00000004	/* compress message */
#define TMEXPND		0x00000008	/* expand message */

#if defined(__cplusplus)
extern "C" {
#endif

struct tmtype_sw_t {
	char type[TMTYPELEN];	   /* type of buffer */
	char subtype[TMSTYPELEN];  /* sub-type of buffer */
	long dfltsize;		   /* default size of buffer */
	/* buffer initialization function pointer */
	int (_TMDLLENTRY *initbuf) _((char _TM_FAR *, long));	
	/* buffer re-initialization function pointer */
	int (_TMDLLENTRY *reinitbuf) _((char _TM_FAR *, long));	
	/* buffer un-initialization function pointer */
	int (_TMDLLENTRY *uninitbuf) _((char _TM_FAR *, long));	
	/* pre-send buffer manipulation func pointer */
	long (_TMDLLENTRY *presend) _((char _TM_FAR *, long, long));	
	/* post-send buffer manipulation func pointer */
	void (_TMDLLENTRY *postsend) _((char _TM_FAR *, long, long));	
	/* post-receive buffer manipulation func pointer*/
	long (_TMDLLENTRY *postrecv) _((char _TM_FAR *, long, long));	
	/* encode/decode function pointer */
	long (_TMDLLENTRY *encdec) _((int, char _TM_FAR *, long, char _TM_FAR *, long));	
	/* routing function pointer */
	int (_TMDLLENTRY *route) _((char _TM_FAR *, char _TM_FAR *, char _TM_FAR *, long, char _TM_FAR *));		
	/* buffer filtering function pointer */
	int (_TMDLLENTRY *filter) _((char _TM_FAR *, long, char _TM_FAR *, long));
	/* buffer formatting function pointer */
	int (_TMDLLENTRY *format) _((char _TM_FAR *, long, char _TM_FAR *, char _TM_FAR *, long));
	/* this space reserved for future expansion */
	void (_TMDLLENTRY *reserved[10]) _((void));
};

#if defined(__cplusplus)
}
#endif

typedef struct tmtype_sw_t TMTYPESW;

_TMITUX2WSC extern TMTYPESW _TM_NEAR tm_itypesw[];  /* internal buffer type switch */

/********************************************************************/
/* EVERYTHING BELOW THIS LINE SHOULD BE MOVED TO AN INTERNAL HEADER */
/********************************************************************/
#define _tmsw ((TUX->_TUX__tm_swindex >= TM_MAX_ITYPES) ? \
	&TUX->_TUX_tm_typeswp[TUX->_TUX__tm_swindex-TM_MAX_ITYPES] : \
	&tm_itypesw[TUX->_TUX__tm_swindex])

#define _tmtype			(_tmsw->type)
#define _tmsubtype		(_tmsw->subtype)
#define _tmsize			(_tmsw->dfltsize)
#ifndef lint
#define _tminitbuf(p, l)	\
	((*(_tmsw->initbuf == NULL ? _dfltinitbuf : _tmsw->initbuf)) \
	 ((p), (l)))
#define _tmreinitbuf(p, l)	\
	((*(_tmsw->reinitbuf == NULL ? _dfltinitbuf : _tmsw->reinitbuf)) \
	 ((p), (l)))
#define _tmuninitbuf(p, l)	\
	((*(_tmsw->uninitbuf == NULL ? _dfltinitbuf : _tmsw->uninitbuf)) \
	 ((p), (l)))
#define _tmpresend(p, l, ml)	\
	((*(_tmsw->presend == NULL ? _dfltblen : _tmsw->presend)) \
	 ((p), (l), (ml)))
#define _tmpostsend(p, l, ml)	\
	((*(_tmsw->postsend == NULL ? _dfltpostsend : _tmsw->postsend)) \
	 ((p), (l), (ml)))
#define _tmpostrecv(p, rl, ml)	\
	((*(_tmsw->postrecv == NULL ? _dfltblen : _tmsw->postrecv)) \
	 ((p), (rl), (ml)))
#define _tmencdec(op, e, el, o, ol)	\
	((*(_tmsw->encdec == NULL ? _dfltencdec : _tmsw->encdec)) \
	 ((op), (e), (el), (o), (ol)))
#define _tmroute(n, s, d, l, g)	\
	((*(_tmsw->route == NULL ? _dfltroute : _tmsw->route)) \
	 ((n), (s), (d), (l), (g)))
#define _tmfilter(p, dl, e, el)		\
	((*(_tmsw->filter == NULL ? _dfltfilter : _tmsw->filter)) \
	 ((p), (dl), (e), (el)))
#define _tmformat(p, dl, f, r, rl)	\
	((*(_tmsw->format == NULL ? _dfltformat : _tmsw->format)) \
	 ((p), (dl), (f), (r), (rl)))
#else 
extern	int	_tminitbuf _((char *, long));
extern	int	_tmreinitbuf _((char *, long));
extern	int	_tmuninitbuf _((char *, long));
extern	long	_tmpresend _((char *, long, long));
extern	void	_tmpostsend _((char *, long, long));
extern	long	_tmpostrecv _((char *, long, long));
extern	long	_tmencdec _((int, char *, long, char *, long));
extern	int	_tmroute _((char *, char *, char *, long, char *));
extern	int	_tmfilter _((char *, long, char *, long));
extern	int	_tmformat _((char *, long, char *, char *, long));
#endif


#if defined(__cplusplus)
extern "C" {
#endif

extern  int     _TMDLLENTRY _dfltinitbuf _((char _TM_FAR *, long));
extern  long    _TMDLLENTRY _dfltblen _((char _TM_FAR *, long, long));
extern  void    _TMDLLENTRY _dfltpostsend _((char _TM_FAR *, long, long));
extern  long    _TMDLLENTRY _dfltencdec _((int, char _TM_FAR *, long, char _TM_FAR *, long));
extern  int     _TMDLLENTRY _dfltroute _((char _TM_FAR *, char _TM_FAR *, char _TM_FAR *, long, char _TM_FAR *));
extern	int	_TMDLLENTRY _dfltfilter _((char _TM_FAR *, long, char _TM_FAR *, long));
extern	int	_TMDLLENTRY _dfltformat _((char _TM_FAR *, long, char _TM_FAR *, char _TM_FAR *, long));
extern  long    _TMDLLENTRY _strpresend _((char _TM_FAR *, long, long));
extern  long    _TMDLLENTRY _strencdec _((int, char _TM_FAR *, long, char _TM_FAR *, long));
extern	int	_TMDLLENTRY _sfilter _((char _TM_FAR *, long, char _TM_FAR *, long));
extern	int	_TMDLLENTRY _sformat _((char _TM_FAR *, long, char _TM_FAR *, char _TM_FAR *, long));
extern  int     _TMDLLENTRY _finit _((char _TM_FAR *, long));
extern  int     _TMDLLENTRY _freinit _((char _TM_FAR *, long));
extern  int     _TMDLLENTRY _funinit _((char _TM_FAR *, long));
extern  long    _TMDLLENTRY _fpresend _((char _TM_FAR *, long, long));
extern  void    _TMDLLENTRY _fpostsend _((char _TM_FAR *, long, long));
extern  long    _TMDLLENTRY _fpostrecv _((char _TM_FAR *, long, long));
extern  long    _TMDLLENTRY _fencdec _((int, char _TM_FAR *, long, char _TM_FAR *, long));
extern  int     _TMDLLENTRY _froute _((char _TM_FAR *, char _TM_FAR *, char _TM_FAR *, long, char _TM_FAR *));
extern	int	_TMDLLENTRY _ffilter _((char _TM_FAR *, long, char _TM_FAR *, long));
extern	int	_TMDLLENTRY _fformat _((char _TM_FAR *, long, char _TM_FAR *, char _TM_FAR *, long));
extern  int     _TMDLLENTRY _finit32 _((char _TM_FAR *, long));
extern  int     _TMDLLENTRY _freinit32 _((char _TM_FAR *, long));
extern  int     _TMDLLENTRY _funinit32 _((char _TM_FAR *, long));
extern  long    _TMDLLENTRY _fpresend32 _((char _TM_FAR *, long, long));
extern  void    _TMDLLENTRY _fpostsend32 _((char _TM_FAR *, long, long));
extern  long    _TMDLLENTRY _fpostrecv32 _((char _TM_FAR *, long, long));
extern  long    _TMDLLENTRY _fencdec32 _((int, char _TM_FAR *, long, char _TM_FAR *, long));
extern  int     _TMDLLENTRY _froute32 _((char _TM_FAR *, char _TM_FAR *, char _TM_FAR *, long, char _TM_FAR *));
extern	int	_TMDLLENTRY _ffilter32 _((char _TM_FAR *, long, char _TM_FAR *, long));
extern	int	_TMDLLENTRY _fformat32 _((char _TM_FAR *, long, char _TM_FAR *, char _TM_FAR *, long));
extern  int     _TMDLLENTRY _vinit _((char _TM_FAR *, long));
extern  int     _TMDLLENTRY _vreinit _((char _TM_FAR *, long));
extern  long    _TMDLLENTRY _vpresend _((char _TM_FAR *, long, long));
extern  long    _TMDLLENTRY _vencdec _((int, char _TM_FAR *, long, char _TM_FAR *, long));
extern  int     _TMDLLENTRY _vroute _((char _TM_FAR *, char _TM_FAR *, char _TM_FAR *, long, char _TM_FAR *));
extern	int	_TMDLLENTRY _vfilter _((char _TM_FAR *, long, char _TM_FAR *, long));
extern	int	_TMDLLENTRY _vformat _((char _TM_FAR *, long, char _TM_FAR *, char _TM_FAR *, long));
extern  int     _TMDLLENTRY _vinit32 _((char _TM_FAR *, long));
extern  int     _TMDLLENTRY _vreinit32 _((char _TM_FAR *, long));
extern  long    _TMDLLENTRY _vpresend32 _((char _TM_FAR *, long, long));
extern  long    _TMDLLENTRY _vencdec32 _((int, char _TM_FAR *, long, char _TM_FAR *, long));
extern  int     _TMDLLENTRY _vroute32 _((char _TM_FAR *, char _TM_FAR *, char _TM_FAR *, long, char _TM_FAR *));
extern	int	_TMDLLENTRY _vfilter32 _((char _TM_FAR *, long, char _TM_FAR *, long));
extern	int	_TMDLLENTRY _vformat32 _((char _TM_FAR *, long, char _TM_FAR *, char _TM_FAR *, long));
extern	int	_TMDLLENTRY _TPINITinit _((char _TM_FAR *, long));
extern	int	_TMDLLENTRY _TPINITreinit _((char _TM_FAR *, long));
extern	long	_TMDLLENTRY _TPINITpresend _((char _TM_FAR *, long , long));
extern	long	_TMDLLENTRY _TPINITencdec _((int, char _TM_FAR *, long, char _TM_FAR *, long));

extern	int	_TMDLLENTRY AEWaddtypesw _((TMTYPESW _TM_FAR *newtype));
extern TMTYPESW _TM_FAR * _TMDLLENTRY _tmtypeswaddr _((void));

extern	int	_tuxdftcmpexp _((int, char *, long *, char *, long));

#if defined(__cplusplus)
}
#endif


#endif
