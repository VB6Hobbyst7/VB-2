
/* ------------------------ usrinc/tx.h ----------------------- */
/*								*/
/*              Copyright (c) 2000 - 2004 Tmax Soft Co., Ltd	*/
/*                   All Rights Reserved  			*/
/*								*/
/* ------------------------------------------------------------ */

#ifndef _TMAX_TX_H
#define _TMAX_TX_H

#ifndef _WIN32
#define __cdecl
#endif

#ifdef _TMAX_KERNEL
#include <include/xa.h>
#else
#include <usrinc/xa.h>
#endif

/*
 * Definitions for tx_ routines
 */
/* commit return values */
typedef long COMMIT_RETURN;
#define TX_COMMIT_COMPLETED 0
#define TX_COMMIT_DECISION_LOGGED 1

/* transaction control values */
typedef long TRANSACTION_CONTROL;
#define TX_UNCHAINED 0
#define TX_CHAINED 1

/* type of transaction timeouts */
typedef long TRANSACTION_TIMEOUT;

/* transaction state values */
typedef long TRANSACTION_STATE;
#define TX_NOT_ACTIVE -1
#define TX_ACTIVE 0
#define TX_TIMEOUT_ROLLBACK_ONLY 1
#define TX_ROLLBACK_ONLY 2

/* structure populated by tx_info() */
struct tx_info_t {
	XID	xid;
	COMMIT_RETURN when_return;
	TRANSACTION_CONTROL transaction_control;
	TRANSACTION_TIMEOUT transaction_timeout;
	TRANSACTION_STATE   transaction_state;
};
typedef struct tx_info_t TXINFO;


#if defined (__cplusplus)
extern "C" {
#endif
int __cdecl tx_begin ();
int __cdecl tx_close ();
int __cdecl tx_commit ();
int __cdecl tx_info (TXINFO *);
int __cdecl tx_open ();
int __cdecl tx_rollback ();
int __cdecl tx_set_commit_return (COMMIT_RETURN);
int __cdecl tx_set_transaction_control (TRANSACTION_CONTROL);
int __cdecl tx_set_transaction_timeout (TRANSACTION_TIMEOUT);
#if defined (__cplusplus)
}
#endif

/*
 * tx_ () return codes (transaction manager reports to application)
 */
#define TX_NOT_SUPPORTED	1	/* option not supported */
#define TX_OK			0	/* normal execution */
#define TX_OUTSIDE		-1	/* application is in an RM local
					   transaction */
#define TX_ROLLBACK		-2	/* transaction was rolled back */
#define TX_MIXED		-3	/* transaction was partially committed
					   and partially rolled back */
#define TX_HAZARD		-4	/* transaction may have been partially
					   committed and partially rolled back*/
#define TX_PROTOCOL_ERROR	-5	/* routine invoked in an improper
					   context */
#define TX_ERROR		-6	/* transient error */
#define TX_FAIL			-7	/* fatal error */
#define TX_EINVAL		-8	/* invalid arguments were given */
#define TX_COMMITTED		-9	/* transaction has heuristically
					   committed */
#define TX_ESYSTEM              -99     /* for internal use */
#define TX_NO_BEGIN		-100	/* transaction committed plus new
					   transaction could not be started */
#define TX_ROLLBACK_NO_BEGIN	(TX_ROLLBACK+TX_NO_BEGIN)
					/* transaction rollback plus new
					   transaction could not be started */
#define TX_MIXED_NO_BEGIN		(TX_MIXED+TX_NO_BEGIN)
					/* mixed plus new transaction could
					   not be started */
#define TX_HAZARD_NO_BEGIN		(TX_HAZARD+TX_NO_BEGIN)
					/* hazard plus new transaction could
					   not be started */
#define TX_COMMITTED_NO_BEGIN		(TX_COMMITTED+TX_NO_BEGIN)
					/* heuristically committed plus new
					   transaction could not be started */
#endif /* ifndef TX_H */

/*
 * End of tx.h header
 */
