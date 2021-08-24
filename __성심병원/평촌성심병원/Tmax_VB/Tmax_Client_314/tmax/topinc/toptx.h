
/* ------------------------- topinc/toptx.h ------------------- */
/*								*/
/*              Copyright (c) 2000 Tmax Soft Co., Ltd		*/
/*                   All Rights Reserved  			*/
/*								*/
/* ------------------------------------------------------------ */

#ifndef _TOPEND_TX_H
#define _TOPEND_TX_H


/*
 * tx_* return values
 */
#define TX_SHUTDOWN   3 /* TMC shutdown */
#define TX_COMMIT_LOGGED 2 /* Early commit return */
#define TX_NOT_SUPPORTED 1 /* No early commit return */
#define TX_OK         0 /* Normal execution */
#define TX_LOCAL     -1 /* Local txn in progress */
#define TX_OUTSIDE   -1 /* Local txn in progress */
#define TX_ROLLBACK  -2 /* Can't commit - rolled back */
#define TX_MIXED     -3 /* Partially committed and rolled back */
#define TX_HAZARD    -4 /* May have been heuristically completed */
#define TX_PROTOCOL_ERROR -5 /* Improper context */
#define TX_ERROR     -6 /* Transient error */
#define TX_RM_ERROR  -6 /* An RM returned error */
#define TX_FAIL      -7 /* Fatal error */
#define TX_TM_ERROR  -7 /* TM encountered an error */
#define TX_EINVAL    -8 /* Invalid argument value */
#define TX_NO_BEGIN -100 /* transaction committed but new
                            transaction could not be started */
#define TX_ROLLBACK_NO_BEGIN (TX_ROLLBACK+TX_NO_BEGIN)
                         /* transaction rollback but new
                            transaction could not be started */
#define TX_MIXED_NO_BEGIN (TX_MIXED+TX_NO_BEGIN)
                         /* mixed plus new transaction could
                            not be started */
#define TX_HAZARD_NO_BEGIN (TX_HAZARD+TX_NO_BEGIN)
                         /* hazard plus new transaction could
                            not be started */
/*
 * commit control values
*/
typedef long COMMIT_RETURN;
#define TX_COMMIT_COMPLETED       0
#define TX_COMMIT_DECISION_LOGGED 1

/* Alias old typedef
 *
 */
#define COMMIT_CONTROL COMMIT_RETURN

/*
 * transaction control modes
*/
typedef long TRANSACTION_CONTROL;
#define TX_UNCHAINED 0
#define TX_CHAINED   1

#define XIDDATASIZE 128 /* XID data size in bytes */
#define MAXGTRIDSIZE 64 
#define MAXBQUALSIZE 64 
#define XIDSIZE (XIDDATASIZE + (3 * sizeof(long)))
struct xid_t {
   long formatID;      /* format code: 0 -- ISO,
        positive -- private, -1 -- NULLXID */
   long gtrid_length;      /* length of global tran. ID */
   long bqual_length;      /* length of branch qualifier */
   char data[XIDDATASIZE]; /* gtrid and bqual */
};
typedef struct xid_t XID;
extern XID NULLXID;  /* Null transaction value */


/*
 * Structure populated by tx_info()
 */
struct tx_info_t {
   XID xid;
   COMMIT_RETURN when_return;
   TRANSACTION_CONTROL trx_control;
};
typedef struct tx_info_t TXINFO;
/*
 * Extern declarations for tx_* functions
 */

#ifdef __cplusplus
extern "C" {
#endif

#if defined(__STDC__) || defined(__cplusplus)
extern int tx_begin(void);    /* Begin a global trans. */
extern int tx_close(void);    /* Close all RMs */
extern int tx_commit(void);   /* Commit the transaction */
extern int tx_info(TXINFO *); /* Get context and XID */
extern int tx_open(void);     /* Open all RMs */
extern int tx_rollback(void); /* Roll back the trans. */
/* The routines referenced by the following four externs
   are restricted and reserved for use in the future.    */
extern int tx_set_commit_return(COMMIT_RETURN);
extern int tx_set_transaction_control(TRANSACTION_CONTROL);
extern void tx_xid_format(XID *, char *);
extern int tx_recover(void);
#else /* non-ANSI C */
extern int tx_begin();    /* Begin a global trans. */
extern int tx_close();    /* Close all RMs */
extern int tx_commit();   /* Commit the transaction */
extern int tx_info();     /* Get context and XID */
extern int tx_open();     /* Open all RMs */
extern int tx_rollback(); /* Roll back the trans. */
/* The routines referenced by the following four externs
   are restricted and reserved for use in the future.    */
extern int tx_set_commit_return();
extern int tx_set_transaction_control();
extern void tx_xid_format();
extern int tx_recover();
#endif /* ifdef __STDC__ */

/* COBOL APIs */
extern void TXBEGIN();
extern void TXCOMMIT();
extern void TXROLLBACK();

#ifdef __cplusplus
}
#endif

#endif /* _TOPEND_TX_H */

