
/* --------------------- usrinc/hlinkapi.h -------------------- */
/*                                                              */
/*              Copyright (c) 2000 - 2004 Tmax Soft Co., Ltd    */
/*                   All Rights Reserved                        */
/*                                                              */
/* ------------------------------------------------------------ */

#ifndef _TMAX_HLINKAPI_H_
#define _TMAX_HLINKAPI_H_

#include <time.h>

#ifndef _WIN32
#define __cdecl
#endif

/* DATA LOGGING TYPE */
#define TMAX_REQUEST            1
#define TMAX_RESPONSE           2
#define RGW_REQUEST             3
#define RGW_RESPONSE            4
#define BID_MESSAGE             5
#define ROP_MESSAGE             6
#define TMAX_PGMREQUEST         7
#define TMAX_PGMRESPONSE        8

/* struct for data logging */
struct logheader {
    int    type;
    int    len;
    int    errcode;
    time_t time;
    int    seconds;
    char   luname[8];
};
typedef struct logheader LOGHEADER;

/* struct for host link process */
struct hlprocinfo {
    int    pid;
    int    innum;
    int    outnum;
    int    line_status;
    char   label[20];
    char   linkname[8];
};
typedef struct hlprocinfo HLPROCINFO;

/* struct for host link session status */
struct hlsessinfo {
    char   luname[8];
    char   wsname[8];
    char   lutype[12];
    char   svcname[16];
    int    status;            /* 0x50: LU-LU, 0x51: LU-SSCP, 0x52: DOWN, 
	                         0x53: CSDN,  0x54: ACTLU,   0x55: INACTLU, 
	                         0x56: NSPE */
    int    send;              /* 1 : host send */
    int    direction;         /* 0 : inbound lu, 1=outbound lu */
    int    count;             /* process count */
};
typedef struct hlsessinfo HLSESSINFO;

#define HOST_TRANS_LENGTH    8
#define HOST_PROGRAM_LENGTH  8
#define TPGWINFO_SIZE        sizeof(struct tpgwinfo)

struct tpgwinfo {
    char  svc[XATMI_SERVICE_NAME_LENGTH];   /* relay or tpacall service name */
    char  trxid[HOST_TRANS_LENGTH];         /* host transaction id. */
    char  pgmname[HOST_PROGRAM_LENGTH];     /* host program name */
};
typedef struct tpgwinfo TPGWINFO_T;


#if defined (__cplusplus)
extern "C" {
#endif

int __cdecl tpgethlinksvr(int shmkey);
int __cdecl tpgethlinkproc(int svrn, HLPROCINFO *info);
HLSESSINFO *__cdecl tpgethlinkinfo(int svrn);

#if defined (__cplusplus)
}
#endif


#endif  /* _TMAX_HLINKAPI_H_ */

