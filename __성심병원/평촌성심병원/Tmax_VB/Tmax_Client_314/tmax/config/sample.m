
###############################################################
#                                                             #
#         Sample Configuration File for Tmax System           #
#         =========================================           #
#                                                             #
#     Copyright(c) 2003 TmaxSoft Inc. All rights reserved     #
#                                                             #
###############################################################
*DOMAIN
tmax1        	SHMKEY   = 77214, MINCLH   = 1, MAXCLH = 1,
             	TPORTNO  = 8888,  BLOCKTIME = 30

*NODE
tmax		TMAXDIR	= "c:\tmax",
		APPDIR	= "c:\tmax\appbin",
             	PATHDIR  = "c:\tmax\path",
             	TLOGDIR  = "c:\tmax\log\tlog",
             	ULOGDIR  = "c:\tmax\log\ulog",
             	SLOGDIR  = "c:\tmax\log\slog"

*SVRGROUP
svg1		NODENAME = tmax

svgora     	NODENAME = tmax,
             	DBNAME   = ORACLE,
             	OPENINFO = "Oracle_XA+Acc=P/scott/tiger+SesTm=60",
             	TMSNAME  = tms_ora,
             	MINTMS   = 1
### for RQ ###
#rqsvg           NODENAME = tmax, SVGTYPE = RQMGR, CPC = 4

#*RQ
#rq1             SVGNAME = rqsvg, BOOT = WARM, FILEPATH="c:\tmax\appbin\rq1.dat"

*SERVER
svr1		SVGNAME  = svg1
svr2		SVGNAME  = svg1
svr3		SVGNAME  = svg1
#svr_ucs        SVGNAME = svg1, SVRTYPE = UCS
#svr_conv       SVGNAME = svg1, CONV = YES
#svr_rq         SVGNAME = svg1

fdltest		SVGNAME  = svgora, MIN = 1, MAX = 2
#			CLOPT = "-o c:/temp/$(SVR).$(PID) -e c:/temp/$(SVR).$(PID)"
sdltest		SVGNAME  = svgora, MIN = 1, MAX = 2


*SERVICE
SDLTOUPPER	SVRNAME = svr1
SDLTOLOWER	SVRNAME = svr1

TOUPPER		SVRNAME = svr2
TOLOWER		SVRNAME = svr2

FDLTOUPPER	SVRNAME = svr3
FDLTOLOWER	SVRNAME = svr3

#LOGIN          SVRNAME = svr_ucs
#TOUPPER_CONV   SVRNAME = svr_conv
#TPENQ          SVRNAME = svr_rq
#TPDEQ          SVRNAME = svr_rq

FDLINS		SVRNAME = fdltest
FDLSEL		SVRNAME = fdltest
FDLUPT		SVRNAME = fdltest
FDLDEL		SVRNAME = fdltest

SDLINS		SVRNAME = sdltest
SDLSEL		SVRNAME = sdltest
SDLUPT		SVRNAME = sdltest
SDLDEL		SVRNAME = sdltest