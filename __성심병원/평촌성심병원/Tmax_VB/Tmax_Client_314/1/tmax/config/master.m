###############################################################
#                                                             #
#         Sample Configuration File for Tmax System           #
#         =========================================           #
#                                                             #
#     Copyright(c) 2002 TmaxSoft Inc. All rights reserved     #
#                                                             #
###############################################################

*DOMAIN
tmax1           SHMKEY =79990, MAXUSER = 1000, MINCLH=1, MAXCLH=3,
                TPORTNO=8888, BLOCKTIME=30,
                SECURITY = "USER_AUTH", OWNER = tmax,
#               MAXSPR=64, MAXSVR=32,
#               MAXCPC=32,
#               CLICHKINT = 10, IDLETIME = 30,
#               MAXSACALL = 10, MAXCACALL = 10,
#               MAXCONV_NODE = 16, MAXCONV_SERVER = 8,
#               IPCPERM = 0600,
#               TXTIME = 10,
#               NLIVEINQ = 15

*NODE
tmax	        TMAXDIR = "c:\\tmax",
                APPDIR  = "c:\\tmax\appbin",
                PATHDIR = "c:\\tmax\path",
                TLOGDIR = "c:\\tmax\log\tlog",
                ULOGDIR = "c:\\tmax\log\ulog",
                SLOGDIR = "c:\\tmax\log\slog"
#               ENVFILE = "c:\\tmax\config\env",
#               TPORTNO2 = 8899,
#               TPORTNO3 = 8909,
#               TPORTNO4 = 8919,
#               TPORTNO5 = 8929
#               REALSVR = "realtest", RSCPC = 2
#               DOMAINNAME = tmax1,
#               CLHQTIMEOUT = 10,
#               LOGOUTSVC = logout


*SVRGROUP
svg1            NODENAME = "tmax"

### svg for load balancing ###
#svg2            NODENAME = "tmax",  COUSIN = svg3
#svg3            NODENAME = "tmax"

### svg for RQ ###
#svg4            NODENAME = "tmax", SVRTYPE = RQMGR, CPC = 1

#*RQ
#rqtest          SVGNAME = svg4, QSIZE = 16, BUFFERING = Y, BOOT = WARM, FILEPATH = "c:\tmax\path\rqtest"

*SERVER
svr1            SVGNAME = svg1, MIN=1, MAX=5, ASQCOUNT = 10, MAXQCOUNT = 10,
                CLOPT = "-o c:/temp/$(SVR).$(PID) -e c:/temp/$(SVR).$(PID)"

#svr2           SVGNAME = svg1, MIN=1, MAX=5, CLOPT = "-o c:/temp/$(SVR).$(PID) -e c:/temp/$(SVR).$(PID)"
#svr3           SVGNAME = svg2, RESTART = Y

### server for RQ ###
#rqsvr          SVGNAME = svg1, MIN=1, MAX=5

### server for conversation ###
#svrconv        SVGNAME = svg2, CONV = Y


### servers for REALSVR ###
#realtest        SVGNAME = svg1, MIN=1, MAX=2, SVRTYPE = REALSVR, MAXRSTART = 0

### server for UCS ###
#ucsServer       SVGNAME = svg2, MIN = 1, MAX = 3, SVRTYPE = UCS
#ucsServer2      SVGNAME = svg1, MIN = 2, SVRTYPE = UCS_LOAD

*SERVICE
SDLTOUPPER      SVRNAME = svr1, PRIO = 2
SDLTOLOWER      SVRNAME = svr1, PRIO = 1
#TOUPPER        SVRNAME = svr2
#TOLOWER        SVRNAME = svr2
#FDLTOUPPER     SVRNAME = svr3, ROUTING = rout1
#FDLTOLOWER     SVRNAME = svr3, ROUTING = rout1

#*ROUTING
#rout1           FIELD = INPUT, SUBTYPE = FIELD,
#                RANGES = "'best':svg2, *:svg3"