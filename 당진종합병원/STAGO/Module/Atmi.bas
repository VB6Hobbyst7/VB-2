Attribute VB_Name = "atmi"
'/* -------------------------- atmi.h ------------------------- */
'/*                                                             */
'/*              Copyright (c) 2000 Tmax Soft Co., Ltd          */
'/*                   All Rights Reserved                       */
'/*                                                             */
'/* ----------------------------------------------------------- */

'/* Flags to tpinit() for Tuxedo compatability */
Const TPU_MASK = &H7    'unsolicited notification mask
Const TPU_SIG = &H1     'signal based notification
Const TPU_DIP = &H2     'dip-in based notification
Const TPU_IGN = &H4     'ignore unsolicited messages
Const TPSA_FASTPATH = &H8
Const TPSA_PROTECTED = &H10

'/* ---------- flags from API ----- */
'/* Most Significant Two Bytes are reserved for internal use */
Global Const TPNOFLAGS = &H0
Global Const TPNOBLOCK = &H1
Global Const TPSIGRSTRT = &H2
Global Const TPNOREPLY = &H4
Global Const TPNOTRAN = &H8
Global Const TPTRAN = &H10
Global Const TPNOTIME = &H20
Global Const TPNOGETANY = &H40
Global Const TPGETANY = &H80
Global Const TPNOCHANGE = &H100
Global Const TPBLOCK = &H200
Global Const TPFLOWCONTROL = &H400
Global Const TPSENDONLY = &H800
Global Const TPRECVONLY = &H1000
Global Const TPUDP = &H2000
Global Const TPRQS = &H4000
Global Const TPFUNC = &H8000

'/* --- flags used in tpstart() --- */
Const TPUNSOL_MASK = &H7
Const TPUNSOL_HND = &H1
Const TPUNSOL_IGN = &H2
Const TPUNSOL_POLL = &H4
Const TPUNIQUE = &H10
Const TPONLYONE = &H20

'/* Flags to tpreturn() */
Const TPFAIL = &H1
Const TPSUCCESS = &H2
Const TPEXIT = &H4
Const TPDOWN = &H8

'/* ------ flags for reply type check ----- */
Const TPREQ = 0
Const TPERR = -1

'/* -------- for Tuxedo Compatability ------- */
'/* Flags to tpscmt() - Valid TP_COMMIT_CONTROL characteristic values */
Const TP_CMT_LOGGED = &H1       '/* return after commit decision is logged */
Const TP_CMT_COMPLETE = &H2     '/* return after commit has completed */

'/* Return values to tpchkauth() */
Const TPNOAUTH = 0      '/* no authentication */
Const TPSYSAUTH = 1     '/* system authentication */
Const TPAPPAUTH = 2     '/* system and application authentication */

'/* unsolicited msg type */
Const TPPOST_MSG = 1
Const TPBROADCAST_MSG = 2
Const TPNOTIFY = 3
Const TPSENDTOCLI = 4

Const XATMI_SERVICE_NAME_LENGTH = 16    '/* where x must be > 15 */

Type tpsvcinfo
    Name    As String * XATMI_SERVICE_NAME_LENGTH
    data    As Long
    len     As Long
    flags   As Long
    cd      As Long
End Type
    
Declare Function gettperrno Lib "TMAX4GL.DLL" () As Long
Declare Function gettpurcode Lib "TMAX4GL.DLL" () As Long

Const TPEBADDESC = 2
Const TPEBLOCK = 3
Const TPEINVAL = 4
Const TPELIMIT = 5
Const TPENOENT = 6
Const TPEOS = 7
Const TPEPROTO = 9
Const TPESVCERR = 10
Const TPESVCFAIL = 11
Const TPESYSTEM = 12
Const TPETIME = 13
Const TPETRAN = 14
Const TPGOTSIG = 15
Const TPEITYPE = 17
Const TPEOTYPE = 18
Const TPEEVENT = 22
Const TPEMATCH = 23
Const TPENOREADY = 24
Const TPESECURITY = 25
Const TPEQFULL = 26
Const TPEQPURGE = 27
Const TPECLOSE = 28
Const TPESVRDOWN = 29
Const TPEPRESVC = 30
Const TPEMAXNO = 31

Const TPUNSOLERR As Long = -1


Global Const UNSOL_TPBROADCAST = &H2
'/* ---- flags used in conv[]: don't use dummy flags ----*/
Global Const TPEV_DISCONIMM = &H1
Global Const TPEV_SVCERR = &H2
Global Const TPEV_SVCFAIL = &H4
Global Const TPEV_SVCSUCC = &H8
Global Const TPEV_SENDONLY = &H20
Const TPCONV_DUMMY1 = &H800     '/* don't use this flag: TPSENDONLY */
Const TPCONV_DUMMY2 = &H1000    '/* don't use this flag: TPRECVONLY */
Const TPCONV_OUT = &H10000
Const TPCONV_IN = &H20000

Const X_OCTET = "X_OCTET"
Const X_C_TYPE = "X_C_TYPE"
Const X_COMMON = "X_COMMON"

Const TMTYPEFAIL = -1
Const TMTYPESUCC = 0

Const MAXTIDENT = XATMI_SERVICE_NAME_LENGTH '/* max len of identifier */

Const MAX_PASSWD_LENGTH = 16
Const MAX_MNAME_LENGTH = 16

Const MAXTIDENTPLUS2 = MAXTIDENT + 2
Const MAX_PASSWD_LENGTHPLUS2 = MAXTIDENT + 2

Type tpstart_t
    usrname     As String * MAXTIDENTPLUS2  '/* usr name */
    cltname     As String * MAXTIDENTPLUS2  '/* application client name */
    dompwd      As String * MAX_PASSWD_LENGTHPLUS2  '/* domain password */
    usrpwd      As String * MAX_PASSWD_LENGTHPLUS2  '/* passwd for usrid */
    flags       As Long
End Type

Global Const TMQNAMELEN = 15
Global Const TMQNAMELENPLUS1 = TMQNAMELEN + 1
Global Const TMMSGIDLEN = 32
Global Const TMCORRIDLEN = 32

Type clientid_t
    clientdata(4) As Long
End Type

Type tpqctl_t                           '/* control parameters to queue primitives */
    flags   As Long                     '/* indicates which of the values are set */
    deq_time    As Long                 '/* absolute/relative  time for dequeuing */
    priority    As Long                 '/* enqueue priority */
    diagnostic  As Long                 '/* indicates reason for failure */
    msgid       As String * TMMSGIDLEN  '/* id of message before which to queue */
    corrid      As String * TMCORRIDLEN '/* correlation id used to identify message */
    replyqueue  As String * TMQNAMELENPLUS1 '/* queue name for reply message */
    failurequeue    As String * TMQNAMELENPLUS1 '/* queue name for failure message */
    cltid   As clientid_t   '/* client identifier for originating client */
    urcode  As Long         '/* application user-return code */
    appkey  As Long         '/* application authentication client key */
End Type

Type tpevctl_t
    flags   As Long
    name1   As String * XATMI_SERVICE_NAME_LENGTH
    name2   As String * XATMI_SERVICE_NAME_LENGTH
    qctl    As tpqctl_t
End Type


'/* ----- client API ----- */
Declare Function tpstart Lib "TMAX4GL.DLL" (ByVal ptpinfo As Long) As Long
Declare Function tpend Lib "TMAX4GL.DLL" () As Long
Declare Function tpalloc Lib "TMAX4GL.DLL" (ByVal buftype As String, ByVal subtype As String, ByVal bufsize As Long) As Long
Declare Function tprealloc Lib "TMAX4GL.DLL" (ByVal pbuffer As Long, ByVal bufsize As Long) As Long
Declare Function tptypes Lib "TMAX4GL.DLL" (ByVal pbuffer As Long, buftype As String, subtype As String) As Long
Declare Sub tpfree Lib "TMAX4GL.DLL" (ByVal pbuffer As Long)
Declare Function tpcall Lib "TMAX4GL.DLL" (ByVal svcname As String, ByVal psendbuf As Long, ByVal sendlen As Long, pprecvbuf As Long, precvlen As Long, ByVal flags As Long) As Long
Declare Function tpacall Lib "TMAX4GL.DLL" (ByVal svcname As String, ByVal psendbuf As Long, ByVal sendlen As Long, ByVal flags As Long) As Long
Declare Function vb_tpgetrply Lib "TMAX4GL.DLL" (ByVal cd As Long, pprecvbuf As Long, precvlen As Long, ByVal flags As Long) As Long
Declare Function tpcancel Lib "TMAX4GL.DLL" (ByVal cd As Long) As Long
Declare Function tmaxreadenv Lib "TMAX4GL.DLL" (ByVal envfile As String, ByVal label As String) As Long

'/* ----- unsolicited messaging API ----- */
Declare Function tpsubscribe Lib "TMAX4GL.DLL" (ByVal eventexpr As String, ByVal filter As String, pctl As tpevctl_t, ByVal flags As Long) As Long
Declare Function tpunsubscribe Lib "TMAX4GL.DLL" (ByVal sd As Long, ByVal flags As Long) As Long
Declare Function tppost Lib "TMAX4GL.DLL" (ByVal eventname As String, ByVal pbuffer As Long, ByVal buflen As Long, ByVal flags As Long) As Long
Declare Function tpbroadcast Lib "TMAX4GL.DLL" (ByVal lnid As String, ByVal usrname As String, ByVal cltname As String, ByVal pbuffer As Long, ByVal buflen As Long, ByVal flags As Long) As Long
Declare Function tpsetunsol Lib "TMAX4GL.DLL" (ByVal pfunc As Long) As Long
'/* ----- tpchkunsol : by jaya (2004.07.30) ----- */
Declare Function tpchkunsol Lib "TMAX4GL.DLL" () As Long
Declare Function tpsetunsol_flag Lib "TMAX4GL.DLL" (ByVal flag As Long) As Long
Declare Function tpgetunsol Lib "TMAX4GL.DLL" (ByVal unsoltype As Long, ppbuffer As Long, precvlen As Long, ByVal flags As Long) As Long
    
'/* ----- conversational API ----- */
Declare Function tpsend Lib "TMAX4GL.DLL" (ByVal cd As Long, ByVal pbuffer As Long, ByVal buflen As Long, ByVal flags As Long, prevent As Long) As Long
Declare Function tprecv Lib "TMAX4GL.DLL" (ByVal cd As Long, ppbuffer As Long, pbuflen As Long, ByVal flags As Long, prevent As Long) As Long
Declare Function tpconnect Lib "TMAX4GL.DLL" (ByVal svcname As String, ByVal pbuffer As Long, ByVal buflen As Long, ByVal flags As Long) As Long
Declare Function tpdiscon Lib "TMAX4GL.DLL" (ByVal cd As Long) As Long
   
'/* ----- transaction API ----- */
Declare Function tx_begin Lib "TMAX4GL.DLL" () As Long
Declare Function tx_commit Lib "TMAX4GL.DLL" () As Long
Declare Function tx_rollback Lib "TMAX4GL.DLL" () As Long
Declare Function tx_set_transaction_timeout Lib "TMAX4GL.DLL" (ByVal timeout As Long) As Long

'/* reliable queue */
Declare Function tpenq Lib "TMAX4GL.DLL" (ByVal qname As String, ByVal svcname As String, ByVal pbuffer As Long, ByVal buflen As Long, ByVal flags As Long) As Long
Declare Function tpdeq Lib "TMAX4GL.DLL" (ByVal qname As String, ByVal svcname As String, ppbuffer As Long, pbuflen As Long, ByVal flags As Long) As Long
Declare Function tp_sleep Lib "TMAX4GL.DLL" (ByVal interval As Long)

'/* ----- etc API ------------- */
Declare Function tpstrerror Lib "TMAX4GL.DLL" (ByVal tperrno As Long) As Long

' /* ----- Useful buffer manipulation function ----- */
Declare Function vb_getstr Lib "TMAX4GL.DLL" (ByVal Fbfr As Long, uloc As Any) As Long
Declare Function vb_getcar Lib "TMAX4GL.DLL" (ByVal Fbfr As Long, uloc As Any, ByVal dlen As Long) As Long
Declare Function vb_putcar Lib "TMAX4GL.DLL" (ByVal Fbfr As Long, uloc As Any, ByVal dlen As Long) As Long

Declare Function vb_tpsleep Lib "TMAX4GL.DLL" (ByVal stime As Long) As Long

'/* ----- env API ------------ */
Declare Function tpgetenv Lib "TMAX4GL.DLL" (ByVal Name As String) As Long
Declare Function tpputenv Lib "TMAX4GL.DLL" (ByVal Name As String) As Long

