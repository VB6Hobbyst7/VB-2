Attribute VB_Name = "odeconst"
'/* COPYRIGHT_TEXT */


'*********************************************************************
'
'                      OEC LITE CONSTANTS for Visual Basic
'
'********************************************************************


' ************* Begin OEC constants (do not remove) **************************


'  0

Global Const DCE_NOERROR = 0
Global Const DCE_NOENVSET = 1
Global Const DCE_NOPORT = 2
Global Const DCE_NOHOST = 3
Global Const DCE_NOBROKER = 4
Global Const DCE_NOSERVER = 5
Global Const DCE_BADHOST = 6
Global Const DCE_BADSERVHOST = 7
Global Const DCE_NOSOCKCREATE = 8
Global Const DCE_NOSOCKBIND = 9


'   10
Global Const DCE_NOSOCKCONN = 10
Global Const DCE_NOSOCKACCEPT = 11
Global Const DCE_NOSOCKTMOUT = 12
Global Const DCE_NOSUCHFUNC = 13
Global Const DCE_LOCALHOSTUNKN = 14
Global Const DCE_NOSOCKLISTEN = 15
Global Const DCE_CLIENTBUSY = 16
Global Const DCE_NOMEMORY = 17
Global Const DCE_BADLOGIN = 18
Global Const DCE_CANTOPENLOG = 19


'  20


Global Const DCE_BADSYNTAX = 20
Global Const DCE_CANTOPENENV = 21
Global Const DCE_NULLENVFILE = 22
Global Const DCE_ENVALREADY = 23
Global Const DCE_PEERERROR = 24
Global Const DCE_LOSTSERVER = 25
Global Const DCE_BADTCPINIT = 26
Global Const DCE_IPCINITBAD = 27
Global Const DCE_IPCCANTCLOSE = 28
Global Const DCE_READFAIL = 29


'   30


Global Const DCE_WRITEFAIL = 30
Global Const DCE_CANTFORK = 31
Global Const DCE_BADPORT = 32
Global Const DCE_NOINTERFACE = 33
Global Const DCE_BADARG = 34
Global Const DCE_BADTICKET = 35
Global Const DCE_SIGINT = 36
Global Const DCE_SERVERFAILED = 37
Global Const DCE_BADINTERFACE = 38
Global Const DCE_NOAUTH = 39


'  40

Global Const DCE_SECREQ = 40
Global Const DCE_NOSEC = 41
Global Const DCE_WRITABLE = 42
Global Const DCE_MAXCAPACITY = 43
Global Const DCE_NOSUBB = 44
Global Const DCE_FSERROR = 45
Global Const DCE_RPCTIMEOUT = 46
Global Const DCE_BADVERSION = 47
Global Const DCE_NODYNINTERFACE = 48
Global Const DCE_NONAMECHANGE = 49


'   50


Global Const DCE_NOENVRESET = 50
Global Const DCE_NOTIMERAVAIL = 51
Global Const DCE_LASTERR = 52


'********************  OEC logging constants **************************

Global Const DCE_LOG_NONE = 0
Global Const DCE_LOG_ERROR = 6
Global Const DCE_LOG_WARNING = 19
Global Const DCE_LOG_DEBUG = 29


'********************  OEC CONSTANTS for window event *****************

Global Const DCE_NOIGNORE = 0
Global Const DCE_IGNORE = 1
Global Const DCE_QUERYSTATUS = 2


'********************  OEC CONSTANTS for DATABASES ********************

Global Const ORA_DCP = 1
Global Const IFX_DCP = 2
Global Const SYB_DCP = 3
Global Const ING_DCP = 4
Global Const DB2_DCP = 5




'This is the return value of a function when an error Occurs

Global Const DCPERROR = 0



'This should be the return value of a function when it's successful

Global Const DCPSUCCESS = 1




'DBMS ERRORS
'-----------

' This should be the default setting  for the error parameter when the
' result is successful.

Global Const DBNOEEROR = 0


'This is typically flagged by the DBMS. This Occurs when a login to the DBMS
'is refused. In case it isn't caught, use this.


Global Const DBCONNREFUSED = -1000000


' This is typically flagged by the DBMS. This will happen when the connection
'to the DBMS is lost. In case it isn't caught, use this.

Global Const DBCONNDOWN = -1000005

'OEC ERRORS
'----------


'OEC Error -- Commit Pending.  This Occurs when the a 2nd begin work
'is issued without committing (or aborting) the previous transaction.


Global Const DBCOMMITPEND = -1000010



'OEC Error -- No Transaction . This Occurs when a commit/abort of a
'transaction is attempted without properly beginning it.


Global Const DBNOTRANSAC = -1000015



'OEC Error -- No Memory.   Occurs when memory allocation fails.
'The server logs should give more detailed info.


Global Const DBNOMEMORY = -1000020


' OEC Error -- Already Connected.  Occurs when you are already connected
' to the DBMS and try a 2nd time (e.g. sql_prepare() RPC called twice
' to a dedicated server).


Global Const DBALREADYCONN = -1000025



'OEC Error -- Use Login RPC.  Occurs when you failed to login with
'the login rpc (sql_prepare()), then you issued a 2nd, different RPC. You must
'1st issue sql_prepare() successfully.


Global Const DBUSELOGINRPC = -1000030



'OEC Error -- No Cursors Allowed.  Occurs when you attempt to use
'cursors (sql_rows() RPC) with  a non-dedicated server.

Global Const DBNOCURSORS = -1000035



'OEC Error -- Query Not Found.  Occurs when a server receives an SQL
'command (RPC) that it cannot handle i.e. knows nothing about. Possibly
'will happen with version skew.

Global Const DBQUERYNOTFOUND = -1000040
                                        


'OEC Error -- DB Last Record. The last record of the query was
'retrieved. This is set so that libdcp.a can manage cursors appropriately.

Global Const DBLASTREC = -1000045



'OEC  Error -- Log File Error.  An error was encountered while trying
'to log a transaction. Most likely a memory allocation error. Check the
'Server Debug Log.

Global Const DBLOGERROR = -1000050



'OEC Error -- Invalid Parameter.  An invalid parameter was passed with
'the function. Check the valid value ranges of the inputs. e.g. sql_rows()
'requires an input >= 0. A negative value would result in this error.

Global Const DBINVALIDPARAM = -1000055


'OEC Error -- Miscellaneous Error.  Check the Server Log for more info.


Global Const DBMISC = -1000100


