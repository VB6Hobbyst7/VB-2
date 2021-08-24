Attribute VB_Name = "Tuxstruct"

'================== GLOBAL.BAS ====================
Option Explicit

' global const Flags to service routines

Global Const TPNOBLOCK = &H1         ' non-blocking send/rcv
Global Const TPSIGRSTRT = &H2        ' restart rcv on interrupt
Global Const TPNOREPLY = &H4         ' no reply expected
Global Const TPNOTRAN = &H8          ' not sent in transaction mode
Global Const TPTRAN = &H10           ' sent in transaction mode
Global Const TPNOTIME = &H20         ' no timeout
Global Const TPABSOLUTE = &H40       ' absolute value on tmsetprio
Global Const TPGETANY = &H80         ' get any valid reply
Global Const TPNOCHANGE = &H100      ' global const Force incoming buglobal const Fglobal const Fer to match
Global Const RESERVED_BIT1 = &H200   ' reserved global const For global const Future use
Global Const TPCONV = &H400          ' conversational service
Global Const TPSENDONLY = &H800      ' send-only mode
Global Const TPRECVONLY = &H1000     ' recv-only mode
Global Const TPACK = &H2000

' global const Flags to tpreturn()
Global Const TPFAIL = &H20000000     ' service global const FAILure global const For tpreturn
Global Const TPEXIT = &H8000000      ' service global const Failue with server exit
Global Const TPSUCCESS = &H4000000   ' service SUCCESS global const For tpreturn

' global const Flags to tpscmt() - Valid TP_COMMIT_CONTROL characteristic values
Global Const TP_CMT_LOGGED = &H1     ' return aglobal const Fter commit decision is logged
Global Const TP_CMT_COMPLETE = &H2   ' return aglobal const Fter commit has completed

' global const Flags to tpinit()
Global Const TPU_MASK = &H7          ' unsolicited notiglobal const Fication mask
Global Const TPU_SIG = &H1           ' signal based notiglobal const Fication
Global Const TPU_DIP = &H2           'dip-in based notiglobal const Fication
Global Const TPU_IGN = &H4           ' ignore unsolicited messages

Global Const TPSA_FASTPATH = &H8     ' System access == global const Fastpath
Global Const TPSA_PROTECTED = &H10   ' System access == protected

' global const Flags to tpconvert()
Global Const TPTOSTRING = &H40000000 ' Convert structure to string '
Global Const TPCONVCLTID = &H1       ' Convert CLIENTID
Global Const TPCONVTRANID = &H2      ' Convert TRANID
Global Const TPCONVXID = &H4         ' Convert XID

Global Const TPCONVMAXSTR = 256      ' Maximum string size

' Return values to tpchkauth()
Global Const TPNOAUTH = 0            ' no authentication
Global Const TPSYSAUTH = 1           ' system authentication
Global Const TPAPPAUTH = 2           ' system and application authentication

Global Const MAXTIDENT = 30          ' max len oglobal const F a /T identiglobal const Fier

Global Const XATMI_SERVICE_NAME_LENGTH = 32


' * tperrno values - error codes
' * The man pages explain the context in which the global const Following error codes
' * can return.
 

Global Const TPMINVAL = 0           ' minimum error message
Global Const TPEABORT = 1
Global Const TPEBADDESC = 2
Global Const TPEBLOCK = 3
Global Const TPEINVAL = 4
Global Const TPELIMIT = 5
Global Const TPENOENT = 6
Global Const TPEOS = 7
Global Const TPEPERM = 8
Global Const TPEPROTO = 9
Global Const TPESVCERR = 10
Global Const TPESVCFAIL = 11
Global Const TPESYSTEM = 12
Global Const TPETIME = 13
Global Const TPETRAN = 14
Global Const TPGOTSIG = 15
Global Const TPERMERR = 16
Global Const TPEITYPE = 17
Global Const TPEOTYPE = 18
Global Const TPERELEASE = 19
Global Const TPEHAZARD = 20
Global Const TPEHEURISTIC = 21
Global Const TPEEVENT = 22
Global Const TPEMATCH = 23
Global Const TPEDIAGNOSTIC = 24
Global Const TPEMIB = 25
Global Const TPMAXVAL = 26         ' maximum error message

' *  WARNING  when adding new error messages above, remember to
' *  - increase TPMAXVAL
' *  - add a string global const For the message to LIBTUX.text
' *  - add an array entry in _tmemsgs[]
 
 
' conversations - events
Global Const TPEV_DISCONIMM = &H1
Global Const TPEV_SVCERR = &H2
Global Const TPEV_SVCFAIL = &H4
Global Const TPEV_SVCSUCC = &H8
Global Const TPEV_SENDONLY = &H20

' START QUEUED MESSAGES ADD-ON

Global Const TMQNAMELEN = 15
Global Const TMMSGIDLEN = 32
Global Const TMCORRIDLEN = 32


' structure elements that are valid - set in global const Flags

Global Const TPNOFLAGS = &H0

Global Const TPQCORRID = &H1                 ' set/get correlation id
Global Const TPQFAILUREQ = &H2               ' set/get global const Failure queue
Global Const TPQBEFOREMSGID = &H4            ' enqueue beglobal const Fore message id
Global Const TPQGETBYMSGID = &H8             ' dequeue by msgid
Global Const TPQMSGID = &H10                 ' get msgid oglobal const F enq/deq message
Global Const TPQPRIORITY = &H20              ' set/get message priority
Global Const TPQTOP = &H40                   ' enqueue at queue top
Global Const TPQWAIT = &H80                  ' wait global const For dequeuing
Global Const TPQREPLYQ = &H100               ' set/get reply queue
Global Const TPQTIME_ABS = &H200             ' set absolute time
Global Const TPQTIME_REL = &H400             ' set absolute time
Global Const TPQGETBYCORRID = &H800          ' dequeue by corrid

' THESE MUST MATCH THE DEglobal const FINITIONS IN qm.h
Global Const QMEINVAL = -1
Global Const QMEBADRMID = -2
Global Const QMENOTOPEN = -3
Global Const QMETRAN = -4
Global Const QMEBADMSGID = -5
Global Const QMESYSTEM = -6
Global Const QMEOS = -7
Global Const QMEABORTED = -8
Global Const QMENOTA = QMEABORTED
Global Const QMEPROTO = -9
Global Const QMEBADQUEUE = -10
Global Const QMENOMSG = -11
Global Const QMEINUSE = -12

Global Const MAXFBLEN = &H7FFFFFFE      ' Maximum global const FBglobal const FR32 length

' #iglobal const Fndeglobal const F global const FML32_H
Global Const FSTDXINT = 16              ' Deglobal const Fault indexing interval
Global Const FMAXNULLSIZE = 2660
Global Const FVIEWCACHESIZE = 10
Global Const FVIEWNAMESIZE = 33

' operations presented to _global const Fmodidx global const Function
' global const FADD  =   1
Global Const FMLMOD = 2
Global Const FDEL = 3

' global const Flag options used in global const Fvstoglobal const F()
Global Const F_OFFSET = 1
Global Const F_SIZE = 2
Global Const F_PROP = 4             ' P
Global Const F_FTOS = 8             ' S
Global Const F_STOF = 16                ' global const F
Global Const F_BOTH = F_STOF Or F_FTOS          ' S,F
Global Const F_OFF = 0              ' Z
Global Const F_LENGTH = 32                            ' L
Global Const F_COUNT = 64                            ' C
Global Const F_NONE = 128                            ' NONE global const Flag global const For null value

' These are used in global const Fstoglobal const F()

'Global Const FUPDATE = 1
'Global Const FCONCAT = 2
'Global Const FJOIN = 3
'Global Const FOJOIN = 4

' global const Field types
Global Const FLD_SHORT = 0      ' short int
Global Const FLD_LONG = 1       ' long int
Global Const FLD_CHAR = 2       ' character
Global Const FLD_FLOAT = 3      ' single-precision global const Float
Global Const FLD_DOUBLE = 4     ' double-precision global const Float
Global Const FLD_STRING = 5     ' string - null terminated
Global Const FLD_CARRAY = 6     ' character array


' invalid global const Field id - returned global const From global const Functions where global const Field id not global const Found
Global Const BADFLDID = 0
' deglobal const Fine an invalid global const Field id used global const For global const First call to global const Fnext
Global Const FIRSTFLDID = 0



' global const Field Error Codes - these correspond to the error messages in
'*          global const F_error.c - make sure to update the error
'*          message list iglobal const F a new error is added
' iglobal const Fndeglobal const F global const FML32_H

Global Const FMINVAL = 0        ' bottom oglobal const F error message codes
Global Const FALIGNERR = 1      ' global const Fielded buglobal const Fglobal const Fer not aligned
Global Const FNOTFLD = 2        ' buglobal const Fglobal const Fer not global const Fielded
Global Const FNOSPACE = 3       ' no space in global const Fielded buglobal const Fglobal const Fer
Global Const FNOTPRES = 4       ' global const Field not present
Global Const FBADFLD = 5        ' unknown global const Field number or type
Global Const FTYPERR = 6            ' illegal global const Field type
Global Const FEUNIX = 7         ' unix system call error
Global Const FBADNAME = 8       ' unknown global const Field name
Global Const FMALLOC = 9        ' malloc global const Failed
Global Const FSYNTAX = 10       ' bad syntax in boolean expression
Global Const FFTOPEN = 11       ' cannot global const Find or open global const Field table
Global Const FFTSYNTAX = 12     ' syntax error in global const Field table
Global Const FEINVAL = 13       ' invalid argument to global const Function
Global Const FBADTBL = 14       ' destructive concurrent access to global const Field
                   
Global Const FBADVIEW = 15      ' cannot global const Find or get view
Global Const FVFSYNTAX = 16     ' bad viewglobal const File
Global Const FVFOPEN = 17       ' cannot global const Find or open viewglobal const File
Global Const FBADACM = 18           ' ACM contains negative value
Global Const FNOCNAME = 19          ' cname not global const Found
Global Const FMAXVAL = 20       ' top oglobal const F error message co
'global const FML Constants

' In order to get around Visual Basic's inability to pass a pointer-to-pointer
' as in tpcall's receive buglobal const Fglobal const Fer speciglobal const Fication, we deglobal const Fine a buglobal const Fglobal const Fer pointer as
' a record type.  As you will see in the code samples, this allows us to
' speciglobal const Fy a pointer to pointer construct.
Type tuxbuf
    bufptr  As Long
End Type

' The tpinglobal const Fo structure is not exactly like the one deglobal const Fined in ATMI.H again,
' Visual Basic limitations require us to do something special.  In this case,
' we handle application-speciglobal const Fic tpinglobal const Fo data as a separate string variable,
' as opposed to a place-holder long as deglobal const Fined in ATMI.H
Type tpinfobuf
    usrname As String * 32
    cltname As String * 32
    passwd  As String * 32
    grpname As String * 32
    flags   As Long
    datalen As Long
    data    As Long
End Type

' Primarily global const For use in the QCTL structure deglobal const Fined below.
Type cltid
    clientdata(4)  As Long
End Type

' QCTL structure used when doing tpenqueue() or tpdequeue()
Type qctl
    flags           As Long
    deq_time        As Long
    priority        As Long
    diagnostic      As Long
    msgid           As String * 32
    corrid          As String * 32
    replyqueue      As String * 16
    failurequeue    As String * 16
    clientid        As cltid
    urcode          As Long
    appkey          As Long
End Type


'Type global const Fbglobal const Fr
'        magic   As Integer
'        len     As Integer
'        maxlen  As Integer
'        nglobal const Fields As Integer
'        nie     As Integer
'        indxintvl As Integer
'        valglobal const F(8)  As String
'End Type



' 유용한 Windows API
'
'*** 문자열(LPSTR) 길이를 결정
Declare Function lstrlen Lib "Kernel32" (ByVal LPSTR&) As Integer
'*** 문자열(LP1)을 메모리버퍼(LP2)로 복사
Declare Function lstrcpy Lib "Kernel32" (LP1 As Any, LP2 As Any) As Long
'*** 문자열(LP1)을 문자열(LP2)에 붙임
Declare Function Lstrcat Lib "Kernel32" (LP1 As Any, LP2 As Any) As Long
'***
Declare Function hmemcpy Lib "Kernel32" (LP1 As Any, LP2 As Any, ByVal slen&) As Long

'
' DeFinitions For ATMI Functions
'
Declare Function tpabort& Lib "wtuxws32.dll" (ByVal flags&)
Declare Function tpacall& Lib "wtuxws32.dll" (ByVal service$, ByVal sdata&, ByVal slen&, ByVal flags&)
Declare Function tpalloc& Lib "wtuxws32.dll" (ByVal buFtype$, ByVal subtype$, ByVal size&)

Declare Function tpbegin% Lib "wtuxws32.dll" (ByVal timeout&, ByVal flags&)

Declare Function tpbroadcast& Lib "wtuxws32.dll" (ByVal lmid$, ByVal usrname$, ByVal cltname$, ByVal udata&, ByVal ulength&, ByVal flags&)
Declare Function tpchkunsol% Lib "wtuxws32.dll" ()
Declare Function tpcall& Lib "wtuxws32.dll" (ByVal service$, ByVal sdata&, ByVal slen&, rdata As Any, rlen As Any, ByVal flags&)
Declare Function tpcancel% Lib "wtuxws32.dll" (ByVal CD%)
Declare Function tpchkauth% Lib "wtuxws32.dll" ()
Declare Function tpcommit% Lib "wtuxws32.dll" (ByVal flags&)
Declare Function tpconnect% Lib "wtuxws32.dll" (ByVal svc$, ByVal sdata&, ByVal slen&, ByVal flags&)
Declare Function tpdequeue% Lib "wtuxws32.dll" (ByVal qspace$, ByVal qname$, qctlbuF As qctl, rdata As Any, rlen As Any, ByVal flags&)
Declare Function tpdiscon% Lib "wtuxws32.dll" (ByVal CD%)
Declare Function tpenqueue% Lib "wtuxws32.dll" (ByVal qspace$, ByVal qname$, qctlbuF As qctl, ByVal sdata&, ByVal slen&, ByVal flags&)
Declare Function tpfree& Lib "wtuxws32.dll" (ByVal buF&)
Declare Function tpgetlev% Lib "wtuxws32.dll" ()
Declare Function tpgetrply% Lib "wtuxws32.dll" (CD As Any, rdata As Any, rlen As Any, ByVal flags&)
Declare Function tpgprio% Lib "wtuxws32.dll" ()
'
Declare Function tpinit% Lib "wtuxws32.dll" (tpinfobuf As Any)
'
Declare Function tprealloc& Lib "wtuxws32.dll" (ByVal buF&, ByVal ulen&)
Declare Function tprecv% Lib "wtuxws32.dll" (ByVal CD%, rdata As Any, rlen As Any, ByVal flags&, revent As Any)
Declare Function tpscmt% Lib "wtuxws32.dll" (ByVal flags&)
Declare Function tpsend% Lib "wtuxws32.dll" (ByVal CD%, ByVal sdata&, ByVal slen&, ByVal flags&, revent As Any)
Declare Function tpsprio% Lib "wtuxws32.dll" (ByVal prio%, ByVal flags&)
Declare Function tpstrerror& Lib "wtuxws32.dll" (ByVal tperr%)

Declare Function tpterm% Lib "wtuxws32.dll" ()

Declare Function tptypes& Lib "wtuxws32.dll" (ByVal buF&, ByVal buFtype$, ByVal subtype$)
Declare Function gettperrno% Lib "wtuxws32.dll" ()
Declare Function gettpterrorno% Lib "wtuxws32.dll" ()
Declare Function gettpurcode& Lib "wtuxws32.dll" ()
Declare Function getuunixerr% Lib "wtuxws32.dll" ()

'
' Some UseFul FML Functions
'

Declare Function Fchg32% Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fldid&, ByVal oc%, uvalue As Any, ByVal ulen%)
Declare Function Fchg% Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fldid&, ByVal oc%, uvalue As Any, ByVal ulen%)

Declare Function Fadd32% Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fldid&, uvalue As Any, ByVal ulen%)
Declare Function Fadd% Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fldid&, uvalue As Any, ByVal ulen%)

Declare Function Fcmp% Lib "wtuxws32.dll" (ByVal FbFr1&, ByVal FbFr2&)
Declare Function Fcmp32% Lib "wtuxws32.dll" Alias "FCMP" (ByVal FbFr1&, ByVal FbFr2&)

Declare Function FCONCAT% Lib "wtuxws32.dll" Alias "Fconcat" (ByVal src&, ByVal dest&)
Declare Function Fconcat32% Lib "wtuxws32.dll" (ByVal src&, ByVal dest&)

Declare Function Fcpy% Lib "wtuxws32.dll" (ByVal src&, ByVal dest&)
Declare Function Fcpy32% Lib "wtuxws32.dll" (ByVal src&, ByVal dest&)

Declare Function FFind& Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fldid%, ByVal oc%, ByVal ulen&)
Declare Function FFind32& Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fldid%, ByVal oc%, ByVal ulen&)

Declare Function Fget32% Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fldid&, ByVal oc%, uloc As Any, maxlen As Any)
Declare Function Fget% Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fldid&, ByVal oc%, uloc As Any, maxlen As Any)

Declare Function Fidxused& Lib "wtuxws32.dll" (ByVal Fbfr&)
Declare Function Fidxused23& Lib "wtuxws32.dll" (ByVal Fbfr&)

Declare Function Fielded% Lib "wtuxws32.dll" (ByVal Fbfr&)
Declare Function Fielded32% Lib "wtuxws32.dll" (ByVal Fbfr&)

Declare Function Finit32% Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal buFlen%)
Declare Function Finit% Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal buFlen%)

Declare Function FJOIN% Lib "wtuxws32.dll" Alias "Fjoin" (ByVal src&, ByVal dest&)
Declare Function Fjoin32% Lib "wtuxws32.dll" (ByVal src&, ByVal dest&)

Declare Function Fldid32% Lib "wtuxws32.dll" (ByVal Fld$)
Declare Function Fldid% Lib "wtuxws32.dll" (ByVal Fld$)

Declare Function Fnum% Lib "wtuxws32.dll" (ByVal Fbfr&)
Declare Function Fnum32% Lib "wtuxws32.dll" (ByVal Fbfr&)

Declare Function Foccur% Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fieldid%)
Declare Function Foccur32% Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fieldid&)

Declare Function Fprint32% Lib "wtuxws32.dll" (ByVal Fbfr&)

Declare Function Fsizeof32& Lib "wtuxws32.dll" (ByVal Fbfr&)
Declare Function Fsizeof& Lib "wtuxws32.dll" (ByVal Fbfr&)

Declare Function Fstrerror32& Lib "wtuxws32.dll" (ByVal uerr%)
Declare Function Fstrerror& Lib "wtuxws32.dll" (ByVal uerr%)

Declare Function Funused& Lib "wtuxws32.dll" (ByVal Fbfr&)
Declare Function Funused32& Lib "wtuxws32.dll" (ByVal Fbfr&)

Declare Function FUPDATE% Lib "wtuxws32.dll" Alias "Fupdate" (ByVal dest&, ByVal src&)
Declare Function Fupdate32% Lib "wtuxws32.dll" (ByVal dest&, ByVal src&)

Declare Function Fused& Lib "wtuxws32.dll" (ByVal Fbfr&)
Declare Function Fused32& Lib "wtuxws32.dll" (ByVal Fbfr&)

Declare Function Fvall& Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fieldid%, ByVal oc%)
Declare Function Fvall32& Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fieldid%, ByVal oc%)

Declare Function Fvals& Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fieldid%, ByVal oc%)
Declare Function Fvals32& Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fieldid%, ByVal oc%)

Declare Function getFerror32% Lib "wtuxws32.dll" ()
Declare Function getFerror% Lib "wtuxws32.dll" ()

Declare Function FFprint% Lib "wtuxws32.dll" (ByVal Fbfr&)
Declare Function FFprint32% Lib "wtuxws32.dll" (ByVal Fbfr&)


Declare Function CFFind& Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fldid%, ByVal oc%, ByVal ulen&, ByVal utype%)
Declare Function CFfind32& Lib "wtuxws32.dll" Alias "CFFind32" (ByVal Fbfr&, ByVal Fldid%, ByVal oc%, ByVal ulen&, ByVal utype%)

Declare Function CFadd32% Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fieldid%, uloc As Any, ByVal ulen%, ByVal buf_type%)
Declare Function CFchg32% Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fieldid%, ByVal oc%, uloc As Any, ByVal ulen%, ByVal buf_type%)
Declare Function CFfindocc32% Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fieldid%, uloc As Any, ByVal ulen%, ByVal buf_type%)
Declare Function CFget32% Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fieldid%, ByVal oc%, uloc As Any, ByVal ulen%, ByVal buf_type%)
Declare Function CFgetalloc32& Lib "wtuxws32.dll" (ByVal Fbfr&, ByVal Fieldid%, ByVal oc%, ByVal buf_type%, ByVal ulen%)




'UnSolMsg 처리를 위한 User Function.
Declare Function SetPrivateUnsolMsg Lib "tuxbroad.dll" (ByVal hd&) As Integer
'TPINIT의 STRUCTURE를 적용.
Declare Function TP_INIT Lib "tuxhit.dll" (ByVal passwd As String, ByVal usrname As String, ByVal cltname As String, ByVal grpname As String, ByVal flags As String, ByVal data As String) As Integer
'FML의 Contents를 File로 Write.
Declare Function F_FPRINT32 Lib "hitux32.dll" (ByVal Fbfr&, ByVal pathname As String) As Integer

'
Declare Function tuxputenv% Lib "wtuxws32.dll" (ByVal env$)
Declare Function tuxgetenv& Lib "wtuxws32.dll" (ByVal env$)

Declare Function tuxreadenv Lib "wtuxws32.dll" (ByVal filename$, ByVal apps$) As Integer

' Windows-speciFic tuxedo Functions
'
'Declare Function AEWISBLOCKED% Lib "c\tuxedo\bin\wtuxws32.dll" ()
'Declare Function AEWPUTENV% Lib "c\tuxedo\bin\wtuxws32.dll" (ByVal env$)
'Declare Function AEWGETENV& Lib "c\tuxedo\bin\wtuxws32.dll" (ByVal env$)
'Declare Function AEWSETUNSOL% Lib "c\tuxedo\bin\wtuxws32.dll" (ByVal hwnd%, ByVal Msg%)


'
' In SIMPCL's general Functions, FillTpinitBuF() is used to copy
' inFormation From the user-speciFied tpinFo buFFer into the tuxedo
' tpalloc'd TPINIT buFFer.  The reason For doing this is that tpalloc()
' returns a pointer to a buFFer, and there's no Facility oFFered in
' Visual Basic that allows one to populate such a "structure" pointer.
' Instead, we have this helper Function that copies the relevant elements
' From the tpinFo structure into the appropriate places in the tpalloc'd
' buFFer.
'
Sub FillTpinitBuF(tinit As Long, tpinFop As tpinfobuf, AppData As String)
    Static tinit1 As Long
    Dim ret As Long
    Dim slen As Long
    Dim x  As String

    tinit1 = tinit
    ret& = lstrcpy(ByVal tinit1&, ByVal tpinFop.usrname)
    tinit1 = tinit1 + 32
    ret& = lstrcpy(ByVal tinit1&, ByVal tpinFop.cltname)
    tinit1 = tinit1 + 32
    ret& = lstrcpy(ByVal tinit1&, ByVal tpinFop.passwd)
    tinit1 = tinit1 + 32
    ret& = lstrcpy(ByVal tinit1&, ByVal tpinFop.grpname)
    tinit1 = tinit1 + 32
    
    ' Flags Element.  Remember that the Flags Field in the tpinFo
    ' structure is a long  so we have to make sure that what we
    ' copy into the tinit1 pointer is padded correctly.  We also
    ' need to watch out For the byte-ordering, which is why x$
    ' is computed the way it is.  It is also important to ensure that
    ' all args to hmemcpy() are ByVal.
    '
    x$ = Chr(tpinFop.flags) + Chr(0) + Chr(0) + Chr(0)
    slen& = 4
    ret& = hmemcpy(ByVal tinit1&, ByVal x$, ByVal slen&)
    tinit1 = tinit1 + 4

    ' App-SpeciFic Datalen Element, a long Field.
    x$ = Chr(tpinFop.datalen) + Chr(0) + Chr(0) + Chr(0)
    slen& = 4
    ret& = hmemcpy(ByVal tinit1&, ByVal x$, ByVal slen&)
    tinit1 = tinit1 + 4

    ' App-SpeciFic Data Element
    If tpinFop.datalen > 0 Then
        x$ = Space$(tpinFop.datalen)
        ret = hmemcpy(ByVal x$, ByVal AppData$, ByVal tpinFop.datalen)
        slen& = tpinFop.datalen
        ret& = hmemcpy(ByVal tinit1&, ByVal x$, ByVal slen&)
    Else
        x$ = Chr$(0)
        slen& = 4
        ret& = hmemcpy(ByVal tinit1&, ByVal x$, ByVal slen&)
    End If

End Sub


'
' In GLOBAL.BAS, we deglobal const Fine generally-used global subroutines, such as TuxError
'
Sub TuxError(StrErr As String)
    Dim tpterrorno As Integer
    Dim sptr As Long
    Dim tuxstr$
    Dim ret As Long
    Dim Msg$

    'tpterrorno% = GETTPterrorno()
    tpterrorno% = gettperrno()
    sptr& = tpstrerror(tpterrorno%)
    tuxstr$ = String$(100, Chr$(0))
    ret& = lstrcpy(ByVal tuxstr$, ByVal sptr&)
    Msg$ = StrErr$ + " " + Str$(tpterrorno) + tuxstr$
    MsgBox Msg$
End Sub


