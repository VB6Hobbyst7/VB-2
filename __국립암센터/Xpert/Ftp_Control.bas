Attribute VB_Name = "Ftp_Control"
'Option Explicit
'
''***INTERNET_FLAG_PASSIVE Mode ***'
'
'Private Const INTERNET_FLAG_PASSIVE_YES = &H8000000
'Private Const INTERNET_FLAG_PASSIVE_NO = 0
'
''*** 화일 복사, 삭제, 이동 라이브러리 ***'
'
'Private Const FO_COPY = &H2&
'Private Const FO_delete = &H3&
'Private Const FO_MOVE = &H1&
'Private Const FO_RENAME = &H4&
'Private Const FOF_ALLOWUNDO = &H40&
'Private Const FOF_CONFIRMMOUSE = &H2&
'Private Const FOF_CREATEPROGRESSDLG = &H0&
'Private Const FOF_FILESONLY = &H80&
'Private Const FOF_MULTIDESTFILES = &H1&
'Private Const FOF_NOCONFIRMATION = &H10&
'Private Const FOF_NOCONFIRMMKDIR = &H200&
'Private Const FOF_RENAMEONCOLLISION = &H8&
'Private Const FOF_SILENT = &H4&
'Private Const FOF_SIMPLEPROGRESS = &H100&
'Private Const FOF_WANTMAPPINGHandLE = &H20&
'
'Private Type SHFILEOPSTRUCT
'    hwnd As Long
'    wFunc As Long
'    pFrom As String
'    pTo As String
'    fFlags As Integer
'    fAnyOperationsAborted As Long
'    hNameMappings As Long
'    lpszProgressTitle As String
'End Type
'
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Any) As Long
'
''화일 OPEN 라이브러리
'
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
'    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
'    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'
'Private Declare Function GetDesktopWindow Lib "user32" () As Long
'
'Const SW_SHOWNORMAL = 1
'
''FTP
'
'Declare Function GetProcessHeap Lib "kernel32" () As Long
'Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
'Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
'Public Const HEAP_ZERO_MEMORY = &H8
'Public Const HEAP_GENERATE_EXCEPTIONS = &H4
'
'Declare Sub CopyMemory1 Lib "kernel32" Alias "RtlMoveMemory" ( _
'         hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
'
'Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" ( _
'         hpvDest As Long, hpvSource As Any, ByVal cbCopy As Long)
'
'Public Const MAX_PATH = 260
'Public Const NO_ERROR = 0
'Public Const FILE_ATTRIBUTE_READONLY = &H1
'Public Const FILE_ATTRIBUTE_HIDDEN = &H2
'Public Const FILE_ATTRIBUTE_SYSTEM = &H4
'Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
'Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
'Public Const FILE_ATTRIBUTE_NORMAL = &H80
'Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
'Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
'Public Const FILE_ATTRIBUTE_OFFLINE = &H1000
'
'Type FILETIME
'    dwLowDateTime As Long
'    dwHighDateTime As Long
'End Type
'
'Type WIN32_FIND_DATA
'        dwFileAttributes As Long
'        ftCreationTime As FILETIME
'        ftLastAccessTime As FILETIME
'        ftLastWriteTime As FILETIME
'        nFileSizeHigh As Long
'        nFileSizeLow As Long
'        dwReserved0 As Long
'        dwReserved1 As Long
'        cFileName As String * MAX_PATH
'        cAlternate As String * 14
'End Type
'
'Public Const ERROR_NO_MORE_FILES = 18
'
'Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
'(ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
'
'Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
'(ByVal hFtpSession As Long, ByVal lpszSearchFile As String, _
'      lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
'
'Public Declare Function FTPGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
'(ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
'      ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, _
'      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
'
'Public Declare Function FTPPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
'(ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
'      ByVal lpszRemoteFile As String, _
'      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
'
'Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
'    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
'
'' *** Initializes an application's use of the Win32 Internet functions ***'
'
'Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
'(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
'ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
'
''*** User agent constant. ***'
'
''Public Const scUserAgent = "vb wininet"
'Public Const scUserAgent = "science"
'
''*** Use registry access settings. ***'
'
'Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
'Public Const INTERNET_OPEN_TYPE_DIRECT = 1
'Public Const INTERNET_OPEN_TYPE_PROXY = 3
'Public Const INTERNET_INVALID_PORT_NUMBER = 0
'
'Public Const FTP_TRANSFER_TYPE_ASCII = &H1
'Public Const FTP_TRANSFER_TYPE_BINARY = &H2
'Public Const INTERNET_FLAG_PASSIVE = &H8000000
'
''*** Opens a HTTP session for a given site. ***'
'
'Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'
'Public Const ERROR_INTERNET_EXTENDED_ERROR = 12003
'Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean
'
''*** Number of the TCP/IP port on the server to connect to. ***'
'
'Public Const INTERNET_DEFAULT_FTP_PORT = 21
'Public Const INTERNET_DEFAULT_GOPHER_PORT = 70
'Public Const INTERNET_DEFAULT_HTTP_PORT = 80
'Public Const INTERNET_DEFAULT_HTTPS_PORT = 443
'Public Const INTERNET_DEFAULT_SOCKS_PORT = 1080
'
'Public Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
'Public Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
'Public Const INTERNET_OPTION_SEND_TIMEOUT = 5
'
'Public Const INTERNET_OPTION_USERNAME = 28
'Public Const INTERNET_OPTION_PASSWORD = 29
'Public Const INTERNET_OPTION_PROXY_USERNAME = 43
'Public Const INTERNET_OPTION_PROXY_PASSWORD = 44
'
'' Type of service to access.
'
'Public Const INTERNET_SERVICE_FTP = 1
'Public Const INTERNET_SERVICE_GOPHER = 2
'Public Const INTERNET_SERVICE_HTTP = 3
'
'' Opens an HTTP request handle.
'
'Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
'(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
'ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'
'' Brings the data across the wire even if it locally cached.
'
'Public Const INTERNET_FLAG_RELOAD = &H80000000
'Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
'Public Const INTERNET_FLAG_MULTIPART = &H200000
'Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000     ' don't write this item to the cache
'
'Public Const GENERIC_READ = &H80000000
'Public Const GENERIC_WRITE = &H40000000
'
'' Sends the specified request to the HTTP server.
'
'Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal _
'hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As _
'String, ByVal lOptionalLength As Long) As Integer
'
'' Queries for information about an HTTP request.
'
'Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
'(ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, _
'ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
'
'' The possible values for the lInfoLevel parameter include:
'
'Public Const HTTP_QUERY_CONTENT_TYPE = 1
'Public Const HTTP_QUERY_CONTENT_LENGTH = 5
'Public Const HTTP_QUERY_EXPIRES = 10
'Public Const HTTP_QUERY_LAST_MODIFIED = 11
'Public Const HTTP_QUERY_PRAGMA = 17
'Public Const HTTP_QUERY_VERSION = 18
'Public Const HTTP_QUERY_STATUS_CODE = 19
'Public Const HTTP_QUERY_STATUS_TEXT = 20
'Public Const HTTP_QUERY_RAW_HEADERS = 21
'Public Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
'Public Const HTTP_QUERY_FORWARDED = 30
'Public Const HTTP_QUERY_SERVER = 37
'Public Const HTTP_QUERY_USER_AGENT = 39
'Public Const HTTP_QUERY_set_COOKIE = 43
'Public Const HTTP_QUERY_REQUEST_METHOD = 45
'Public Const HTTP_STATUS_DENIED = 401
'Public Const HTTP_STATUS_PROXY_AUTH_REQ = 407
'
''*** Add this flag to the about flags to get request header. ***'
'
'Public Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
'Public Const HTTP_QUERY_FLAG_NUMBER = &H20000000
'
''*** Reads data from a handle opened by the HttpOpenRequest function. ***'
'
'Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
'Public Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumberOfBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
'Public Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal sFilename As String, ByVal lAccess As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'Public Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
'Public Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Long
'Public Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Long
'Public Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long
'Public Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer
'Public Declare Function InternetSetOptionStr Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hInternet As Long, ByVal lOption As Long, ByVal sBuffer As String, ByVal lBufferLength As Long) As Integer
'
''*** Closes a single Internet handle or a subtree of Internet handles. ***'
'Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
'
''*** Queries an Internet option on the specified handle ***'
'Public Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" _
'(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long) As Integer
'
''*** Returns the version number of Wininet.dll. ***'
'Public Const INTERNET_OPTION_VERSION = 40
'
''*** Contains the version number of the DLL that contains the Windows Internet
'
''    functions (Wininet.dll). This structure is used when passing the
'
''    INTERNET_OPTION_VERSION flag to the InternetQueryOption function. ***'
'
'Public Type tWinInetDLLVersion
'    lMajorVersion As Long
'    lMinorVersion As Long
'End Type
'
''*** Adds one or more HTTP request headers to the HTTP request handle. ***'
'Public Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" _
'                                          (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, _
'                                           ByVal lModifiers As Long) As Integer
'
''**** Flags to modify the semantics of this function. Can be a combination of these values:***'
'Public Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000 '<--Adds the header only if it does not already exist; otherwise, an error is returned.
'Public Const HTTP_ADDREQ_FLAG_ADD = &H20000000 '<--Adds the header if it does not exist. Used with REPLACE.
'Public Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000 '<-- Replaces or removes a header. If the header value is empty and the header is found,
'
'                                                   '    it is removed. If not empty, the header value is replaced
'
'Public hOpen As Long
'Public hConnection As Long
'Public Const dwType = FTP_TRANSFER_TYPE_BINARY
'
''====================================
'' FTP_Connection 함수
''------------------------------------
'' FTP서버에 연결..
''====================================
'
'Public Function FTP_Connection() As Boolean
'Dim mmFTP As String
'Dim mmFID As String
'Dim mmFPASS As String
'Dim mmPORT As Long
'
'mmFTP = ftp_server
'mmFID = ftp_id
'mmFPASS = ftp_pass
'mmPORT = ftp_port
'   FTP_Connection = True
'   hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
'
'   If hOpen = 0 Then
'      'MsgBox "인터넷 연결에 실패하였습니다.", vbExclamation
'      FTP_Connection = False
'      Exit Function
'   End If
'   If ftp_passive_mode = "Y" Then
'      hConnection = InternetConnect(hOpen, mmFTP, mmPORT, mmFID, mmFPASS, INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE_YES, 0)
'   Else
'      hConnection = InternetConnect(hOpen, mmFTP, mmPORT, mmFID, mmFPASS, INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE_NO, 0)
'   End If
'
'   If hConnection = 0 Then
'      'MsgBox "인터넷 연결에 실패하였습니다.", vbExclamation
'      FTP_Connection = False
'      Exit Function
'   End If
'End Function
'
''====================================
'' FTP_DisConnect 함수
''------------------------------------
'' FTP서버로부터 연결해제..
''====================================
'Public Function FTP_DisConnect() As Boolean
'
'    FTP_DisConnect = True
'    If hConnection <> 0 Then
'        InternetCloseHandle (hConnection)
'    End If
'    If hOpen <> 0 Then
'        InternetCloseHandle (hOpen)
'    End If
'End Function
'
''====================================
'' FTP_Download 함수
''------------------------------------
'' FTP서버로부터 파일받기..
''====================================
'
'Public Function FTP_Download(ServerName As String, ClientName As String) As Boolean
'
'   Dim bRet As Boolean
'   FTP_Download = True
'   bRet = FTPGetFile(hConnection, ServerName, ClientName, False, INTERNET_FLAG_RELOAD, dwType, 0)
'   If bRet = False Then
'      FTP_Download = False
'   End If
'End Function
'
''====================================
'' FTP_Upload 함수
''------------------------------------
'' FTP서버로 파일보내기..
''====================================
'Public Function FTP_Upload(ClientName As String, ServerDir As String, Filename As String) As Boolean
'
'   Dim bRet As Boolean
'   FTP_Upload = True
'   If hOpen = 0 Or hConnection = 0 Then
'      MsgBox "현재 FTP서버에 연결이 끊겼습니다.", vbInformation, App.Title
'      FTP_Upload = False
'   End If
'
'   bRet = FtpSetCurrentDirectory(hConnection, ServerDir)
'
'   If bRet = False Then
'      FTP_Upload = False
'   End If
'
'   bRet = FTPPutFile(hConnection, ClientName, Filename, dwType, 0)
'
'   If bRet = False Then
'      FTP_Upload = False
'   End If
'End Function
'
''====================================
'' FTP_Delete 함수
''------------------------------------
'' FTP서버에서 FileDelete..
''====================================
'Public Function FTP_Delete(Filename As String) As Boolean
'   Dim bRet As Boolean
'   FTP_Delete = True
'   If hOpen = 0 Or hConnection = 0 Then
'      MsgBox "현재 FTP서버에 연결이 끊겼습니다.", vbInformation, App.Title
'      FTP_Delete = False
'   End If
'
'   bRet = FtpDeleteFile(hConnection, Filename)
'
'   If bRet = False Then
'       FTP_Delete = False
'   End If
'
'End Function
'
''====================================
'' FTP_CreateDir 함수
''------------------------------------
'' FTP서버에 디렉토리 생성..
''====================================
'Public Function FTP_CreateDir(DirName As String) As Boolean
'
'   Dim bRet As Boolean
'
'   FTP_CreateDir = True
'
'   If hOpen = 0 Or hConnection = 0 Then
'      MsgBox "현재 FTP서버에 연결이 끊겼습니다.", vbInformation, App.Title
'      FTP_CreateDir = False
'   End If
'
'   bRet = FtpCreateDirectory(hConnection, DirName)
'
'   If bRet = False Then
'       FTP_CreateDir = False
'   End If
'End Function
'
''====================================
'' FTP_RemoveDir 함수
''------------------------------------
'' FTP서버에서 디렉토리 제거..
''====================================
'Public Function FTP_RemoveDir(DirName As String) As Boolean
'   Dim bRet As Boolean
'
'   FTP_RemoveDir = True
'
'   If hOpen = 0 Or hConnection = 0 Then
'      MsgBox "현재 FTP서버에 연결이 끊겼습니다.", vbInformation, App.Title
'      FTP_RemoveDir = False
'   End If
'
'   bRet = FtpRemoveDirectory(hConnection, DirName)
'
'   If bRet = False Then
'       FTP_RemoveDir = False
'   End If
'
'End Function
'
''====================================
'' FTP_GetDir 함수
''------------------------------------
'' FTP서버에서 디렉토리 제거..
''====================================
'Public Function FTP_GetDir(DirName As String) As Boolean
'
'   Dim bRet As Boolean
'   szDir = Space(260)
'
'   FTP_GetDir = True
'
'   bRet = FtpSetCurrentDirectory(hConnection, DirName)
'
'   If bRet = False Then
'      FTP_GetDir = False
'   End If
'
'   bRet = FtpGetCurrentDirectory(hConnection, szDir, Len(szDir))
'
'   If bRet = False Then
'      FTP_GetDir = False
'      Exit Function
'   Else
'      szDir = Left(szDir, InStr(1, szDir, Chr(0)) - 1)
'      szDir = szDir & IIf((Right(szDir, 1) = "/"), "*.*", "/*.*")
'   End If
'End Function
'
'Public Function FTPPathChk(ServerDir As String) As Boolean
'
'    '경로체크
'
'    Dim bRet As Boolean
'    FTPPathChk = True
'    bRet = FtpSetCurrentDirectory(hConnection, ServerDir)
'    If bRet = False Then
''        MsgBox "경로체크가 실패하였습니다.", vbExclamation
'        FTPPathChk = False
'    End If
'
'End Function
'
'Public Function FileDelete(Source As String, Dest As String)
'
'    Dim lenFileop As Long
'    Dim foBuf() As Byte
'    Dim fileop As SHFILEOPSTRUCT
'    lenFileop = LenB(fileop)
'    ReDim foBuf(1 To lenFileop)
'    With fileop
'         '.hwnd = Me.hwnd
'         .wFunc = FO_delete
'         .pFrom = Source & vbNullChar & vbNullChar & vbNullChar
'         .pTo = Dest & vbNullChar & vbNullChar
'         .fFlags = FOF_SIMPLEPROGRESS Or FOF_NOCONFIRMATION
'    End With
'
'    Call CopyMemory(foBuf(1), fileop, lenFileop)
'    Call CopyMemory(foBuf(19), foBuf(21), 12)
'    If SHFileOperation(foBuf(1)) <> 0 Then
''        MsgBox "FILE Delete OPERATION FAILED! " & Chr$(13) & "ERROR CODE: " & Err.LastDllError, vbCritical Or vbOKOnly
'    Else
'       If fileop.fAnyOperationsAborted <> 0 Then
'          MsgBox "FILE Delete Operation Failed", vbCritical Or vbOKOnly
'       End If
'    End If
'
'End Function
'
'
