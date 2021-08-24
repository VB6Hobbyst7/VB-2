VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_PATH = 260
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const NO_ERROR = 0
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_OFFLINE = &H1000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Sec                              As Integer
Private Min                              As Integer
Private Hr                               As Integer
Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As Currency
  ftLastAccessTime As Currency
  ftLastWriteTime As Currency
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Private Const ERROR_NO_MORE_FILES = 18
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_INVALID_PORT_NUMBER = 0
Private Const INTERNET_SERVICE_FTP = 1

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Any, lpLocalFileTime As Any) As Long

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function FTPGetFile Lib "Wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpRenameFile Lib "Wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszOldName As String, ByVal lpszNewName As String) As Boolean
Private Declare Function FtpCreateDirectory Lib "Wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszName As String) As Boolean
Private Declare Function FtpRemoveDirectory Lib "Wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszName As String) As Boolean

Private Declare Function FtpDeleteFile Lib "Wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function FtpOpenFile Lib "Wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal sBuff As String, ByVal Access As Long, ByVal Flags As Long, ByVal Context As Long) As Long
Private Declare Function FTPPutFile Lib "Wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpSetCurrentDirectory Lib "Wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpGetCurrentDirectory Lib "Wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Boolean
Private Declare Function FtpFindFirstFile Lib "Wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpLibFileName As String) As Long

Private Declare Function InternetFindNextFile Lib "Wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
Private Declare Function InternetWriteFile Lib "Wininet.dll" (ByVal hFile As Long, ByRef sBuffer As Byte, ByVal lNumBytesToWrite As Long, dwNumberOfBytesWritten As Long) As Integer
Private Declare Function InternetReadFile Lib "Wininet.dll" (ByVal hFile As Long, ByRef sBuffer As Byte, ByVal lNumBytesToRead As Long, dwNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "Wininet.dll" (ByVal hInet As Long) As Long
Private Declare Function InternetOpen Lib "Wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetConnect Lib "Wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetGetLastResponseInfo Lib "Wininet.dll" Alias "InternetGetLastResponseInfoA" (ByRef lpdwError As Long, ByVal lpszErrorBuffer As String, ByRef lpdwErrorBufferLength As Long) As Boolean

Private Const rDayZeroBias As Double = 109205#   ' Abs(CDbl(#01-01-1601#))
Private Const rMillisecondPerDay As Double = 10000000# * 60# * 60# * 24# / 10000#

Private Const BUFFERSIZE = 255

Private hOpen As Long
Private hConnection As Long
Private hFile As Long

Private dwType      As FTP_TRANSFER_TYPE
Private dwSeman     As INTERNET_FLAG

Private szErrorMessage As String

Private mDirCol As New clsDirList

Public Enum INTERNET_FLAG
    INTERNET_FLAG_ACTIVE = &H0
    INTERNET_FLAG_PASSIVE = &H8000000
End Enum

Public Enum FTP_TRANSFER_TYPE
    FTP_TRANSFER_TYPE_ASCII = &H1
    FTP_TRANSFER_TYPE_BINARY = &H2
End Enum

Public Enum FTP_CONNECTION_STATES
    FTP_CONNECTION_RESOLVING_HOST
    FTP_CONNECTION_HOST_RESOLVED
    FTP_CONNECTION_CONNECTED
    FTP_CONNECTION_AUTHENTICATION
    FTP_USER_LOGGED
    FTP_ESTABLISHING_DATA_CONNECTION
    FTP_DATA_CONNECTION_ESTABLISHED
    FTP_RETRIEVING_DIRECTORY_INFO
    FTP_DIRECTORY_INFO_COMPLETED
    FTP_TRANSFER_STARTING
    FTP_TRANSFER_COMLETED
End Enum
'all possible reply codes that can be sent by ftp server
Private Enum FTP_RESPONSE_CODES
    FTP_RESPONSE_RESTART_MARKER_REPLY = 110
    FTP_RESPONSE_SERVICE_READY_IN_MINUTES = 120
    FTP_RESPONSE_DATA_CONNECTION_ALREADY_OPEN = 125
    FTP_RESPONSE_FILE_STATUS_OK = 150
    FTP_RESPONSE_COMMAND_OK = 200
    FTP_RESPONSE_COMMAND_NOT_IMPLEMENTED_SUPERFLUOUS_AT_THIS_SITE = 202 'superfluous at this site
    FTP_RESPONSE_SYSTEM_STATUS_OR_SYSTEM_HELP_REPLY = 211
    FTP_RESPONSE_DIRECTORY_STATUS = 212
    FTP_RESPONSE_FILE_STATUS = 213
    FTP_RESPONSE_HELP_MESSAGE = 214
    FTP_RESPONSE_NAME_SYSTEM_TYPE = 215
    FTP_RESPONSE_SERVICE_READY_FOR_NEW_USER = 220
    FTP_RESPONSE_SERVICE_CLOSING_CONTROL_CONNECTION = 221
    FTP_RESPONSE_DATA_CONNECTION_OPEN = 225
    FTP_RESPONSE_CLOSING_DATA_CONNECTION = 226
    FTP_RESPONSE_ENTERING_PASSIVE_MODE = 227
    FTP_RESPONSE_USER_LOGGED_IN = 230
    FTP_RESPONSE_REQUESTED_FILE_ACTION_OK_COMPLETED = 250
    FTP_RESPONSE_PATHNAME_CREATED = 257
    FTP_RESPONSE_USER_NAME_OK_NEED_PASSWORD = 331
    FTP_RESPONSE_NEED_ACCOUNT_FOR_LOGIN = 332
    FTP_RESPONSE_REQUESTED_FILE_ACTION_PENDING_FURTHER_INFO = 350
    FTP_RESPONSE_SERVICE_NOT_AVAILABLE_CLOSING_CONTROL_CONNECTION = 421
    FTP_RESPONSE_CANNOT_OPEN_DATA_CONNECTION = 425
    FTP_RESPONSE_CONNECTION_CLOSED_TRANSFER_ABORTED = 426
    FTP_RESPONSE_REQUESTED_FILE_ACTION_NOT_TAKEN = 450
    FTP_RESPONSE_REQUESTED_ACTION_ABORTED = 451
    FTP_RESPONSE_REQUESTED_ACTION_NOT_TAKEN = 452
    FTP_RESPONSE_SYNTAX_ERROR_COMMAND_UNRECOGNIZED = 500
    FTP_RESPONSE_SYNTAX_ERROR_IN_PARAMETERS_OR_ARGUMENTS = 501
    FTP_RESPONSE_COMMAND_NOT_IMPLEMENTED = 502
    FTP_RESPONSE_BAD_SEQUENCE_OF_COMMANDS = 503
    FTP_RESPONSE_COMMAND_NOT_IMPLEMENTED_FOR_THAT_PARAMETER = 504
    FTP_RESPONSE_NOT_LOGGED_IN = 530
    FTP_RESPONSE_NEED_ACCOUNT_FOR_STORING_FILES = 532
    FTP_RESPONSE_REQUESTED_ACTION_NOT_TAKEN_FILE_UNAVAILABLE = 550
    FTP_RESPONSE_REQUESTED_ACTION_ABORTED_PAGE_TYPE_UNKNOWN = 551
    FTP_RESPONSE_REQUESTED_FILE_ACTION_ABORTED_EXCEEDED_STORAGE_ALLOCATION = 552
    FTP_RESPONSE_REQUESTED_ACTION_NOT_TAKEN_FILE_NAME_NOT_ALLOWED = 553
End Enum
Public Enum FtpErrors
    ERROR_FTP_WINSOCK_AddressInUse
    ERROR_FTP_WINSOCK_AddressNotAvailable
    ERROR_FTP_WINSOCK_AlreadyComplete
    ERROR_FTP_WINSOCK_AlreadyConnected
    ERROR_FTP_WINSOCK_BadState
    ERROR_FTP_WINSOCK_ConnectAborted
    ERROR_FTP_WINSOCK_ConnectionRefused
    ERROR_FTP_WINSOCK_ConnectionReset
    ERROR_FTP_WINSOCK_GetNotSupported
    ERROR_FTP_WINSOCK_HostNotFound
    ERROR_FTP_WINSOCK_HostNotFoundTryAgain
    ERROR_FTP_WINSOCK_InProgress
    ERROR_FTP_WINSOCK_InvalidArg
    ERROR_FTP_WINSOCK_InvalidArgument
    ERROR_FTP_WINSOCK_InvalidOp
    ERROR_FTP_WINSOCK_InvalidPropertyValue
    ERROR_FTP_WINSOCK_MsgTooBig
    ERROR_FTP_WINSOCK_NetReset
    ERROR_FTP_WINSOCK_NetworkSubsystemFailed
    ERROR_FTP_WINSOCK_NetworkUnreachable
    ERROR_FTP_WINSOCK_NoBufferSpace
    ERROR_FTP_WINSOCK_NoData
    ERROR_FTP_WINSOCK_NonRecoverableError
    ERROR_FTP_WINSOCK_NotConnected
    ERROR_FTP_WINSOCK_NotInitialized
    ERROR_FTP_WINSOCK_NotSocket
    ERROR_FTP_WINSOCK_OpCanceled
    ERROR_FTP_WINSOCK_OutOfMemory
    ERROR_FTP_WINSOCK_OutOfRange
    ERROR_FTP_WINSOCK_PortNotSupported
    ERROR_FTP_WINSOCK_SetNotSupported
    ERROR_FTP_WINSOCK_SocketShutdown
    ERROR_FTP_WINSOCK_Success
    ERROR_FTP_WINSOCK_Timedout
    ERROR_FTP_WINSOCK_Unsupported
    ERROR_FTP_WINSOCK_WouldBlock
    ERROR_FTP_WINSOCK_WrongProtocol
    ERROR_FTP_PROTOCOL_SERVICE_READY_IN_MINUTES
    ERROR_FTP_PROTOCOL_USER_NAME_OK_NEED_PASSWORD
    ERROR_FTP_PROTOCOL_NEED_ACCOUNT_FOR_LOGIN
    ERROR_FTP_PROTOCOL_REQUESTED_FILE_ACTION_PENDING_FURTHER_INFO
    ERROR_FTP_PROTOCOL_SERVICE_NOT_AVAILABLE_CLOSING_CONTROL_CONNECTION
    ERROR_FTP_PROTOCOL_CANNOT_OPEN_DATA_CONNECTION
    ERROR_FTP_PROTOCOL_CONNECTION_CLOSED_TRANSFER_ABORTED
    ERROR_FTP_PROTOCOL_REQUESTED_FILE_ACTION_NOT_TAKEN
    ERROR_FTP_PROTOCOL_REQUESTED_ACTION_ABORTED
    ERROR_FTP_PROTOCOL_REQUESTED_ACTION_NOT_TAKEN
    ERROR_FTP_PROTOCOL_SYNTAX_ERROR_COMMAND_UNRECOGNIZED
    ERROR_FTP_PROTOCOL_SYNTAX_ERROR_IN_PARAMETERS_OR_ARGUMENTS
    ERROR_FTP_PROTOCOL_COMMAND_NOT_IMPLEMENTED
    ERROR_FTP_PROTOCOL_BAD_SEQUENCE_OF_COMMANDS
    ERROR_FTP_PROTOCOL_COMMAND_NOT_IMPLEMENTED_FOR_THAT_PARAMETER
    ERROR_FTP_PROTOCOL_NOT_LOGGED_IN
    ERROR_FTP_PROTOCOL_NEED_ACCOUNT_FOR_STORING_FILES
    ERROR_FTP_PROTOCOL_REQUESTED_ACTION_NOT_TAKEN_FILE_UNAVAILABLE
    ERROR_FTP_PROTOCOL_REQUESTED_ACTION_ABORTED_PAGE_TYPE_UNKNOWN
    ERROR_FTP_PROTOCOL_REQUESTED_FILE_ACTION_ABORTED_EXCEEDED_STORAGE_ALLOCATION
    ERROR_FTP_PROTOCOL_REQUESTED_ACTION_NOT_TAKEN_FILE_NAME_NOT_ALLOWED
    ERROR_FTP_USER_TIMEOUT
    ERROR_FTP_USER_TRANSFER_IN_PROGRESS
End Enum

Public Event Connect()
Public Event StateChanged(state As FTP_CONNECTION_STATES)
Public Event FileStart(ITotalByte As Long)
Public Event FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long)
Public Event FileEnd(ITotalByte As Long)
Public Event FtpError(szFunc As String, ErrCode As Long, Desc As String)

Private Function Win32ToVbTime(ft As Currency) As Date
    Dim ftl As Currency
    ' Call API to convert from UTC time to local time
    If FileTimeToLocalFileTime(ft, ftl) Then
        ' Local time is nanoseconds since 01-01-1601
        ' In Currency that comes out as milliseconds
        ' Divide by milliseconds per day to get days since 1601
        ' Subtract days from 1601 to 1899 to get VB Date equivalent
        Win32ToVbTime = CDate((ftl / rMillisecondPerDay) - rDayZeroBias)
    Else
        Call MsgBox(Err.LastDllError)
    End If
End Function

Property Get Directory() As clsDirList
   Set Directory = mDirCol
End Property

Private Sub ErrorOut(ByVal dwError As Long, ByRef szFunc As String)
    Dim dwRet As Long
    Dim dwTemp As Long
    Dim szString As String * 2048
    
    dwRet = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, _
                      GetModuleHandle("wininet.dll"), dwError, 0, _
                      szString, 256, 0)
    szErrorMessage = szFunc & " error code: " & dwError & " Message: " & szString
    If (dwError = 12003) Then
        ' Extended error information was returned
        dwRet = InternetGetLastResponseInfo(dwTemp, szString, 2048)
        szErrorMessage = szString
    End If
    RaiseEvent FtpError(szFunc, dwError, szString)
    
End Sub

Property Get GetLastErrorMessage() As String
    GetLastErrorMessage = szErrorMessage
End Property

Public Sub CloseConnection()
    If hConnection <> 0 Then
        Call InternetCloseHandle(hConnection)
    End If
    hConnection = 0
End Sub

Public Function SimpleFTPPutFile(sLocal As String, sRemote As String) As Boolean
    If (FTPPutFile(hConnection, sLocal, sRemote, dwType, 0) = False) Then
        Call ErrorOut(Err.LastDllError, "SimpleFtpPutFile")
        SimpleFTPPutFile = False
        Exit Function
    Else
        SimpleFTPPutFile = True
    End If
End Function
 
Public Function RenameFTPFile(sExisting As String, sNewName As String) As Boolean
    If (FtpRenameFile(hConnection, sExisting, sNewName) = False) Then
        Call ErrorOut(Err.LastDllError, "RenameFTPFile")
        RenameFTPFile = False
        Exit Function
    Else
        RenameFTPFile = True
    End If
End Function

Public Function CreateFTPDirectory(sDirectory As String) As Boolean
    If (FtpCreateDirectory(hConnection, sDirectory) = False) Then
        Call ErrorOut(Err.LastDllError, "CreateFTPDirectory")
        CreateFTPDirectory = False
        Exit Function
    Else
        CreateFTPDirectory = True
    End If
End Function

Public Function RemoveFTPDirectory(sDirectory As String) As Boolean
    If (FtpRemoveDirectory(hConnection, sDirectory) = False) Then
        Call ErrorOut(Err.LastDllError, "RemoveFTPDirectory")
        RemoveFTPDirectory = False
        Exit Function
    Else
        RemoveFTPDirectory = True
    End If
End Function
 
Public Function DeleteFTPFile(sRemote As String) As Boolean
    If (FtpDeleteFile(hConnection, sRemote) = False) Then
        Call ErrorOut(Err.LastDllError, "DeleteFTPFile")
        DeleteFTPFile = False
        Exit Function
    Else
        DeleteFTPFile = True
    End If
End Function

Public Function OpenConnection(sServer As String, sPort As String, sUser As String, sPassword As String) As Boolean
    If hConnection <> 0 Then
        Call InternetCloseHandle(hConnection)
    End If
    
    hConnection = InternetConnect(hOpen, sServer, sPort, sUser, sPassword, INTERNET_SERVICE_FTP, dwSeman, 0)
    If hConnection = 0 Then
        Call ErrorOut(Err.LastDllError, "InternetConnect")
        OpenConnection = False
        Exit Function
    Else
        RaiseEvent Connect
        OpenConnection = True
    End If
End Function

Public Function FTPUploadFile(sLocal As String, sRemote As String) As Boolean
    Dim Data(BUFFERSIZE - 1) As Byte
    Dim Written As Long
    Dim Size As Long
    Dim Sum As Long
    Dim lBlock As Long
    Dim FN As Long

On Error GoTo ErrorRoutine

    Sum = 0
    lBlock = 0
    sLocal = Trim(sLocal)
    sRemote = Trim(sRemote)
    
    If sLocal <> "" And sRemote <> "" Then
      hFile = FtpOpenFile(hConnection, sRemote, GENERIC_WRITE, dwType, 0)
      If hFile = 0 Then
          Call ErrorOut(Err.LastDllError, "FtpOpenFile:PutFile")
          FTPUploadFile = False
          RaiseEvent FileEnd(Size)
          Exit Function
      End If
      FN = FreeFile
      Open sLocal For Binary Access Read As FN
      Size = LOF(FN)
      RaiseEvent FileStart(Size)
      For lBlock = 1 To Size \ BUFFERSIZE
          Get FN, , Data
          If (InternetWriteFile(hFile, Data(0), BUFFERSIZE, Written) = 0) Then
              FTPUploadFile = False
              Call ErrorOut(Err.LastDllError, "InternetWriteFile")
              RaiseEvent FileEnd(Size)
              Exit Function
          End If
          DoEvents
          Sum = Sum + BUFFERSIZE
          RaiseEvent FileTransferProgress(Sum, Size)
      Next lBlock
      
      Get FN, , Data
      If (InternetWriteFile(hFile, Data(0), Size Mod BUFFERSIZE, Written) = 0) Then
          FTPUploadFile = False
          Call ErrorOut(Err.LastDllError, "InternetWriteFile2")
          RaiseEvent FileEnd(Size)
          Exit Function
      End If
      
      Sum = Sum + (Size Mod BUFFERSIZE)
      Size = Sum
      RaiseEvent FileTransferProgress(Sum, Size)
      Close FN
      Call InternetCloseHandle(hFile)
      RaiseEvent FileEnd(Size)
      FTPUploadFile = True
   End If
Exit Function
ErrorRoutine:
    If FN Then
        Close FN
    End If
    RaiseEvent FileEnd(Size)

End Function

Public Function FTPDownloadFile(sLocal As String, sRemote As String) As Boolean
    Dim Data(BUFFERSIZE - 1) As Byte ' array of 100 elements 0 to 99
    Dim Written As Long
    Dim Size As Long
    Dim Sum As Long
    Dim lBlock As Long
    Dim FN As Long

On Error GoTo ErrorRoutine

    DoEvents
        DoEvents
        FTPDownloadFile = False
    DoEvents
        DoEvents
        RaiseEvent StateChanged(FTP_TRANSFER_STARTING)
    DoEvents
        DoEvents
        Sum = 0
        lBlock = 0
    DoEvents
        DoEvents
        sLocal = sLocal
        sRemote = sRemote
    DoEvents
        DoEvents
        If sLocal = "" Or sRemote = "" Then
            RaiseEvent FileEnd(Size)
            Exit Function
        End If
        Size = GetFTPFileSize(sRemote)
    DoEvents
        DoEvents
    
    If Size >= 0 Then
    DoEvents
        hFile = FtpOpenFile(hConnection, sRemote, GENERIC_READ, dwType, 0)
    DoEvents
        DoEvents
        If hFile = 0 Then
            Call ErrorOut(Err.LastDllError, "FtpOpenFile:GetFile")
            RaiseEvent FileEnd(Size)
            Exit Function
        End If
    DoEvents
        
        FN = FreeFile
        RaiseEvent FileStart(Size)
        Open sLocal For Binary Access Write As FN '#1
    DoEvents
        DoEvents
        Seek FN, 1
        Sum = 1
    DoEvents
        DoEvents
        For lBlock = 1 To Size \ BUFFERSIZE
            If (InternetReadFile(hFile, Data(0), BUFFERSIZE, Written) = 0) Then
                Call ErrorOut(Err.LastDllError, "InternetReadFile")
                Close FN '#1
                RaiseEvent FileEnd(Size)
                Exit Function
            End If
                  DoEvents
            DoEvents
                Put FN, , Data
                DoEvents
                Sum = Sum + BUFFERSIZE
                DoEvents
                RaiseEvent FileTransferProgress(Sum, Size)
                  DoEvents
            DoEvents
        Next lBlock
        
        ReDim Data2((Size Mod BUFFERSIZE) - 1) As Byte
    DoEvents
        DoEvents
        If (InternetReadFile(hFile, Data2(0), Size Mod BUFFERSIZE, Written) = 0) Then
            Call ErrorOut(Err.LastDllError, "InternetReadFile2")
            Close FN
                  DoEvents
            RaiseEvent FileEnd(Size)
            DoEvents
            Exit Function
        End If
                    
        Put FN, , Data2
              DoEvents
    DoEvents
        Sum = Sum + (Size Mod BUFFERSIZE)
        Size = Sum
        RaiseEvent FileTransferProgress(Sum, Size)
              DoEvents
    DoEvents
        Close FN
    DoEvents
        DoEvents
        Call InternetCloseHandle(hFile)
        FTPDownloadFile = True
    DoEvents
        DoEvents
        RaiseEvent StateChanged(FTP_TRANSFER_COMLETED)
        RaiseEvent FileEnd(Size)
    End If
Exit Function
ErrorRoutine:
    If FN Then
        Close FN
    End If
    RaiseEvent FileEnd(Size)
End Function

Public Function SimpleFTPGetFile(sLocal As String, sRemote As String) As Boolean
   ' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
    If (FTPGetFile(hConnection, sRemote, sLocal, False, FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
        Call ErrorOut(Err.LastDllError, "SimpleFtpGetFile")
        SimpleFTPGetFile = False
        Exit Function
    Else
        SimpleFTPGetFile = True
    End If
End Function

Public Function GetFTPDirectory() As String
    Dim szDir As String
    szDir = String(1024, Chr$(0))
    If (FtpGetCurrentDirectory(hConnection, szDir, 1024) = False) Then
        Call ErrorOut(Err.LastDllError, "FtpGetCurrentDirectory")
        Exit Function
    Else
        GetFTPDirectory = left(szDir, InStr(1, szDir, String(1, 0), vbBinaryCompare) - 1)
    End If
End Function

Public Function SetFTPDirectory(sDir As String)
    If (FtpSetCurrentDirectory(hConnection, sDir) = False) Then
        Call ErrorOut(Err.LastDllError, "FtpSetCurrentDirectory")
        SetFTPDirectory = False
        Exit Function
    Else
        SetFTPDirectory = True
    End If
End Function

Public Function GetFTPFileSize(sFile As String) As Long
    Dim szDir As String
    Dim hFind As Long
    Dim nLastError As Long
    Dim pData As WIN32_FIND_DATA
    
    hFind = FtpFindFirstFile(hConnection, sFile, pData, 0, 0)
    nLastError = Err.LastDllError
    If hFind = 0 Then
        If (nLastError = ERROR_NO_MORE_FILES) Then
            GetFTPFileSize = -1  ' File not found
        Else
            GetFTPFileSize = -2  ' Other error
            Call ErrorOut(Err.LastDllError, "FtpFindFirstFile")
        End If
    End If
           
    GetFTPFileSize = pData.nFileSizeLow
    Call InternetCloseHandle(hFind)
End Function

Public Function GetDirectoryListing(sFilter As String) As clsDirList
    Dim szDir As String
    Dim hFind As Long
    Dim nLastError As Long
    Dim dError As Long
    Dim ptr As Long
    Dim pData As WIN32_FIND_DATA
    Dim sFileName As String
        
    Set mDirCol = Nothing
    DoEvents
    hFind = FtpFindFirstFile(hConnection, sFilter, pData, 0, 0)
    DoEvents
    nLastError = Err.LastDllError
    DoEvents
    If hFind = 0 Then
        If (nLastError <> ERROR_NO_MORE_FILES) Then
            Call ErrorOut(Err.LastDllError, "FtpFindFirstFile")
        End If
        Exit Function
    End If
    
    dError = NO_ERROR
    Dim bRet As Boolean

    sFileName = left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
    Call mDirCol.Add(pData.dwFileAttributes, Win32ToVbTime(pData.ftCreationTime), Win32ToVbTime(pData.ftLastAccessTime), Win32ToVbTime(pData.ftLastWriteTime), pData.nFileSizeLow, sFileName)
    Do
        DoEvents
        pData.cFileName = String(MAX_PATH, 0)
        bRet = InternetFindNextFile(hFind, pData)
        DoEvents
        If Not bRet Then
            dError = Err.LastDllError
            If dError = ERROR_NO_MORE_FILES Then
                Exit Do
            Else
                Call ErrorOut(Err.LastDllError, "InternetFindNextFile")
                Call InternetCloseHandle(hFind)
                Exit Function
            End If
            DoEvents
        Else
            sFileName = left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
            Call mDirCol.Add(pData.dwFileAttributes, Win32ToVbTime(pData.ftCreationTime), Win32ToVbTime(pData.ftLastAccessTime), Win32ToVbTime(pData.ftLastWriteTime), pData.nFileSizeLow, sFileName)
        End If
        DoEvents
    Loop
       DoEvents
    Set GetDirectoryListing = mDirCol
    DoEvents
    Call InternetCloseHandle(hFind)
End Function

Public Property Get TransferType() As FTP_TRANSFER_TYPE
    dwType = TransferType
End Property

Public Property Let TransferType(iType As FTP_TRANSFER_TYPE)
    dwType = iType
    DoEvents
    DoEvents
End Property

Public Property Get ConnectMode() As INTERNET_FLAG
    ConnectMode = dwSeman
End Property

Public Property Let ConnectMode(iMode As INTERNET_FLAG)
    dwSeman = iMode
    DoEvents
    DoEvents
End Property

Public Property Get IsConncted() As Boolean
    If hConnection Then
        IsConncted = True
    Else
        IsConncted = False
    End If
    
End Property

Private Sub Class_Initialize()
    hOpen = InternetOpen("SMC FTP Client", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If hOpen = 0 Then
        Call ErrorOut(Err.LastDllError, "InternetOpen")
    End If
    dwType = FTP_TRANSFER_TYPE_ASCII
    dwSeman = INTERNET_FLAG_ACTIVE
    hConnection = 0
End Sub

Private Sub Class_Terminate()
    Call InternetCloseHandle(hOpen)
End Sub




