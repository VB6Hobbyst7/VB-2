Attribute VB_Name = "modNetDir"
'***************************************************************
'Windows API/Global Declarations for :Connect, Disconnect Network
'     s Drives ( EASILY )
'***************************************************************


Public Enum SpecialFolderIDs
    sfidDESKTOP = &H0
    sfidPROGRAMS = &H2
    sfidPERSONAL = &H5
    sfidFAVORITES = &H6
    sfidSTARTUP = &H7
    sfidRECENT = &H8
    sfidSENDTO = &H9
    sfidSTARTMENU = &HB
    sfidDESKTOPDIRECTORY = &H10
    sfidNETHOOD = &H13
    sfidFONTS = &H14
    sfidTEMPLATES = &H15
    sfidCOMMON_STARTMENU = &H16
    sfidCOMMON_PROGRAMS = &H17
    sfidCOMMON_STARTUP = &H18
    sfidCOMMON_DESKTOPDIRECTORY = &H19
    sfidAPPDATA = &H1A
    sfidPRINTHOOD = &H1B
    sfidProgramFiles = &H10000
    sfidCommonFiles = &H10001
End Enum

Public Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As SpecialFolderIDs, ByRef pIdl As Long) As Long
Public Declare Function SHGetPathFromIDListA Lib "shell32" (ByVal pIdl As Long, ByVal pszPath As String) As Long
Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Const RESOURCETYPE_DISK = &H1
Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Declare Function WNetDisconnectDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long

Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long

Declare Function VerInstallFile Lib "version.dll" Alias "VerInstallFileA" (ByVal Flags&, ByVal SrcName$, ByVal DestName$, ByVal SrcDir$, ByVal DestDir$, ByVal CurrDir As Any, ByVal TmpName$, lpTmpFileLen&) As Long
Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Declare Function OSGetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Declare Function GetTempFilename32 Lib "kernel32" Alias "GetTempFileNameA" (ByVal strWhichDrive As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFilename As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)

    
Public Const CSIDL_COMMON_STARTUP = 24
Public Const MAX_PATH = 260
Global Const gintMAX_SIZE% = 255                        'Maximum buffer size
Global Const gintMAX_PATH_LEN% = 260                    ' Maximum allowed path length including path, filename,

Public Const FO_MOVE As Long = &H1
Public Const FO_COPY As Long = &H2
Public Const FO_DELETE As Long = &H3
Public Const FO_RENAME As Long = &H4

Public Const FOF_MULTIDESTFILES As Long = &H1
Public Const FOF_CONFIRMMOUSE As Long = &H2
Public Const FOF_SILENT As Long = &H4
Public Const FOF_RENAMEONCOLLISION As Long = &H8
Public Const FOF_NOCONFIRMATION As Long = &H10
Public Const FOF_WANTMAPPINGHANDLE As Long = &H20
Public Const FOF_CREATEPROGRESSDLG As Long = &H0
Public Const FOF_ALLOWUNDO As Long = &H40
Public Const FOF_FILESONLY As Long = &H80
Public Const FOF_SIMPLEPROGRESS As Long = &H100
Public Const FOF_NOCONFIRMMKDIR As Long = &H200

Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Long
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer ' e.g. = &h0000 = 0
    dwStrucVersionh As Integer ' e.g. = &h0042 = .42
    dwFileVersionMSl As Integer ' e.g. = &h0003 = 3
    dwFileVersionMSh As Integer ' e.g. = &h0075 = .75
    dwFileVersionLSl As Integer ' e.g. = &h0000 = 0
    dwFileVersionLSh As Integer ' e.g. = &h0031 = .31
    dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
    dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
    dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
    dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
    dwFileFlagsMask As Long ' = &h3F For version "0.42"
    dwFileFlags As Long ' e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long ' e.g. VOS_DOS_WINDOWS16
    dwFileType As Long ' e.g. VFT_DRIVER
    dwFileSubtype As Long ' e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long ' e.g. 0
    dwFileDateLS As Long ' e.g. 0
End Type

Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type


Type FileList
    FileNm As String
    FileSize As Long
    FileDtTm As String
    FileVersion As String
    FileExtend As String
    SvrPath As String
    DestPath As String
    flag As Boolean
End Type

Global Const gsAppName = "GetNewVersion.exe"

Global gNetDriveChar() As String    '네트워크 드라이브명
Global gNetDrive() As String        '
Global gClientPath() As String
Global gcServerPath() As String

Global gNetCount As String  '업그레이드 받을 서버의 갯수


' App Path
Global RegHdApp As String
Global Const RegSsApp As String = "App"
Global Const RegK1App As String = "Path"
Global Const RegK2App As String = "ExeName"
' File Server Path
Global RegHdSet As String
Global Const RegSsSet As String = "Setup"
Global Const RegK1Set As String = "Server IP"

Global Const gstrSEP_DIR$ = "\"                         ' Directory separator character
Global Const gstrSEP_AMPERSAND$ = "@"
Global Const gstrSEP_REGKEY$ = "\"                      ' Registration key separator character.
Global Const gstrSEP_DRIVE$ = ":"                       ' Driver separater character, e.g., C:\
Global Const gstrSEP_DIRALT$ = "/"                      ' Alternate directory separator character
Global Const gstrSEP_EXT$ = "."                         ' Filename extension separator character
Global Const gstrSEP_PROGID = "."
Global Const gstrSEP_FILE$ = "|"                        ' Use the character for delimiting filename lists because it is not a valid character in a filename.
Global Const gstrSEP_LIST = "|"
Global Const gstrSEP_URL$ = "://"                       ' Separator that follows HPPT in URL address
Global Const gstrSEP_URLDIR$ = "/"                      ' Separator for dividing directories in URL addresses.

Global Const gstrUNC$ = "\\"                            'UNC specifier \\
Global Const gstrCOLON$ = ":"
Global Const gstrSwitchPrefix1 = "-"
Global Const gstrSwitchPrefix2 = "/"
Global Const gstrCOMMA$ = ","
Global Const gstrDECIMAL$ = "."
Global Const gstrQUOTE$ = """"
Global Const gstrCCOMMENT$ = "//"                       ' Comment specifier used in C, etc.
Global Const gstrASSIGN$ = "="
Global Const gstrINI_PROTOCOL = "Protocol"
Global Const gstrREMOTEAUTO = "RA"
Global Const gstrDCOM = "DCOM"

Global gsWinPath As String
Global gsSysPath As String
Global gsAppPath As String


Global gsFileInfo() As FileList
Global aryCmd() As Variant
Global strProjectId As String
Global strExeName As String
'Global blnDownloadMyself As Boolean
Global Const strCommonPath = "Common"

Sub Main()
'    aryCmd = GetCommandLine(3)
    gExeFile = Command$
    
'    ReDim aryCmd(1)
'    aryCmd(1) = "APS"
'    strProjectId = aryCmd(1)
'    strExeName = aryCmd(2)
    'If Trim(strExeName) = "" Then strExeName = "getnewversion.exe"
    
'    blnDownloadMyself = IIf(aryCmd(2) = "1", True, False)

'    RegHdApp = App.LegalTrademarks & " " & strProjectId
'    RegHdSet = App.LegalTrademarks & " " & strProjectId
    frmVersionCheck.Show
End Sub
   

Public Function GetFileVersion(FilenameAndPath As Variant) As Variant

    Dim lDummy As Long, lSize As Long, rc As Long
    Dim lVerbufferLen As Long, lVerPointer As Long
    Dim sBuffer() As Byte
    Dim udtVerBuffer As VS_FIXEDFILEINFO
    Dim ProdVer As String
    Dim strTmp As String
    
    On Error GoTo HandelCheckFileVersionError
    
    strTmp = FileDateTime(FilenameAndPath)
    
    GetFileVersion = "1"    'vbNullString
    
    lSize = GetFileVersionInfoSize(FilenameAndPath, lDummy)
    If lSize < 1 Then Exit Function
    
    ReDim sBuffer(lSize)
    
    rc = GetFileVersionInfo(FilenameAndPath, 0&, lSize, sBuffer(0))
    rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
    MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
    
    '**** Determine Product Version number ****
    GetFileVersion = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl) & "." & Format$(udtVerBuffer.dwProductVersionLSl)
    'ProdVer = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl)
    'ProdVer = ProdVer & vbCrLf & Format$(udtVerBuffer.dwProductVersionLSh) & "." & Format$(udtVerBuffer.dwProductVersionLSl)
    Exit Function
    
HandelCheckFileVersionError:
    GetFileVersion = ""

End Function


Public Function GetFileDateTime(FilenameAndPath As Variant) As Variant
    On Error GoTo FileNotFound
    GetFileDateTime = FileDateTime(FilenameAndPath)
    Exit Function
FileNotFound:
    GetFileDateTime = ""
    
End Function


Public Function GetSysDir() As String

    Dim Temp As String * 256
    Dim X As Integer
    
    X = GetSystemDirectory(Temp, Len(Temp)) ' Make API Call (Temp will hold return value)
    GetSysDir = Left$(Temp, X) ' Trim Buffer and return String

End Function


Public Function GetWinDir() As String

    Dim Temp As String * 256
    Dim X As Integer
    
    X = GetWindowsDirectory(Temp, Len(Temp)) ' Make API Call (Temp will hold return value)
    GetWinDir = Left$(Temp, X) ' Trim Buffer and return String

End Function
 

 
Public Function NetConnect(Index As Integer, Optional ByVal NetPath As String, Optional ByVal pindex As Long) As Boolean

    Dim X As Long
    
    NetConnect = True
    If Index = 0 Then
        X = WNetAddConnection(NetPath, "", gNetDriveChar(0))
        If X <> 0 Then NetConnect = False
    Else
        X = WNetCancelConnection(gNetDriveChar(0), 1)
        If X <> 0 Then NetConnect = False
    End If

End Function


 
 '-----------------------------------------------------------
 ' FUNCTION GetShortPathName
 '
 ' Retrieve the short pathname version of a path possibly
 '   containing long subdirectory and/or file names
 '-----------------------------------------------------------
 '
 Function GetShortPathName(ByVal strLongPath As String) As String
     Const cchBuffer = 300
     Dim strShortPath As String
     Dim lResult As Long

     On Error GoTo 0
     strShortPath = String(cchBuffer, Chr$(0))
     lResult = OSGetShortPathName(strLongPath, strShortPath, cchBuffer)
     If lResult = 0 Then
         'Error 53 ' File not found
         'Vegas#51193, just use the long name as this is usually good enough
         GetShortPathName = strLongPath
     Else
         GetShortPathName = StripTerminator(strShortPath)
     End If
 End Function
 
'-----------------------------------------------------------
' FUNCTION: GetTempFilename
' Get a temporary filename for a specified drive and
' filename prefix
' PARAMETERS:
'   strDestPath - Location where temporary file will be created.  If this
'                 is an empty string, then the location specified by the
'                 tmp or temp environment variable is used.
'   lpPrefixString - First three characters of this string will be part of
'                    temporary file name returned.
'   wUnique - Set to 0 to create unique filename.  Can also set to integer,
'             in which case temp file name is returned with that integer
'             as part of the name.
'   lpTempFilename - Temporary file name is returned as this variable.
' RETURN:
'   True if function succeeds; false otherwise
'-----------------------------------------------------------
'
Function GetTempFilename(ByVal strDestPath As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, lpTempFilename As String) As Boolean
    If strDestPath = vbNullString Then
        '
        ' No destination was specified, use the temp directory.
        '
        strDestPath = String(gintMAX_PATH_LEN, vbNullChar)
        If GetTempPath(gintMAX_PATH_LEN, strDestPath) = 0 Then
            GetTempFilename = False
            Exit Function
        End If
    End If
    lpTempFilename = String(gintMAX_PATH_LEN, vbNullChar)
    GetTempFilename = GetTempFilename32(strDestPath, lpPrefixString, wUnique, lpTempFilename) > 0
    lpTempFilename = StripTerminator(lpTempFilename)
End Function

'-----------------------------------------------------------
' FUNCTION: GetFileName
'
' Return the filename portion of a path
'
'-----------------------------------------------------------
'
Function GetFileName(ByVal strPath As String) As String
    Dim strFilename As String
    Dim iSep As Integer
    
    strFilename = strPath
    Do
        iSep = InStr(strFilename, gstrSEP_DIR)
        If iSep = 0 Then iSep = InStr(strFilename, gstrCOLON)
        If iSep = 0 Then
            GetFileName = strFilename
            Exit Function
        Else
            strFilename = Right(strFilename, Len(strFilename) - iSep)
        End If
    Loop
End Function


'-----------------------------------------------------------
' FUNCTION: IsWin32
'
' Returns true if this program is running under Win32 (i.e.
'   any 32-bit operating system)
'-----------------------------------------------------------
'
Function IsWin32() As Boolean
    IsWin32 = (IsWindows95() Or IsWindowsNT())
End Function

'-----------------------------------------------------------
' FUNCTION: IsWindows95
'
' Returns true if this program is running under Windows 95
'   or successor
'-----------------------------------------------------------
'
Function IsWindows95() As Boolean
    Const dwMask95 = &H1&
    IsWindows95 = (GetWinPlatform() And dwMask95)
End Function

'-----------------------------------------------------------
' FUNCTION: IsWindowsNT
'
' Returns true if this program is running under Windows NT
'-----------------------------------------------------------
'
Function IsWindowsNT() As Boolean
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

'-----------------------------------------------------------
' FUNCTION: IsWindowsNT4WithoutSP2
'
' Determines if the user is running under Windows NT 4.0
' but without Service Pack 2 (SP2).  If running under any
' other platform, returns False.
'
' IN: [none]
'
' Returns: True if and only if running under Windows NT 4.0
' without at least Service Pack 2 installed.
'-----------------------------------------------------------
'
Function IsWindowsNT4WithoutSP2() As Boolean
    IsWindowsNT4WithoutSP2 = False
    
    If Not IsWindowsNT() Then
        Exit Function
    End If
    
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    strCSDVersion = StripTerminator(osvi.szCSDVersion)
    
    'Is this Windows NT 4.0?
    Const NT4MajorVersion = 4
    Const NT4MinorVersion = 0
    If (osvi.dwMajorVersion <> NT4MajorVersion) Or (osvi.dwMinorVersion <> NT4MinorVersion) Then
        'No.  Return False.
        Exit Function
    End If
    
    'If no service pack is installed, or if Service Pack 1 is
    'installed, then return True.
    Const strSP1 = "SERVICE PACK 1"
    If strCSDVersion = "" Then
        IsWindowsNT4WithoutSP2 = True 'No service pack installed
    ElseIf strCSDVersion = strSP1 Then
        IsWindowsNT4WithoutSP2 = True 'Only SP1 installed
    End If
End Function

'----------------------------------------------------------
' FUNCTION: GetWinPlatform
' Get the current windows platform.
' ---------------------------------------------------------
Public Function GetWinPlatform() As Long
    
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    GetWinPlatform = osvi.dwPlatformId
End Function


Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function


