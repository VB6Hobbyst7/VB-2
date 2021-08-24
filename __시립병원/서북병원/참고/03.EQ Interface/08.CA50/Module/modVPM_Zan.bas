Attribute VB_Name = "modVPM_Zan"
Option Explicit

Private Declare Function GetDefaultPrinter Lib "winspool.drv" Alias "GetDefaultPrinterA" (ByVal pszBuffer As String, pcchBuffer As Long) As Long
Private Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" (ByVal pszPrinter As String) As Long

Private Declare Function SHGetFolderPath Lib "shell32" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Private Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
Private Declare Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer

'Used CSIDL values.
'For the full list, search MSDN for "CSIDL Values"
Private Const CSIDL_APPDATA As Long = &H1A
Private Const CSIDL_PROGRAM_FILES  As Long = &H26
Private Const CSIDL_COMMON_APPDATA As Long = &H23

Private Const SHGFP_TYPE_CURRENT = 0
Private Const S_OK As Long = 0

Private Const MAX_PATH = 260

Private Function GetZanPrinterFolder(csidl As Long) As String
    Dim sPath               As String * MAX_PATH
    Dim SpecialFolderPath   As String
    Dim lResult             As Long
  
    sPath = Space$(MAX_PATH)
    lResult = SHGetFolderPath(0, csidl, 0, SHGFP_TYPE_CURRENT, sPath)
                       
    SpecialFolderPath = Left$(sPath, InStr(sPath, vbNullChar) - 1)
       
    GetZanPrinterFolder = SpecialFolderPath
End Function

Public Sub ZanPrinterSetting()
    Dim IniFileFolder   As String
    Dim SaveFolder      As String * MAX_PATH
    Dim lngLength       As Long
    Dim PopupDlgInd     As Integer

    'get the save.ini file full path.
    'Use CSIDL_COMMON_APPDATA instead if
    'all users share the same settings

    IniFileFolder = GetZanPrinterFolder(CSIDL_COMMON_APPDATA)
    If Right$(IniFileFolder, 1) <> "\" Then
        IniFileFolder = IniFileFolder & "\"
    End If

'''    IniFileFolder = IniFileFolder & "zvprt50\Zan Image Printer(color)\save.ini"
    IniFileFolder = IniFileFolder & "zvprt50\" & gtypEQ_INFO.ZIPNM & "\save.ini"
    
    lngLength = GetPrivateProfileString("save", "folder", "", SaveFolder, MAX_PATH, IniFileFolder)
    
    WritePrivateProfileString "save", "basefilename", "[%Year]-[02d%Month]-[02d%Day] [02d%Time+12]", IniFileFolder
'''    PopupDlgInd = GetPrivateProfileInt("save", "popupdialog", 0, IniFileFolder)
'''    WritePrivateProfileString "save", "filexistact", "1", IniFileFolder
    WritePrivateProfileString "save", "folder", gtypEQ_INFO.EQIMGFILEPATH & "\", IniFileFolder
End Sub

Public Sub SET_DEFAULT_PRINTER(ArgPrinterName As String)
    SetDefaultPrinter ArgPrinterName
End Sub

Public Function GET_DEFAULT_PRINTER() As String
    Dim sPrinterName As String
    Dim sNameBuff As String, lLen As Long
    
    GET_DEFAULT_PRINTER = ""
    
    GetDefaultPrinter vbNullChar, lLen
    sNameBuff = Space$(lLen)
    GetDefaultPrinter sNameBuff, lLen
    
    GET_DEFAULT_PRINTER = Left$(sNameBuff, InStr(sNameBuff, vbNullChar) - 1)
End Function


