Attribute VB_Name = "modCommon"
Option Explicit

Private Const PROCESS_TERMINATE = &H1&
Private Const SYNCHRONIZE = &H100000
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5
Private Const WS_CAPTION = &HC00000         '## 타이틀바를 가진윈도우 생성
Private Const GWL_STYLE = (-16)             '## 윈도우 스타일 설정
Private Const WM_SETICON As Long = &H80
Private Const SHGFI_ICON = &H100

Private Type SHELLEXECUTEINFO
   cbSize As Long
   fMask As Long
   hwnd As Long
   lpVerb As String
   lpFile As String
   lpParameters As String
   lpDirectory As String
   nShow As Long
   hInstApp As Long
   lpIDList As Long
   lpClass As String
   hkeyClass As Long
   dwHotKey As Long
   hIcon As Long
   hProcess As Long
End Type

 

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

 

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

 

Private Const SEE_MASK_IDLIST = &H4

 

Private Type SHFILEINFO
   hIcon As Long                       '## 실행파일의 아콩핸들
   iIcon As Long                       '## 시스템이미지 목록의 아콩 인덱스 번호
   dwAttributes As Long                '## 파일속성 Flag
   szDisplayName As String * 128       '## 실행파일 이름
   szTypeName As String * 80           '## 실행파일 유형
End Type

Public Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Public Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public xlApp As Object    '## 엑셀App 인스턴스 생성시에 사용되는 할당되는 전역변수
Public XlWb  As Object    '## 엑셀App Workbook개체를 할당할 전역변수

Public Const XL_NOTRUNNING As Long = 429

Public bSubFlag As Boolean                          '## 프린트와 엑셀저장 FLAG
Public bLoop    As Boolean                          '## 프린트 서브폼의 비 활성화여부






Declare Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lplFileName As String) As Long

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

'통신설정
Type config
    gPort       As String
    gSpeed      As String
    gParity     As String
    gDataBit    As String
    gStopBit    As String
    gStartBit   As String
    gRTSEnable  As String
    gDTREnable  As String
    ACKUse      As String
End Type


Public gSetup As config
Public gPart As String
Public gGubun As Integer

Public gLIS As String           'LIS 업체
Public gEquip As String         '장비명
Public gEquipCode As String     '장비코드 = MDB에 등록된코드와 매치
Public gMode As String          '화면모드(0:분리, 1:통합)
Public gSave As String          '저장여부(0:자동, 1:수동)
Public gBarLen As Integer       '바코드자리수
Public gCommProtocol As String  '통신프로토콜

Public gIP As String
Public gOrderExam As String
Public gAllExam As String
Public gOrder As String

Public gSndState As String
Public gRecodeType As String

Public gQCEquip As String
Public gPreSpecID As String
Public gPreRow As Long
Public gOrdRow As Long
Public gEquipID As String

Public gCurMsgCnt As String

Public gHeader As String
Public gPatient As String

Public gMsgEnd As String

Type DB_Parm
    DBType  As String
    Driver  As String
    USER    As String
    Passwd  As String
    Server  As String
    DB      As String
    HostName    As String
    LocalDB As String
End Type

Public gDB_Parm As DB_Parm

Type DRDB_Parm
    DBType  As String
    Driver  As String
    USER    As String
    Passwd  As String
    Server  As String
    DB      As String
    HostName    As String
    LocalDB     As String
    ServerIP    As String
    ServerPort  As String
End Type

Public gDRDB_Parm As DRDB_Parm

Type DBTBL_Parm
    ORDTABLE As String
    RSLTTABLE As String
    MSTTABLE As String
End Type

Public gDBTBL_Parm As DBTBL_Parm

Type DBCOLUMN_Parm
    ORDDATE As String
    RSLTDATE As String
    BARCODE As String
    pID As String
    PNAME As String
    PSEX As String
    PAGE As String
    TESTCD As String
    RESULT As String
    INTRESULT As String
    STATUS As String
    JUDGE As String
    MACHCD As String
    USER  As String
End Type

Public gDBCOLUMN_Parm As DBCOLUMN_Parm


Public gUserID As String

Public comState As String
Public comsignal As String
Public comSend As String

Public gOrderMessage As String
Public gOrderCnt As Integer
Public gNACKCnt As Integer
Public gPreMsg As String
Public gACKSig As Integer
Public gIFUser As String


Public gArrEquip() As String


'-- Order
Public Const colSpecNo = 0  '미사용
Public Const colCheckBox = 1
Public Const colSeqNo = 2
Public Const colOrdDate = 3
Public Const colBarcode = 4
Public Const colRack = 5
Public Const colPos = 6
Public Const colPID = 7
Public Const colPName = 8
'Public Const colSex = 9 '-- Sex/Age
Public Const colWell = 9 '-- Position
Public Const colOther1 = 10
Public Const colOther2 = 11
'Public Const colOther3 = 11
Public Const colOCnt = 12
Public Const colRCnt = 13
Public Const colState = 14

'-- Result
Public Const colRstCheck = 1
Public Const colTestNm = 2
Public Const colTestCd = 3
Public Const colChannel = 4
Public Const colEqpResult = 5
Public Const colLisResult = 6
Public Const colComment = 7
Public Const colFlag = 8
Public Const colN = 9
Public Const colD = 10
Public Const colP = 11
Public Const colC = 12
Public Const colPrevResult = 13
Public Const colPrevTestDt = 14
Public Const colPrevBarcode = 15
Public Const colReference = 16
Public Const colOther = 17
'=================================

Type RecvData
    NoOrder     As Boolean
    BarNo       As String
    Seq         As String
    RackNo      As String
    TubePos     As String
    Order       As String
    IsSending   As Boolean
    SendCnt     As Integer
End Type

Public mOrder As RecvData

Type IntfData
    SpcmNo   As String
    PatNo    As String
    BarNo    As String
    Seq         As String
    TESTCD   As String
    RackNo   As String
    TubePos  As String
    MnmCd    As String
    MnmNm    As String
    MCnt     As String
    Rst      As String
    SpcPos   As String
End Type

Public mResult As IntfData

Type MicroDic
    MicrosCnt        As Integer
    MicroRst         As String
End Type

Public mMicro As MicroDic
Public gComment_All As String
Public gComment_Code As String

'=================================

Public Function STX() As String
    STX = Chr(2)
End Function

Public Function ETX() As String
    ETX = Chr(3)
End Function

Public Function SOH() As String
    SOH = Chr(1)
End Function

Public Function chrEOT() As String
    chrEOT = Chr(4)
End Function

Public Function chrENQ() As String
    chrENQ = Chr(5)
End Function

Public Function ACK() As String
    ACK = Chr(6)
End Function

Public Function cTAB() As String
    cTAB = Chr(9)
End Function

Public Function LF() As String
    LF = Chr(10)
End Function

Public Function CR() As String
    CR = Chr(13)
End Function

Public Function NAK() As String
    NAK = Chr(15)
End Function

Public Function cSPC() As String
    cSPC = Chr(20)
End Function

Public Function ETB() As String
    ETB = Chr(23)
End Function

Function GetSetup() As Boolean
'---------------------------------------------------------------------------------------------------------------------
'                       Setup  File을 읽어온다.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 100

    db_tmp = ""

    GetSetup = False

    '== 장비 관련 설정  ==============================================================================
    db_tmp = ""
    Call GetPrivateProfileString("config", "Equip", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquip = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "EquipCode", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquipCode = Trim(frmInterface.txtTemp)
    
    '== 통신 관련 설정  ==============================================================================
    db_tmp = ""
    Call GetPrivateProfileString("config", "gPort", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gPort = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gSpeed", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gSpeed = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gParity", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gParity = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gDataBit", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gDataBit = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gStopBit", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gStopBit = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gStartBit", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gStartBit = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "gRTSEnable", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gRTSEnable = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gDTREnable", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gDTREnable = Trim(frmInterface.txtTemp)
   
    ' 화면모드
    db_tmp = ""
    Call GetPrivateProfileString("config", "IFMode", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gMode = Trim(frmInterface.txtTemp)

    ' 자동/수동 저장
    db_tmp = ""
    Call GetPrivateProfileString("config", "AutoSave", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSave = Trim(frmInterface.txtTemp)
    
    ' 바코드 길이
    db_tmp = ""
    Call GetPrivateProfileString("config", "BarLen", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gBarLen = Trim(frmInterface.txtTemp)

    ' 통신프로토콜
    db_tmp = ""
    Call GetPrivateProfileString("config", "CommProtocol", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gCommProtocol = Trim(frmInterface.txtTemp)
   
   
    '== DB 관련 설정    ==============================================================================
    Call GetPrivateProfileString("DATABASE", "dbtype", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.DBType = Trim(frmInterface.txtTemp)
    
    Call GetPrivateProfileString("DATABASE", "server", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.Server = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "uid", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.USER = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "pwd", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.Passwd = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "database", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.DB = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("Server", "IFUser", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gUserID = Trim(frmInterface.txtTemp)

    '-- osw 추가
    Call GetPrivateProfileString("DRDATABASE", "server", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDRDB_Parm.Server = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DRDATABASE", "uid", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDRDB_Parm.USER = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DRDATABASE", "pwd", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDRDB_Parm.Passwd = Trim(frmInterface.txtTemp)
    
    '==  Winsock 관련 설정    ==============================================================================
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "ServerIP", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDRDB_Parm.ServerIP = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "ServerPort", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDRDB_Parm.ServerPort = Trim(frmInterface.txtTemp)
        
    '== DB Table 관련 설정    ==============================================================================
    db_tmp = ""
    Call GetPrivateProfileString("Order", "ORDTABLE", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBTBL_Parm.ORDTABLE = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("Order", "RSLTTABLE", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBTBL_Parm.RSLTTABLE = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "MSTTABLE", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBTBL_Parm.MSTTABLE = Trim(frmInterface.txtTemp)
        
    '== DB Table Column 관련 설정    =======================================================================
    db_tmp = ""
    Call GetPrivateProfileString("Order", "ORDDATE", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.ORDDATE = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "ORDDATE", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.ORDDATE = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "RSLTDATE", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.RSLTDATE = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "BARCODE", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.BARCODE = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "PID", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.pID = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "PNAME", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.PNAME = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "PSEX", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.PSEX = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "PAGE", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.PAGE = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "TESTCD", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.TESTCD = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "RESULT", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.RESULT = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Order", "INTRESULT", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.INTRESULT = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Order", "STATUS", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.STATUS = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Order", "JUDGE", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.JUDGE = Trim(frmInterface.txtTemp)
  
    db_tmp = ""
    Call GetPrivateProfileString("Order", "MACHCD", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.MACHCD = Trim(frmInterface.txtTemp)
  
    db_tmp = ""
    Call GetPrivateProfileString("Order", "USER", "", db_tmp, 100, App.Path & "\didim.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.USER = Trim(frmInterface.txtTemp)
  
    GetSetup = True

End Function

Public Function ASTM_CSum(ByVal CheSum As String) As String
    Dim Tot  As Integer
    Dim sStr As String
    Dim i As Integer
    
    For i = 1 To Len(CheSum)
        Tot = Tot + Asc(Mid(CheSum, i, 1))
    Next i
    
    sStr = Hex(Tot)
    If Len(sStr) = 1 Then
        sStr = "0" & sStr
    End If
    
    ASTM_CSum = Right(sStr, 2)
End Function

Public Function CS(ByVal CheSum As String) As String
    Dim Tot  As Integer
    Dim sStr As String
    Dim i As Integer
    
    For i = 1 To Len(CheSum)
        Tot = Tot + Asc(Mid(CheSum, i, 1))
    Next i
    
    Tot = 256 - (Tot Mod 256)
    sStr = Hex(Tot)
    If Len(sStr) = 1 Then
        sStr = "0" & sStr
    End If
    
    CS = Right(sStr, 2)
End Function

Public Function MOR() As String
    MOR = Chr(2) & ">" & Chr(3) & "3E" & Chr(13)
End Function

Public Function REP() As String
    REP = Chr(2) & "?" & Chr(3) & "3F" & Chr(13)
End Function

Public Sub SetRawData(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    
    If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.Path & "\Log")
    End If
    
    sFileName = Format(CDate(frmInterface.dtpToday), "yyyy-mm-dd")
    
    Open App.Path & "\Log\" & sFileName & ".txt" For Append As FilNum
'    Print #FilNum, Format(Time, "hh:nn:ss") & " " & argSQL
    Print #FilNum, Format(Time, "hh:nn:ss") & " " & argSQL
    Close FilNum
    
End Sub



'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열을 구분자를 이용해 구분해 지정한 위치의 문자열을 구함
'   인수 :
'       1.pText      : 구분자로 구성된 문자열
'       2.pPosiion   : 위치
'       3.pDelimiter : 구분자
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition 인수가 1인 경우 For문 Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function


Public Function ReplaceVal(ByVal pValue As String) As String
    ReplaceVal = Replace(pValue, """", "")
End Function

'-----------------------------------------------------------------------------'
'   기능 : 수신한 Result Flags에 대한 상세설명 조회
'-----------------------------------------------------------------------------'
Public Function GetInfo(ByVal pFlag As String)
    Dim strInfo     As String

    If pFlag = "" Then Exit Function

    Select Case pFlag
        Case "+":   strInfo = "Over the upper control limit"
        Case "-":   strInfo = "Under the lower control limit"
        Case "*":   strInfo = "Analysis error occurred, disparate data of mean data occurred, or Fbg was over analysis range."
        Case "!":   strInfo = "Coagulation time was obtained by re-dilution analysis."
        Case ">":   strInfo = "Over the upper report limit."
        Case "<":   strInfo = "Under the lower report limit."
    End Select

    GetInfo = strInfo
End Function

'-----------------------------------------------------------------------------'
'   기능 : 수신한 Abnormal Flag에 대한 설명조회
'-----------------------------------------------------------------------------'
Public Function GetInfo_Centaur(ByVal pFlag As String) As String
    Dim aryFlags() As String
    Dim strInfo    As String
    Dim i          As Long
    
    aryFlags = Split(pFlag, "\")
    
    For i = LBound(aryFlags) To UBound(aryFlags)
        If i > 0 Then
            strInfo = strInfo & vbCrLf & Space(2)
        Else
            strInfo = "[Abnormal Flags]" & vbCrLf & Space(2)
        End If
        
        Select Case aryFlags(i)
            Case "L":   strInfo = strInfo & "Below Reference Range"
            Case "H":   strInfo = strInfo & "Above Reference Range"
            Case "<":   strInfo = strInfo & "Below Concentration Range"
            Case ">":   strInfo = strInfo & "Above Concentration Range"
        End Select
    Next i
    GetInfo_Centaur = strInfo
End Function

Public Sub SpreadSheetSort(ByRef Spread As fpSpread, ByVal Col As Integer, Optional ByVal SortType As Integer = 1)
    Dim intCount As Integer
    Dim strDataField As String
    'SortType
    ' 0 : none
    ' 1 : ascending
    ' 2 : descending

    With Spread
        .Col = 1: .Col2 = .MaxCols
        .Row = 1: .Row2 = .DataRowCnt
        .SortBy = 0
        .SortKey(1) = Col       '정렬키 열번호

        If SortType = 0 Then
            .SortKeyOrder(1) = SortKeyOrderNone
        ElseIf SortType = 1 Then
            .SortKeyOrder(1) = SortKeyOrderAscending
        ElseIf SortType = 2 Then
            .SortKeyOrder(1) = SortKeyOrderDescending
        Else
            .SortKeyOrder(1) = SortKeyOrderAscending
        End If

        .Action = ActionSort
    End With

End Sub
