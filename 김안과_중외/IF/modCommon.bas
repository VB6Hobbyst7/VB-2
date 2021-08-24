Attribute VB_Name = "modCommon"
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long
    
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
                    
'########## 병원정보 관련 ############################
Type HospParameter
    HOSPCD      As String
    HOSPNM      As String
    LABCD       As String
    LABNM       As String
    PARTCD      As String
    PARTNM      As String
    MACHCD      As String
    MACHNM      As String
    USERID      As String
    USERPW      As String
    USERNM      As String
    LOGINYN     As String
    SAVEPW      As String
    BARUSE      As String
    SAVEAUTO    As String
    SAVELIS     As String
    RSTTYPE     As String
    QCPATH      As String
    APPNM       As String
    LOQWRITE    As String
End Type

Public gHOSP        As HospParameter
'########## 병원정보 관련 ############################

'########## 통신정보 관련 (시리얼/소켓) ##############
Type ComParameter
    COMTYPE     As String
    COMPORT     As String
    SPEED       As String
    DATABIT     As String
    STARTBIT    As String
    STOPBIT     As String
    Parity      As String
    RTSEnable   As Boolean
    DTREnable   As Boolean
    TCPTYPE     As String
    TCPIP       As String
    TCPPORT     As String
End Type

Public gComm        As ComParameter
'########## 통신정보 관련 ############################

Public gOCS         As String
Public strSetup     As String * 100
Public strSetUp1    As String

Public gArrEQP()    As String
Public gAllTestCd   As String   '-- 인터페이스에 등록된 전체검사코드
Public gAllOrdCd    As String   '-- 인터페이스에 등록된 전체오더코드
Public gPatOrdCd    As String   '-- 검체별 검사코드
Public gRow         As Long     '-- 작업중 Row

Public gCENXPCD     As String
Public gADV18CD     As String

Declare Function GetDefaultCommConfig Lib "kernel32" Alias "GetDefaultCommConfigA" (ByVal lpszName As String, lpCC As COMMCONFIG, lpdwSize As Long) As Long

Type DCB
        DCBlength As Long
        BaudRate As Long
        fBitFields As Long
        wReserved As Integer
        XonLim As Integer
        XoffLim As Integer
        ByteSize As Byte
        Parity As Byte
        StopBits As Byte
        XonChar As Byte
        XoffChar As Byte
        ErrorChar As Byte
        EofChar As Byte
        EvtChar As Byte
        wReserved1 As Integer
End Type

Type COMMCONFIG
    dwSize As Long
    wVersion As Integer
    wReserved As Integer
    dcbx As DCB
    dwProviderSubType As Long
    dwProviderOffset As Long
    dwProviderSize As Long
    wcProviderData As Byte
End Type

Public gCOLWIDTH         As String

Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long


Sub Main()
    
    '두번 실행 하지 않음
    If App.PrevInstance Then
       MsgBox "프로그램이 이미 실행중입니다.!", vbExclamation
       End
    End If
    
    '-- INI 정보
    Call GetExeVersion
    
    '-- INI 정보
    Call GetSetup
    
    '-- 병원코드
    If Len(gHOSP.HOSPCD) = 0 Then
        frmHospInfo.Show vbModal
    End If
    
    
    If Len(gLocalDB.PATH) = 0 Then
        frmDB_Local.Show vbModal
    End If
    
    If gDBTYPE = "1" Then
        If Len(gORADB.SID) = 0 Then
            frmDB_Oracle.Show vbModal
        End If
    ElseIf gDBTYPE = "2" Then
        If Len(gSQLDB.IP) = 0 Then
            frmDB_MSSQL.Show vbModal
        End If
    Else
        MsgBox App.PATH & "\OKSOFT.ini 파일에서" & vbNewLine & vbNewLine & "DBTYPE을 먼저 설정하세요 ", vbOKOnly + vbInformation, "DB TYPE 설정"
        End
    End If
    
    
    '-- 로컬 DB 접속
    If Not DbConnect_Local Then
        If vbYes = MsgBox("로컬 데이터베이스가 없습니다. 찾으시겠습니까? ", vbCritical + vbYesNo) Then
            frmDB_Local.Show vbModal
        Else
            End
        End If
    Else
        cn_Local_Flag = True
    End If
       
    If gDBTYPE = "1" Then
        '-- ORACLE DB 접속
        If Not DbConnect_ORACLE Then
            If vbYes = MsgBox("오라클 데이터베이스가 없습니다. 등록하시겠습니까? ", vbCritical + vbYesNo) Then
                frmDB_Oracle.Show vbModal
            Else
                End
            End If
        Else
            cn_Server_Flag = True
        End If
    ElseIf gDBTYPE = "2" Then
        '-- MSSQL DB 접속
        If Not DbConnect_SQL Then
            If vbYes = MsgBox("MS-SQL 데이터베이스가 없습니다. 등록하시겠습니까? ", vbCritical + vbYesNo) Then
                frmDB_MSSQL.Show vbModal
            Else
                End
            End If
        Else
            cn_Server_Flag = True
        End If
        
        
''        If Not DbConnect_SQL_QC Then
''            MsgBox "QC 데이터베이스가 없습니다.", vbCritical, "QC"
''        End If
        
    Else
        MsgBox "데이터베이스 설정을 확인하세요.", vbCritical, "데이터베이스 설정"
        End
    End If
    
    '-- 로그인 사용자
    
    '-- 컨트롤초기화
    Call CtlInitializing
    
    '-- 장비통신정보 읽어오기
        
    '-- 장비검사정보 읽어오기(검사항목 리스트업)
'    Call getTestNms(mEqpCd)
    
    '-- 포트 열기





'1. Read Ini
'   - 서버정보
'   - 장비통신정보
'   - 사용자정보
'
'
'5. Local DB Open
'
'6. Server DB Open(1)
'   Server DB Open(2)
'
'4. Test List Get
'   Server(1) Test Get
'   Server(2) Test Get
'
'3. Control Initial
'   - Default Set
'
'2. Spread Set
'   - Column View Get
'   - Column View Set
'
'4. Communication Open
'
'7. Error Handling
'
'8.
    If gHOSP.LOGINYN = "Y" Then
        Call frmLogin.Show
    Else
        Call frmMain.Show
    End If
    
End Sub


Public Sub CtlInitializing()
         
    frmMain.frame1.ZOrder 0

End Sub

Public Sub SetMenu()
    
    With frmMain
        '-- 바코드사용
        If gHOSP.BARUSE = "Y" Then
            .mnuBarcode.Checked = True
            .mnuSeqno.Checked = False
            .optBarSeq(0).Value = True
        Else
            .mnuBarcode.Checked = False
            .mnuSeqno.Checked = True
            .optBarSeq(1).Value = True
        End If
        
        '-- 결과전송
        If gHOSP.SAVEAUTO = "Y" Then
            .mnuSaveAuto.Checked = True
            .mnuSaveManual.Checked = False
            .optTrans(0).Value = True
        Else
            .mnuSaveAuto.Checked = False
            .mnuSaveManual.Checked = True
            .optTrans(1).Value = True
        End If
        
        '-- 적용결과
        If gHOSP.SAVELIS = "Y" Then
            .mnuLisResult.Checked = True
            .mnuEqpResult.Checked = False
            .optSaveResult(1).Value = True
        Else
            .mnuLisResult.Checked = False
            .mnuEqpResult.Checked = True
            .optSaveResult(0).Value = True
        End If
        
        
    End With
    
    
End Sub

Public Sub GetExeVersion()
    
    '-- HOSPITAl INFO GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("EXE", "APPNAME", "", strSetup, 100, App.PATH & "\OKSOFT.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.APPNM = Trim(strSetUp1)
    
End Sub

Public Sub GetSetup()
    
    '-- HOSPITAl INFO GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "HOSPCD", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.HOSPCD = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "HOSPNM", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.HOSPNM = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "LABCD", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.LABCD = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "LABNM", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.LABNM = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "PARTCD", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.PARTCD = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "PARTNM", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.PARTNM = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "MACHCD", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.MACHCD = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "MACHNM", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.MACHNM = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "USERID", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.USERID = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "USERPW", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.USERPW = Trim(strSetUp1)
        
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "USERNM", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.USERNM = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "LOGINYN", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.LOGINYN = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "SAVEPW", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.SAVEPW = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "BARUSE", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.BARUSE = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "SAVELIS", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.SAVELIS = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "SAVEAUTO", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.SAVEAUTO = Trim(strSetUp1)
    
    '-- 바코드 미사용시 결과받는 형태
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "RSTTYPE", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.RSTTYPE = Trim(strSetUp1)
    
    '-- QC결과 저장경로
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "QCPATH", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.QCPATH = Trim(strSetUp1)
    
    '-- ADVIA1800-2 장비코드
'    strSetup = "":    strSetUp1 = ""
'    Call GetPrivateProfileString("HOSP", "ADVIA1800", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
'    strSetUp1 = Trim(strSetup)
'    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'    gADV18CD = Trim(strSetUp1)
'
'    '-- CENTAURXP 장비코드
'    strSetup = "":    strSetUp1 = ""
'    Call GetPrivateProfileString("HOSP", "CENTAURXP", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
'    strSetUp1 = Trim(strSetup)
'    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'    gCENXPCD = Trim(strSetUp1)
    
    '-- LOG 기록여부
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "LOGWRITE", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.LOQWRITE = Trim(strSetUp1)
    '-- HOSPITAl INFO GET END
    
    '-- OCS
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "OCS", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gOCS = Trim(strSetUp1)
    
    '-- DB TYPE GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "DBTYPE", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gDBTYPE = Trim(strSetUp1)
    
    '-- LOCAL DB GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "LOCALPATH", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gLocalDB.PATH = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "LOCALUID", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gLocalDB.UID = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "LOCALPWD", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gLocalDB.PWD = Trim(strSetUp1)

    '-- ORACLE DB GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "ORACLESID", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gORADB.SID = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "ORACLEUID", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gORADB.UID = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "ORACLEPWD", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gORADB.PWD = Trim(strSetUp1)

    '-- MSSQL DB GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "MSSQLIP", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB.IP = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "MSSQLDB", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB.DB = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "MSSQLUID", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB.UID = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "MSSQLPWD", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB.PWD = Trim(strSetUp1)

    '-- MSSQL QC DB GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "MSSQLIP_QC", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB_QC.IP = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "MSSQLDB_QC", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB_QC.DB = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "MSSQLUID_QC", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB_QC.UID = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "MSSQLPWD_QC", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB_QC.PWD = Trim(strSetUp1)


    '-- COMM INFO GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "COMTYPE", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.COMTYPE = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "COMPORT", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.COMPORT = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "SPEED", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.SPEED = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "PARITY", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.Parity = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "DATABIT", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.DATABIT = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "STARTBIT", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.STARTBIT = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "STOPBIT", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.STOPBIT = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "RTSEnable", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.RTSEnable = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "DTREnable", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.DTREnable = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "TCPTYPE", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.TCPTYPE = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "TCPIP", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.TCPIP = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "TCPPORT", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.TCPPORT = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("VIEW", "COLWIDTH", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gCOLWIDTH = Trim(strSetUp1)
    
End Sub

Public Sub SetExamCode()
    Dim i As Integer
    
    
    With frmMain.spdOrder
        .MaxCols = colSTATE + UBound(gArrEQP)
        For i = 0 To UBound(gArrEQP) - 1
            .Col = colSTATE + (i + 1)
            .Row = -1
            .CellType = CellTypeEdit
            .TypeEditCharSet = TypeEditCharSetASCII
            .TypeEditCharCase = TypeEditCharCaseSetNone
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            Call SetText(frmMain.spdOrder, Trim(gArrEQP(i + 1, 5)), 0, colSTATE + (i + 1))    '-- 5 : 약어명
            .ColWidth(colSTATE + (i + 1)) = gCOLWIDTH
        Next
    End With
    
    With frmMain.spdROrder
        .MaxCols = colSTATE + UBound(gArrEQP)
        For i = 0 To UBound(gArrEQP) - 1
            .Col = colSTATE + (i + 1)
            .Row = -1
            .CellType = CellTypeEdit
            .TypeEditCharSet = TypeEditCharSetASCII
            .TypeEditCharCase = TypeEditCharCaseSetNone
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            Call SetText(frmMain.spdROrder, Trim(gArrEQP(i + 1, 5)), 0, colSTATE + (i + 1))    '-- 5 : 약어명
            .ColWidth(colSTATE + (i + 1)) = gCOLWIDTH
        Next
    End With
    
End Sub

' Function : SetSpreadSort
' Author   : 서은아빠(http://cafe.naver.com/xlsvba)
' LA Time  : 2009-12-29 23:24
' Purpose  : Farpoint Spread8.0 내의 지정 스프레드의 정렬을 설정한다.
'            Farpoint Spread8.0 디자이너라면 Columns and Rows 메뉴에서 Sort Indicate에서 설정하믄
'            되지만 이미 사용중에 다수의 스프레드 변경이라믄 이 방법두 나쁘지 않아 보인다.
' Param    : SP - 스프레드명, iSortoption - Sort Option
'   0(ColUserSortIndicatorNone) - None (Default) No pointer appears
'           No sorting occurred. The BeforeUserSort and AfterUserSort events did not occur.
'   1(ColUserSortIndicatorAscending) - Ascending The  pointer appears when the column is sorted
'           Ascending sort occurred. The BeforeUserSort and AfterUserSort events occurred.
'   2(ColUserSortIndicatorDescending) - Descending The  pointer appears when the column is sorted
'           Descending sort occurred. The BeforeUserSort and AfterUserSort events occurred.
'   3(ColUserSortIndicatorDisabled) - Disabled No pointer appears
'           No sorting can occur. The BeforeUserSort and AfterUserSort events did not occur
'=========================================================================
Public Sub SetSpreadSort(SP As Object, Optional ByVal iSortoption As Integer = 0)

    Dim i As Integer
    
    '## Setting Sort Indicate
    For i = 1 To SP.MaxCols
        SP.ColUserSortIndicator(i) = iSortoption
    Next
    
    SP.UserColAction = UserColActionSort
   
End Sub

Public Function EnumSerPorts(port As Integer) As Long
    Dim cc As COMMCONFIG, ccsize As Long
    ccsize = LenB(cc)
    EnumSerPorts = GetDefaultCommConfig("COM" + Trim(Str(port)) + Chr(0), cc, ccsize)
End Function
