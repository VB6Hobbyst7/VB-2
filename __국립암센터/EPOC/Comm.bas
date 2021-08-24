Attribute VB_Name = "Comm"
Option Explicit

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
Public gEquip As String
Public gIP As String
Public gAllExam     As String

Public gQCEquip As String

Public gEquipID As String

Type DB_Parm
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
    
'    Driver  As String
'    USER    As String
'    Passwd  As String
'    Server  As String
'    DB      As String
'    HostName    As String
'    LocalDB As String
End Type

Public gDB_Parm As DB_Parm

Public comState As String
Public comsignal As String
Public comSend As String

Public gOrderMessage As String
Public gOrderCnt As Integer
Public gNACKCnt As Integer
Public gPreMsg As String
Public gACKSig As Integer

Public gArrEquip() As String
Public gUserID As String


'-- Result S[read Column Seq Num
Public Const colSpecNo = 0     '미사용
Public Const colCHECKBOX = 1
Public Const colSAVESEQ = 2    '저장순번(날짜별)
Public Const colEXAMDATE = 3   '검사일자
Public Const colHOSPDATE = 4   '병원접수일자
Public Const colBARCODE = 5
Public Const colCHARTNO = 6
Public Const colPID = 7        '병록번호(내원번호)
Public Const colINOUT = 8      '입원/외래
Public Const colDISKNO = 9
Public Const colPOSNO = 10
Public Const colPNAME = 11
Public Const colPSEX = 12
Public Const colPAGE = 13
Public Const colOCNT = 14
Public Const colRCNT = 15
Public Const colState = 16

Public Const colEQUIPCODE = 1
Public Const colEXAMCODE = 2
Public Const colEXAMNAME = 3
Public Const colMachResult = 4
Public Const colRESULT = 5
Public Const colSeq = 6
Public Const colFLAG = 7
Public Const colSUBCODE = 8


'-- 수신한 오더정보
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

'-- 수신한 결과정보
Type IntfData
    SpcmNo   As String
    PatNo    As String
    BarNo    As String
    TESTCD   As String
    RackNo   As String
    TubePos  As String
    MnmCd    As String
    MnmNm    As String
    MCnt     As String
    Rst      As String
    SpcPos   As String
    RsltDate As String
    RsltSeq  As String
    OperatorID  As String
End Type

Public mResult As IntfData
Public gEquipCode   As String


'-- 오늘 검사한 날짜의 Max + 1 번호를 가져온다
Public Function getMaxTestNum(ByVal strDate As String) As Long

    getMaxTestNum = 1
    
    '-- 결과업데이트
          SQL = "SELECT MAX(SAVESEQ) as SEQ FROM PATRESULT  "
    SQL = SQL & " WHERE MID(EXAMDATE,1,8) = '" & strDate & "' " & vbCrLf
    
    res = GetDBSelectColumn(gLocal, SQL)
    
    If res > 0 Then
        If Trim(gReadBuf(0)) = "" Then
            getMaxTestNum = 1
        Else
            getMaxTestNum = Trim(gReadBuf(0)) + 1
        End If
    End If
    
    If getMaxTestNum >= 99999 Then
        getMaxTestNum = 99999
    End If
    
End Function

Public Function chrSTX() As String
    chrSTX = Chr(2)
End Function

Public Function chrETX() As String
    chrETX = Chr(3)
End Function

Public Function chrSOH() As String
    chrSOH = Chr(1)
End Function

Public Function chrEOT() As String
    chrEOT = Chr(4)
End Function

Public Function chrENQ() As String
    chrENQ = Chr(5)
End Function

Public Function chrACK() As String
    chrACK = Chr(6)
End Function

Public Function chrTAB() As String
    chrTAB = Chr(9)
End Function

Public Function chrLF() As String
    chrLF = Chr(10)
End Function

Public Function chrCR() As String
    chrCR = Chr(13)
End Function

Public Function chrNACK() As String
    chrNACK = Chr(15)
End Function

Public Function chrSPC() As String
    chrSPC = Chr(20)
End Function

Public Function chrETB() As String
    chrETB = Chr(23)
End Function

Function GetSetup() As Boolean
'---------------------------------------------------------------------------------------------------------------------
'                       Setup  File을 읽어온다.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 100

    db_tmp = ""

    GetSetup = False
    

    db_tmp = ""
    Call GetPrivateProfileString("config", "Equip", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquip = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "QCEquip", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gQCEquip = Trim(frmInterface.txtTemp)
    
''    db_tmp = ""
''    Call GetPrivateProfileString("config", "gPort", "", db_tmp, 100, App.Path & "\Interface.ini")
''    frmInterface.txtTemp = Trim(db_tmp)
''    gSetup.gPort = Trim(frmInterface.txtTemp)
''
''    db_tmp = ""
''    Call GetPrivateProfileString("config", "gSpeed", "", db_tmp, 100, App.Path & "\Interface.ini")
''    frmInterface.txtTemp = Trim(db_tmp)
''    gSetup.gSpeed = Trim(frmInterface.txtTemp)
''
''    db_tmp = ""
''    Call GetPrivateProfileString("config", "gParity", "", db_tmp, 100, App.Path & "\Interface.ini")
''    frmInterface.txtTemp = Trim(db_tmp)
''    gSetup.gParity = Trim(frmInterface.txtTemp)
''
''    db_tmp = ""
''    Call GetPrivateProfileString("config", "gDataBit", "", db_tmp, 100, App.Path & "\Interface.ini")
''    frmInterface.txtTemp = Trim(db_tmp)
''    gSetup.gDataBit = Trim(frmInterface.txtTemp)
''
''    db_tmp = ""
''    Call GetPrivateProfileString("config", "gStopBit", "", db_tmp, 100, App.Path & "\Interface.ini")
''    frmInterface.txtTemp = Trim(db_tmp)
''    gSetup.gStopBit = Trim(frmInterface.txtTemp)
''
''    db_tmp = ""
''    Call GetPrivateProfileString("config", "gStartBit", "", db_tmp, 100, App.Path & "\Interface.ini")
''    frmInterface.txtTemp = Trim(db_tmp)
''    gSetup.gStartBit = Trim(frmInterface.txtTemp)
''
''    db_tmp = ""
''    Call GetPrivateProfileString("config", "gRTSEnable", "", db_tmp, 100, App.Path & "\Interface.ini")
''    frmInterface.txtTemp = Trim(db_tmp)
''    gSetup.gRTSEnable = Trim(frmInterface.txtTemp)
''
''    db_tmp = ""
''    Call GetPrivateProfileString("config", "gDTREnable", "", db_tmp, 100, App.Path & "\Interface.ini")
''    frmInterface.txtTemp = Trim(db_tmp)
''    gSetup.gDTREnable = Trim(frmInterface.txtTemp)
''
''    db_tmp = ""
''    Call GetPrivateProfileString("Server", "ServerPath", "", db_tmp, 100, App.Path & "\Interface.ini")
''    frmInterface.txtTemp = Trim(db_tmp)
''    gServerPath = Trim(frmInterface.txtTemp)
        
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "OCS", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gOCS = Trim(frmInterface.txtTemp)
'
    '== 장비 관련 설정  ==============================================================================
    db_tmp = ""
    Call GetPrivateProfileString("config", "Equip", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquip = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "EquipCode", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquipCode = Trim(frmInterface.txtTemp)
    
    
    '== 통신 관련 설정  ==============================================================================
    db_tmp = ""
    Call GetPrivateProfileString("config", "gPort", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gPort = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gSpeed", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gSpeed = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gParity", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gParity = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gDataBit", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gDataBit = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gStopBit", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gStopBit = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gStartBit", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gStartBit = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "gRTSEnable", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gRTSEnable = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gDTREnable", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gDTREnable = Trim(frmInterface.txtTemp)
   
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "ComFormat", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gCOMFormat = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "ASTMFormat", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gASTMFormat = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "OPTVersion", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gOPTVersion = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "IFMode", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gIFMode = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "AutoSave", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gSave = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "IFScreen", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gScreen = Trim(frmInterface.txtTemp)



    '== DB 관련 설정    ==============================================================================
    Call GetPrivateProfileString("DATABASE", "dbtype", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.DBType = Trim(frmInterface.txtTemp)
    
    Call GetPrivateProfileString("DATABASE", "server", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.Server = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "uid", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.USER = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "pwd", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.Passwd = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "database", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.DB = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("Server", "IFUser", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gUserID = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Server", "LocalIP", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gIP = Trim(frmInterface.txtTemp)


    '-- osw 추가
'    Call GetPrivateProfileString("DRDATABASE", "dbtype", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDRDB_Parm.DBType = Trim(frmInterface.txtTemp)
'
'    Call GetPrivateProfileString("DRDATABASE", "server", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDRDB_Parm.Server = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("DRDATABASE", "uid", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDRDB_Parm.USER = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("DRDATABASE", "pwd", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDRDB_Parm.Passwd = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("DRDATABASE", "database", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDRDB_Parm.DB = Trim(frmInterface.txtTemp)
    
    '==  Winsock 관련 설정    ==============================================================================
'    db_tmp = ""
'    Call GetPrivateProfileString("CONFIG", "ServerIP", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDRDB_Parm.ServerIP = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("CONFIG", "ServerPort", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDRDB_Parm.ServerPort = Trim(frmInterface.txtTemp)
        
    '== DB Table 관련 설정    ==============================================================================
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "ORDTABLE", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBTBL_Parm.ORDTABLE = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "RSLTTABLE", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBTBL_Parm.RSLTTABLE = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "MSTTABLE", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBTBL_Parm.MSTTABLE = Trim(frmInterface.txtTemp)
'
'    '== DB Table Column 관련 설정    =======================================================================
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "ORDDATE", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBCOLUMN_Parm.ORDDATE = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "ORDDATE", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBCOLUMN_Parm.ORDDATE = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "RSLTDATE", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBCOLUMN_Parm.RsltDate = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "BARCODE", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBCOLUMN_Parm.BARCODE = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "PID", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBCOLUMN_Parm.PID = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "PNAME", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBCOLUMN_Parm.PName = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "PSEX", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBCOLUMN_Parm.pSex = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "PAGE", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBCOLUMN_Parm.Page = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "TESTCD", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBCOLUMN_Parm.TESTCD = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "RESULT", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBCOLUMN_Parm.RESULT = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "INTRESULT", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBCOLUMN_Parm.INTRESULT = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "STATUS", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBCOLUMN_Parm.Status = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "JUDGE", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBCOLUMN_Parm.JUDGE = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "MACHCODE", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBCOLUMN_Parm.MACHCD = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Order", "USER", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gDBCOLUMN_Parm.USER = Trim(frmInterface.txtTemp)
  
    '-- 검사파트
'    db_tmp = ""
'    Call GetPrivateProfileString("CONFIG", "GumPart", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gGumPart = Trim(frmInterface.txtTemp)
'
'    '== 지누스 DLL 서비스 관련 설정    =======================================================================
'    db_tmp = ""
'    Call GetPrivateProfileString("SERVICE", "URL", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gGINUS_Parm.url = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("SERVICE", "SVC", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gGINUS_Parm.SVC = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("SERVICE", "HCD", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gGINUS_Parm.HCD = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("SERVICE", "MCD", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gGINUS_Parm.MCD = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Server", "ORDER", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gOrderPath = Trim(frmInterface.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Server", "RESULT", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gResultPath = Trim(frmInterface.txtTemp)
'
    db_tmp = ""
    Call GetPrivateProfileString("Server", "IFUser", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gIFUser = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("Server", "ServerPath", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gServerPath = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Server", "Panel", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gPanel = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Server", "Doctor", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDoctorID = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Server", "Department", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDepartment = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Server", "HostSN", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gHostSN = Trim(frmInterface.txtTemp)

            
            
    GetSetup = True

End Function

Public Sub SetSQLData(ByVal strName As String, ByVal argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
        
    If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.Path & "\Log")
    End If
    
    sFileName = Format(CDate(frmInterface.dtpToday), "yyyy-mm-dd") & "_" & strName
    
    Open App.Path & "\Log\" & sFileName & ".txt" For Append As FilNum
    Print #FilNum, Format(Time, "hh:nn:ss") & " " & argSQL
    Close FilNum
    
End Sub


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
Public Sub Save_Raw_Data(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    
    If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.Path & "\Log")
    End If
    
    sFileName = Format(CDate(frmInterface.Text_Today.Text), "yyyymmdd")
    
    Open App.Path & "\Log\" & sFileName & ".txt" For Append As FilNum
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

