Attribute VB_Name = "modCommon"
Option Explicit


'-- GINUS DLL
'Public Declare Function W2ACALL2 Lib "c:\windows\system32\w2afun.dll" (ByVal sSVC As String, ByVal sRequest As String, ByVal sURL As String) As String

'-- 고대병원 DLL
Public Declare Function TuxedoInit_V2 Lib "C:\DEV\DLL\P_SLDLL.DLL" Alias "TuxedoInit_V2A" (ByVal in_usrname As String, ByVal in_cltname As String, ByVal in_svrid As String) As Variant

'Declare Function TuxedoInit_V2 Lib "P_SLDLL.dll" Alias "TuxedoInit_V2A" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any) As Long
    
    
Public Declare Function UserChk Lib "C:\DEV\DLL\P_SLDLL.DLL" (ByVal in_userid As String, ByVal in_pass As String, ByVal in_locate As String, ByRef out_usernm As Variant) As Integer
Public Declare Function AutoAccept_poct Lib "C:\DEV\DLL\P_SLDLL.DLL" (ByVal in_spcid As String, ByVal in_userid As String, ByVal out_msg As String) As Integer
Public Declare Function ExaminfoList Lib "C:\DEV\DLL\P_SLDLL.DLL" (ByVal in_Flagas As String, ByVal in_spcid As String, ByVal in_execdate As String, ByVal out_order As Variant) As Integer
Public Declare Function ResultList Lib "C:\DEV\DLL\P_SLDLL.DLL" (ByVal out_jobsect As String, ByVal out_Userid As String, ByVal out_Result As Variant, ByVal out_eqipcd As String, ByVal out_autoryn As String) As Integer

'Public Declare Function TuxedoInit_V2 Lib "C:\WINDOWS\SYSTEM32\P_SLDLL.dll" (ByVal in_usrname As String, ByVal in_cltname As String, ByVal in_svrid As String) As Integer
'Public Declare Function UserChk Lib "C:\WINDOWS\SYSTEM32\P_SLDLL.dll" (ByVal in_userid As String, ByVal in_pass As String, ByVal in_locate As String, ByRef out_usernm As Variant) As Integer
'Public Declare Function AutoAccept_poct Lib "C:\WINDOWS\SYSTEM32\P_SLDLL.dll" (ByVal in_spcid As String, ByVal in_userid As String, ByRef out_msg As String) As Integer
'Public Declare Function ExaminfoList Lib "C:\WINDOWS\SYSTEM32\P_SLDLL.dll" (ByVal in_Flagas As String, ByVal in_spcid As String, ByVal in_execdate As String, ByRef out_order As Variant) As Integer
'Public Declare Function ResultList Lib "C:\WINDOWS\SYSTEM32\P_SLDLL.dll" (ByVal out_jobsect As String, ByVal out_Userid As String, ByVal out_Result As Variant, ByRef out_eqipcd As String, ByRef out_autoryn As String) As Integer


'1. function TuxedoInit_V2(in_usrname,in_cltname, in_svrid: pChar): Integer; StdCall; external 'C:\DEV\DLL\P_SLDLL.dll';
'   -    DR 환경 구성에 따른 호출 인자 추가 (사용자 로그인 창에 추가 요청)
'   -    (정상운영시 ‘01’, DR 상황 발생시 ‘DR’
'   - in_usrname, in_cltname null로 보내도 됨
'     예) if TuxedoInit_V2('','','01')  = 1 이면 TMAX 연결 OK
'
'?/2.  function UserChk(in_userid, in_pass, in_locate : string ; var out_usernm: variant): Integer; StdCall; external 'C:\DEV\DLL\P_SLDLL.dll';
'   -    in_userid <- 사번, in_pass <- 패스워드, in_locate <- 병원구분(AA,GR,AS로 입력: AA(안암),GR(구로),AS(안산)
'   -  Return 값이 1이면, out_usernm 에 사용자 이름
'      (5분마다 타이머로 UserChk 함수 호출로 세션 끊김 방지 처리)
'
'
'//3. function AutoAccept_poct(in_spcid, in_userid: String; var out_msg :string):integer; StdCall 'C:\DEV\DLL\P_SLDLL.dll';
'   - 검체를 채취(B) 상태에서 결과입력상태(E)로 변경
'   -  AutoAccept_POCT('102020202',’88888’, out_msg);
'                      ----------
'                       검체번호
'       성공 시 --> return value  1 & out_msg = 'OK'
'       실패 시 --> return value -1 & out_msg 는 에러 관련 메시지
'
'
'오더 4.  function ExaminfoList(in_Flag, in_spcid, in_execdate: string ; var out_order: variant) : Integer; StdCall;  external 'C:\DEV\DLL\P_SLDLL.dll';
'
'    -  ExaminfoList('', '1231231231', '20080908', out_msg);
'
'       성공적으로 조회 시 out_msg 다음과 같은 형식으로 보내줌 , return value 는 검사 건수
'       구분자를 '|' 로 구분함.
'       examinfo.sPatno   [ii] + '|' +    --> 환자 등록번호
'       examinfo.sPatname [ii] + '|' +    --> 환자 명
'       examinfo.sSex     [ii] + '|' +    --> 성별 (F,M)
'       examinfo.sAge     [ii] + '|' +    --> 나이
'       examinfo.sOrdseqno[ii] + '|' +    --> 처방순번
'       examinfo.sWorkno  [ii] + '|' +    --> 작업번호
'       examinfo.sOrddate [ii] + '|' +    --> 처방일자(yyyy-mm-dd')
'       examinfo.sExamcode[ii] + '|' +    --> 검사코드
'       examinfo.sAcptdate[ii] + '|' +    --> 접수일자
'       examinfo.sAreano  [ii] + '|' +    --> 접수번호
'       examinfo.sMeddept [ii] + '|' +    --> 진료과
'       examinfo.sSpccode [ii] + '|' +    --> 검체코드
'       examinfo.sSpcname [ii];           --> 검체명
'
'       예)
'       00001234|홍길동|F|62|2001|1|2008-09-08|BG2200|2008-09-08|1|OG|101|Whole Blood
'
'저장 5. function ResultList(out_jobsect, out_Userid:string; out_Result : variant ; out_eqipcd: string='';out_autoryn: string='N') : Integer; StdCall;
'       ResultList('3', '88888', out_Result, 'EQIP_1', 'N');
'       '2'     --> 보고처리 ( 결과만 입력인 경우에는 '3' 으로 처리)
'       '99999' --> login 사번
'       'EQIP_1' --> 검사 실시한 장비코드
'       out_result 처리 방법
'       검사결과를 '|' 로 구분하여 Upload 처리
'      검체번호|등록번호|처방일자|처방순번|검사코드|검사결과|검사특기시항|delta|Panic|정상치flag|장비flag|
'       **       **       **       **       **        **
'       순으로 보내줌 ( ** 는 필수 입력사항임)
'
'       예)
'       8080004709|80000412|2008-07-28|2001|BG23201A01|2|||||
'       8080004709|80000412|2008-07-28|2001|BG23201A02|2|||||
'       8080004709|80000412|2008-07-28|2001|BG23201A03|0|||||

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

'-- 통신형태
Public gCOMFormat  As String

'-- ASTM 형태
Public gASTMFormat  As String

'-- 장비 S/W Version
Public gOPTVersion  As String

Public gSetup       As config
Public gPart        As String
Public gGubun       As Integer
Public gEquip       As String
Public gEquipCode   As String

Public gIP          As String
Public gOrderExam   As String
Public gAllExam     As String
Public gOrder       As String

Public gSndState    As String
Public gRecodeType  As String

Public gQCEquip     As String
Public gPreSpecID   As String
Public gPreRow      As Long
Public gOrdRow      As Long
Public gEquipID     As String

Public gCurMsgCnt   As String

Public gHeader      As String
Public gPatient     As String
Public gMsgEnd      As String

'-- Origin DB
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

'-- BackUp DB
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

'-- Table Information
Type DBTBL_Parm
    ORDTABLE As String
    RSLTTABLE As String
    MSTTABLE As String
End Type

Public gDBTBL_Parm As DBTBL_Parm

'-- Table Column Information
Type DBCOLUMN_Parm
    ORDDATE As String
    RSLTDATE As String
    BARCODE As String
    PID As String
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

'-- 지누스 DLL 서비스 Information
Type GINUS_Parm
    URL As String
    SVC As String
    HCD As String
End Type

Public gGINUS_Parm As GINUS_Parm

'-- User ID
Public gUserID As String

'Public comState As String
'Public comsignal As String
'Public comSend As String
'
'Public gOrderMessage As String
'Public gOrderCnt As Integer
'Public gNACKCnt As Integer
'Public gPreMsg As String
'Public gACKSig As Integer
Public gIFUser As String


'Public gArrEquip() As String

'-- Result S[read Column Seq Num
Public Const colSpecNo = 0  '미사용
Public Const colCheckBox = 1
Public Const colSeqNo = 2
Public Const colOrdDate = 3
Public Const colBarcode = 4
Public Const colRack = 5
Public Const colPos = 6
Public Const colPID = 7
Public Const colPName = 8
Public Const colSex = 9
Public Const colAge = 10
Public Const colOCnt = 11
Public Const colRCnt = 12
Public Const colState = 13

'Public Const colA1c = 13
'Public Const colIFCC = 15
'Public Const coleAg = 17



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
End Type

Public mResult As IntfData

       
       
'Type MicroDic
'    MicrosCnt        As Integer
'    MicroRst         As String
'End Type
'
'Public mMicro As MicroDic

'Public gComment_All As String
'Public gComment_Code As String


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
    Call GetPrivateProfileString("config", "Equip", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gEquip = Trim(frmLogin.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "EquipCode", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gEquipCode = Trim(frmLogin.txtTemp)
    
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "QCEquip", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmlogin.txtTemp = Trim(db_tmp)
'    gQCEquip = Trim(frmlogin.txtTemp)
    
    '== 통신 관련 설정  ==============================================================================
    db_tmp = ""
    Call GetPrivateProfileString("config", "gPort", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gSetup.gPort = Trim(frmLogin.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gSpeed", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gSetup.gSpeed = Trim(frmLogin.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gParity", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gSetup.gParity = Trim(frmLogin.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gDataBit", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gSetup.gDataBit = Trim(frmLogin.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gStopBit", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gSetup.gStopBit = Trim(frmLogin.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gStartBit", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gSetup.gStartBit = Trim(frmLogin.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "gRTSEnable", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gSetup.gRTSEnable = Trim(frmLogin.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gDTREnable", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gSetup.gDTREnable = Trim(frmLogin.txtTemp)
   
    db_tmp = ""
    Call GetPrivateProfileString("config", "ComFormat", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gCOMFormat = Trim(frmLogin.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "ASTMFormat", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gASTMFormat = Trim(frmLogin.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "OPTVersion", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gOPTVersion = Trim(frmLogin.txtTemp)
    

    '== DB 관련 설정    ==============================================================================
    Call GetPrivateProfileString("DATABASE", "dbtype", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDB_Parm.DBType = Trim(frmLogin.txtTemp)
    
    Call GetPrivateProfileString("DATABASE", "server", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDB_Parm.Server = Trim(frmLogin.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "uid", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDB_Parm.USER = Trim(frmLogin.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "pwd", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDB_Parm.Passwd = Trim(frmLogin.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "database", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDB_Parm.DB = Trim(frmLogin.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("Server", "IFUser", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gUserID = Trim(frmLogin.txtTemp)

    '-- osw 추가
    Call GetPrivateProfileString("DRDATABASE", "server", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDRDB_Parm.Server = Trim(frmLogin.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DRDATABASE", "uid", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDRDB_Parm.USER = Trim(frmLogin.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DRDATABASE", "pwd", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDRDB_Parm.Passwd = Trim(frmLogin.txtTemp)
    
    '==  Winsock 관련 설정    ==============================================================================
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "ServerIP", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDRDB_Parm.ServerIP = Trim(frmLogin.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "ServerPort", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDRDB_Parm.ServerPort = Trim(frmLogin.txtTemp)
        
    '== DB Table 관련 설정    ==============================================================================
    db_tmp = ""
    Call GetPrivateProfileString("Order", "ORDTABLE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBTBL_Parm.ORDTABLE = Trim(frmLogin.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("Order", "RSLTTABLE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBTBL_Parm.RSLTTABLE = Trim(frmLogin.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "MSTTABLE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBTBL_Parm.MSTTABLE = Trim(frmLogin.txtTemp)
        
    '== DB Table Column 관련 설정    =======================================================================
    db_tmp = ""
    Call GetPrivateProfileString("Order", "ORDDATE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.ORDDATE = Trim(frmLogin.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "ORDDATE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.ORDDATE = Trim(frmLogin.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "RSLTDATE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.RSLTDATE = Trim(frmLogin.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "BARCODE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.BARCODE = Trim(frmLogin.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "PID", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.PID = Trim(frmLogin.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "PNAME", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.PNAME = Trim(frmLogin.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "PSEX", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.PSEX = Trim(frmLogin.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "PAGE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.PAGE = Trim(frmLogin.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "TESTCD", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.TESTCD = Trim(frmLogin.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "RESULT", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.RESULT = Trim(frmLogin.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Order", "INTRESULT", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.INTRESULT = Trim(frmLogin.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Order", "STATUS", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.STATUS = Trim(frmLogin.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Order", "JUDGE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.JUDGE = Trim(frmLogin.txtTemp)
  
    db_tmp = ""
    Call GetPrivateProfileString("Order", "MACHCODE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.MACHCD = Trim(frmLogin.txtTemp)
  
    db_tmp = ""
    Call GetPrivateProfileString("Order", "USER", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.USER = Trim(frmLogin.txtTemp)
  
    '== 지누스 DLL 서비스 관련 설정    =======================================================================
    db_tmp = ""
    Call GetPrivateProfileString("SERVICE", "URL", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gGINUS_Parm.URL = Trim(frmLogin.txtTemp)
  
    db_tmp = ""
    Call GetPrivateProfileString("SERVICE", "SVC", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gGINUS_Parm.SVC = Trim(frmLogin.txtTemp)
  
    db_tmp = ""
    Call GetPrivateProfileString("SERVICE", "HCD", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gGINUS_Parm.HCD = Trim(frmLogin.txtTemp)
  
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

Public Sub SpreadSheetSort(ByRef Spread As vaSpread, ByVal Col As Integer, Optional ByVal SortType As Integer = 1)
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
