Attribute VB_Name = "modCommon"
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

'��ż���
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

'-- OCS ��ü
Public gOCS         As String

'-- �������
Public gCOMFormat   As String

'-- ASTM ����
Public gASTMFormat  As String

'-- ��� S/W Version
Public gOPTVersion  As String

Public gSetup       As config
Public gPart        As String
Public gGubun       As Integer
Public gEquip       As String
Public gEquipCode   As String

Public gIP          As String
Public gOrderExam   As String
Public gAllExam     As String
Public gAllExam_Bit As String
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

Public gGumPart     As String

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
    RsltDate As String
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


Public gArrEquip() As String

'-- Result S[read Column Seq Num
Public Const colSpecNo = 0     '�̻��
Public Const colCheckBox = 1
Public Const colSAVESEQ = 2    '�������(��¥��)
Public Const colEXAMDATE = 3   '�˻�����
Public Const colHOSPDATE = 4   '������������
Public Const colBARCODE = 5
Public Const colCHARTNO = 6
Public Const colPID = 7        '���Ϲ�ȣ(������ȣ)
Public Const colINOUT = 8      '�Կ�/�ܷ�
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


'-- ������ ��������
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

'-- ������ �������
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
End Type

Public mResult As IntfData

Public gSave   As String
Public gIFMode As String
Public gScreen As String

'-- ������ DLL ========================================================================================================================================
Type GINUS_Parm
    URL As String
    SVC As String
    HCD As String
    MCD As String
End Type

Public gGINUS_Parm As GINUS_Parm
Public Declare Function W2ACALL2 Lib "c:\windows\system32\w2afun.dll" (ByVal sSVC As String, ByVal sRequest As String, ByVal sURL As String) As String
'-- ������ DLL ========================================================================================================================================


Global Const gDept_Code As String = "06"

'���񽺸� ��ҹ��� �����Ѵ�!!!
Public UPLOAD_SVC As String '= "HAMA0111"
Public DWLOAD_SVC As String '= "HAMA0112"
Public LOGIN_SVC  As String '= "HAMA0125"
 
Global strbuf As tuxbuf
Global rcvbuf As tuxbuf

'Type MicroDic
'    MicrosCnt        As Integer
'    MicroRst         As String
'End Type
'
'Public mMicro As MicroDic

'Public gComment_All As String
'Public gComment_Code As String


'=================================
Global Const STATLIN = 167772165


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
'                       Setup  File�� �о�´�.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 100

    db_tmp = ""

    GetSetup = False
    
'    Dim strEqpCd As String
'    Dim GumEqpCd As String
'
'    strEqpCd = "C2411"
'
'    db_tmp = ""
'    Call GetPrivateProfileString("EQPCD", strEqpCd, "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    strEqpCd = frmInterface.txtTemp
                   
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "OCS", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gOCS = Trim(frmInterface.txtTemp)
    
    '== ��� ���� ����  ==============================================================================
    db_tmp = ""
    Call GetPrivateProfileString("config", "Equip", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquip = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "EquipCode", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquipCode = Trim(frmInterface.txtTemp)
    
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "QCEquip", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gQCEquip = Trim(frmInterface.txtTemp)
    
    '== ��� ���� ����  ==============================================================================
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
   
    db_tmp = ""
    Call GetPrivateProfileString("config", "ComFormat", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gCOMFormat = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "ASTMFormat", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gASTMFormat = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "OPTVersion", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gOPTVersion = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "IFMode", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gIFMode = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "AutoSave", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSave = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "IFScreen", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gScreen = Trim(frmInterface.txtTemp)



    '== DB ���� ����    ==============================================================================
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

    '-- osw �߰�
    Call GetPrivateProfileString("DRDATABASE", "server", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDRDB_Parm.Server = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DRDATABASE", "uid", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDRDB_Parm.USER = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DRDATABASE", "pwd", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDRDB_Parm.Passwd = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DRDATABASE", "database", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDRDB_Parm.DB = Trim(frmInterface.txtTemp)
    
    '==  Winsock ���� ����    ==============================================================================
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "ServerIP", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDRDB_Parm.ServerIP = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "ServerPort", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDRDB_Parm.ServerPort = Trim(frmInterface.txtTemp)
        
    '== DB Table ���� ����    ==============================================================================
    db_tmp = ""
    Call GetPrivateProfileString("Order", "ORDTABLE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBTBL_Parm.ORDTABLE = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("Order", "RSLTTABLE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBTBL_Parm.RSLTTABLE = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "MSTTABLE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBTBL_Parm.MSTTABLE = Trim(frmInterface.txtTemp)
        
    '== DB Table Column ���� ����    =======================================================================
    db_tmp = ""
    Call GetPrivateProfileString("Order", "ORDDATE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.ORDDATE = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "ORDDATE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.ORDDATE = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "RSLTDATE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.RsltDate = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "BARCODE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.BARCODE = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "PID", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.PID = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "PNAME", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.PNAME = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "PSEX", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.PSEX = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "PAGE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.PAGE = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "TESTCD", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.TESTCD = Trim(frmInterface.txtTemp)
        
    db_tmp = ""
    Call GetPrivateProfileString("Order", "RESULT", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.RESULT = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Order", "INTRESULT", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.INTRESULT = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Order", "STATUS", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.STATUS = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Order", "JUDGE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.JUDGE = Trim(frmInterface.txtTemp)
  
    db_tmp = ""
    Call GetPrivateProfileString("Order", "MACHCODE", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.MACHCD = Trim(frmInterface.txtTemp)
  
    db_tmp = ""
    Call GetPrivateProfileString("Order", "USER", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.USER = Trim(frmInterface.txtTemp)
  
    '-- �˻���Ʈ
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "GumPart", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gGumPart = Trim(frmInterface.txtTemp)
    
    '== ������ DLL ���� ���� ����    =======================================================================
    db_tmp = ""
    Call GetPrivateProfileString("SERVICE", "URL", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gGINUS_Parm.URL = Trim(frmInterface.txtTemp)
  
    db_tmp = ""
    Call GetPrivateProfileString("SERVICE", "SVC", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gGINUS_Parm.SVC = Trim(frmInterface.txtTemp)
  
    db_tmp = ""
    Call GetPrivateProfileString("SERVICE", "HCD", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gGINUS_Parm.HCD = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("SERVICE", "MCD", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gGINUS_Parm.MCD = Trim(frmInterface.txtTemp)
    
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
'argSQL�� ������ ���Ϸ� ����
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    
    If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.Path & "\Log")
    End If
    
    sFileName = Format(CDate(frmInterface.dtpToday), "yyyy-mm-dd")
    
    Open App.Path & "\Log\" & sFileName & ".txt" For Append As FilNum
'    Print #FilNum, Format(Time, "hh:nn:ss") & " " & argSQL
    Print #FilNum, argSQL
    Close FilNum
    
End Sub

Public Sub SetSQLData(ByVal strName As String, ByVal argSQL As String)
'argSQL�� ������ ���Ϸ� ����
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
        
    If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.Path & "\Log")
    End If
    
    sFileName = strName
    
    Open App.Path & "\Log\" & sFileName & ".txt" For Output As FilNum
    Print #FilNum, argSQL
    Close FilNum
    
End Sub



'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڿ��� �����ڸ� �̿��� ������ ������ ��ġ�� ���ڿ��� ����
'   �μ� :
'       1.pText      : �����ڷ� ������ ���ڿ�
'       2.pPosiion   : ��ġ
'       3.pDelimiter : ������
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition �μ��� 1�� ��� For�� Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '�ش� �÷�
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
'   ��� : ������ Result Flags�� ���� �󼼼��� ��ȸ
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
'   ��� : ������ Abnormal Flag�� ���� ������ȸ
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
        .SortKey(1) = Col       '����Ű ����ȣ

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



Public Function tmaxexit() As Boolean
Dim ErrMsg As String

    If tpend() = -1 Then
        ErrMsg = tmaxerrdesc(gettperrno())
'        MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
'        MsgBox ("tp end error") & "," & gettperrno()
        tmaxexit = False
        Exit Function
    End If

    tmaxexit = True
    
End Function

Public Function tmaxerrdesc(ByVal intErrnum As Integer) As String
    
    tmaxerrdesc = ""
    tmaxerrdesc = vbNewLine
    
    Select Case intErrnum
        Case 1:  tmaxerrdesc = ""
        Case 2:  tmaxerrdesc = "[TPEBADDESC] cd�� ��ȿ���� ���� �������Դϴ�."
        Case 3:  tmaxerrdesc = "[TPEBLOCK] �ش� ���񽺰� �ٸ� ���ο� ���� ���ŷ�Ǿ� �ֽ��ϴ�."
        Case 4:  tmaxerrdesc = "[TPEINVAL] �Լ��� �μ��� ��ȿ���� �ʽ��ϴ�"
        Case 5:  tmaxerrdesc = "[TPELIMIT] �ý��� �ڿ� �Ǵ� Tmax���� �����ϴ� �ڿ��� �����մϴ�."
        Case 6:  tmaxerrdesc = "[TPENOENT] ȯ�������� ���񽺸��� Ȯ���ϼ���"
        Case 7:  tmaxerrdesc = "[TPEOS] ���ü�� �����Դϴ�. Tmax �ý����� �⵿�Ǿ����� �ʾҽ��ϴ�." & vbNewLine & _
                               "Ŭ���̾�Ʈ������ ���� ���� ip�� ��Ʈ��ȣ�� Ȯ���ϼ���"
        Case 8:  tmaxerrdesc = ""
        Case 9:  tmaxerrdesc = "[TPEPROTO] �������� ��Ȳ���� Tmax API�� ȣ��Ǿ����ϴ�."
        Case 10: tmaxerrdesc = "[TPESVCERR] ���� ���μ������� �ڿ��� �����ϰų�," & vbNewLine & _
                               "TPELIMIT ���� Ȥ�� ����Ÿ�Ӿƿ��� �ɷ� �ֽ��ϴ�"
        Case 11: tmaxerrdesc = "[TPESVCFAIL] ���� ���� �� ���� ���α׷� �������� ������ �߻��Ͽ����ϴ�."
        Case 12: tmaxerrdesc = ""
        Case 13: tmaxerrdesc = "[TPETIME] BLOCKTIME�� �ʰ� �Ͽ����ϴ�."
        Case 14: tmaxerrdesc = ""
        Case 15: tmaxerrdesc = ""
        Case 16: tmaxerrdesc = ""
        Case 17: tmaxerrdesc = "[TPEITYPE] �Էµ� ������ ������ Ȯ���Ͻʽÿ�"
        Case 18: tmaxerrdesc = "[TPEOTYPE] �۽��ڿ� ������ ���� ���� �ٸ� ���� ������ ����߽��ϴ�."
        Case 22: tmaxerrdesc = "[TPEEVENT] ��ȭ������ �̺�Ʈ�� �߻��Ͽ����ϴ�"
        Case 23: tmaxerrdesc = "[TPEMATCH] RQ ����, tpdeq �ϴ� ���� �̸����� RQ�� ����� ���񽺰� �����ϴ�."
        Case 24: tmaxerrdesc = "[TPENOREADY] Tmax ���񽺰� �غ���� �ʾҽ��ϴ�."
        Case 25: tmaxerrdesc = "[TPESECURITY] Tmax ���� ������ ���� ���Ͽ����ϴ�."
        Case 26: tmaxerrdesc = "[TPEQFULL] ��û�� ���񽺰� ������ Max Queue�� �����߽��ϴ�"
        Case 27: tmaxerrdesc = "[TPEQPURGE] ť�� Purge �Ǿ����ϴ�."
        Case 28: tmaxerrdesc = "[TPESVRDOWN] ���� ������ �ٿ�Ǿ����ϴ�."
        Case Else: tmaxerrdesc = "Tmax ����"
    End Select
    
    

End Function


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


' *************************************************************************************
'   ���� Error Process ���
'
'   Fmlptr  : FML buffer pointer    (long type)
'   service : call �Լ���           (string type)
'   msg_type: Error Massage ����    (integer type)
'             0    -> ATMI �޽��� ���
'             1    -> ATMI �޽��� + STATLIN �޽��� ���
'             2    -> ATMI �޽��� + STATLIN �޽��� + SQL code ���
'             else -> call �Լ��� ���
' *************************************************************************************
Function ErrorMsg(ByVal Fmlptr&, service As String, msg_type As Integer) As Integer

    Dim lret As Long
    Dim ret As Long
    Dim ErrMsgCl As String
    Dim ErrMsgSv As String
    Dim tp_err_no As Long
    Dim errptr As Long
    Dim Tpurcode As Long
    
    ErrMsgCl = String$(100, Chr$(32))
    ErrMsgSv = String$(100, Chr$(32))
    
    Select Case msg_type
        Case 0              '< TPINIT, ,TPBEGIN, TPALLOC ...>
            ' ATMI error �޼���
            tp_err_no = gettperrno()
            errptr = tpstrerror(tp_err_no)
            ret = lstrcpy(ByVal ErrMsgCl$, ByVal errptr&)
            Screen.MousePointer = 1         '����ȭ
            MsgBox service & " Failed." + Chr$(13) + "ATMI Msg : " + Left(ErrMsgCl, Len(Trim(ErrMsgCl)) - 1)
        Case 1              '< TPCALL fail ...>
            ' ATMI error �޼���
            tp_err_no = gettperrno()
            errptr = tpstrerror(tp_err_no)
            ret = lstrcpy(ByVal ErrMsgCl$, ByVal errptr&)
            ' ���� �޼��� (STATLIN)
            lret = Fvals32(ByVal Fmlptr&, ByVal STATLIN, 0)
            ret = lstrcpy(ByVal ErrMsgSv$, ByVal lret&)
            Screen.MousePointer = 1         '����ȭ
            MsgBox service & " Failed." + Chr$(13) + "ATMI Msg : " + Left(ErrMsgCl, Len(Trim(ErrMsgCl)) - 1) + Chr$(13) + "Server Msg : " + Trim(Left(Trim(ErrMsgSv), Len(Trim(ErrMsgSv)) - 1))
        Case 2              '< TPCALL fail: SQL error��...>
            ' ATMI error �޼���
            tp_err_no = gettperrno()
            errptr = tpstrerror(tp_err_no)
            ret = lstrcpy(ByVal ErrMsgCl$, ByVal errptr&)
            ' ���� �޼��� (STATLIN)
            lret = Fvals32(ByVal Fmlptr&, ByVal STATLIN, 0)
            ret = lstrcpy(ByVal ErrMsgSv$, ByVal lret&)
            ' sqlcode���� �޴´�.
            Tpurcode = gettpurcode()
            Screen.MousePointer = 1         '����ȭ
            MsgBox service & " Failed." + Chr$(13) + "ATMI Msg : " + Left(ErrMsgCl, Len(Trim(ErrMsgCl)) - 1) + Chr$(13) + "Server Msg : " + Trim(Left(Trim(ErrMsgSv), Len(Trim(ErrMsgSv)) - 1)) + Chr$(13) + "Sqlcode: " + Str(Tpurcode)
        Case 3              '< TPCALL fail: SQL error��...>
            ' ATMI error �޼���
            tp_err_no = getFerror32()
            errptr = Fstrerror32(tp_err_no)
            ret = lstrcpy(ByVal ErrMsgCl$, ByVal errptr&)
            ' ���� �޼��� (STATLIN)
            'lret = FVALS32(ByVal Fmlptr&, ByVal STATLIN, 0)
            'ret = lstrcpy(ByVal ErrMsgSv$, ByVal lret&)
            ' sqlcode���� �޴´�.
            'Tpurcode = gettpurcode()
            'Screen.MousePointer = 1         '����ȭ
           ' MsgBox service & " Failed." + Chr$(13) + "ATMI Msg : " + Left(ErrMsgCl, Len(Trim(ErrMsgCl)) - 1) + Chr$(13) + "Server Msg : " + Trim(Left(Trim(ErrMsgSv), Len(Trim(ErrMsgSv)) - 1)) + Chr$(13) + "Sqlcode: " + Str(Tpurcode)
             MsgBox service & " Failed." + ErrMsgCl
        Case 4              '< TPCALL fail: SQL error��...>
            ' ATMI error �޼���
            tp_err_no = getFerror32()
            errptr = Fstrerror32(tp_err_no)
            ret = lstrcpy(ByVal ErrMsgCl$, ByVal errptr&)
            ' ���� �޼��� (STATLIN)
            'lret = FVALS32(ByVal Fmlptr&, ByVal STATLIN, 0)
            'ret = lstrcpy(ByVal ErrMsgSv$, ByVal lret&)
            ' sqlcode���� �޴´�.
            'Tpurcode = gettpurcode()
            'Screen.MousePointer = 1         '����ȭ
           ' MsgBox service & " Failed." + Chr$(13) + "ATMI Msg : " + Left(ErrMsgCl, Len(Trim(ErrMsgCl)) - 1) + Chr$(13) + "Server Msg : " + Trim(Left(Trim(ErrMsgSv), Len(Trim(ErrMsgSv)) - 1)) + Chr$(13) + "Sqlcode: " + Str(Tpurcode)
             MsgBox service & " Failed." + ErrMsgCl
        Case Else           '< FML Function fail�� ...>
            Screen.MousePointer = 1         '����ȭ
            MsgBox service & " Failed."
    End Select
    
End Function

