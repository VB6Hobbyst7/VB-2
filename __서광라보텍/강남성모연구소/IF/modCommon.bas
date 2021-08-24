Attribute VB_Name = "modCommon"
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                                ByVal lpKeyName As Any, ByVal lpString As Any, _
                                                                                                ByVal lplFileName As String) As Long

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
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

'-- OCS 업체
Public gOCS         As String
Public gSetup       As config
Public gPart        As String
Public gGubun       As Integer
Public gEquip       As String
Public gEquipCode   As String
Public gMach        As String

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
Public gIFUser As String

Public gArrEquip() As String

'-- Result S[read Column Seq Num
Public Const colSpecNo = 0     '미사용
Public Const colCheckBox = 1
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
    TestTime As String
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

Public gIFMode  As String
Public gScreen  As String
Public gUnit    As Boolean
Public gPOri    As String

Sub Main()
    
On Error GoTo onError
    
    '-- 두번 실행 하지 않음
    If App.PrevInstance Then
       MsgBox "프로그램이 이미 실행중입니다.!", vbExclamation
       End
    End If
    
    frmSplash.Show
    
    frmSplash.labMsg.Caption = "인터페이스 프로그램 로딩중입니다."
    DoEvents
    
    frmSplash.Timer1.Interval = 500
    frmSplash.Timer1.Enabled = True
        
Exit Sub

onError:

End Sub

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
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "OCS", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gOCS = Trim(frmInterface.txtTemp)
    
    '== 장비 관련 설정  ==============================================================================
    db_tmp = ""
    Call GetPrivateProfileString("config", "Equip", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquip = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "EquipCode", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquipCode = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "MACH", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gMach = Trim(frmInterface.txtTemp)
    
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
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "IFMode", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gIFMode = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "IFScreen", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gScreen = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "IFUser", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gUserID = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "IncludUnit", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    If Trim(frmInterface.txtTemp) = "Y" Then
        gUnit = True
    Else
        gUnit = False
    End If
    
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "IncludUnit", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    If Trim(frmInterface.txtTemp) = "l" Then
        gPOri = "L" '가로출력
    Else
        gPOri = "P" '세로출력
    End If
    
    GetSetup = True

End Function

Public Sub SetRawData(argSQL As String)
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.Path & "\Log")
    End If
    
    sFileName = Format(Now, "yyyy-mm-dd")
    
    Open App.Path & "\Log\" & sFileName & ".txt" For Append As FilNum
    Print #FilNum, argSQL;
    Close FilNum
    
End Sub

Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, ByVal pDelimiter As String) As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
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
