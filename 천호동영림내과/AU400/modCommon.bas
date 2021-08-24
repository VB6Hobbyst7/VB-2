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
Public gEquipCode As String

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
    RsltDate As String
    BARCODE As String
    PID As String
    PNAME As String
    PSEX As String
    PAGE As String
    TESTCD As String
    Result As String
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


'''Public Const colSpecNo = 0  '미사용
'''Public Const colCheckBox = 1
'''Public Const colBarcode = 2
'''Public Const colRack = 3
'''Public Const colPos = 4
'''Public Const colPID = 5
'''Public Const colPName = 6
'''Public Const colSex = 7
'''Public Const colAge = 8
'''Public Const colOCnt = 9
'''Public Const colRCnt = 10
'''Public Const colState = 11
''''Public Const colA1c = 13
''''Public Const colIFCC = 15
''''Public Const coleAg = 17

Public Const colSpecNo = 0 '미사용
Public Const colCheckBox = 1
Public Const colHospDate = 2
Public Const colGubun = 3
Public Const colBarcode = 4
Public Const colRack = 5
Public Const colDISK = 5
Public Const colPos = 6
Public Const colPID = 7
Public Const colPName = 8
Public Const colSex = 9
Public Const colAge = 10
Public Const colOCnt = 11
Public Const colRCnt = 12
Public Const colState = 13

Public Const colA1c = 14
Public Const colIFCC = 15
Public Const coleAg = 16

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
Public blnSameRecord As Boolean



Type typeXMLInData
    Company     As String
    HospCode    As String
    ChartNo     As String
    PatName     As String
    PatJumin    As String
    PatNo       As String
    CommDate    As String
    ExamNo      As String
    ExamID      As String
    ComExamID   As String
    Specimen    As String
    Result      As String
    Reference   As String
    Remark      As String
    RsltDate    As String
    IOFlag      As String
End Type

Public XMLInData As typeXMLInData

Type typeXMLOutData
    Company     As String
    HospCode    As String
    ChartNo     As String
    PatName     As String
    PatJumin    As String
    PatNo       As String
    CommDate    As String
    ExamNo      As String
    ExamID      As String
    ComExamID   As String
    Specimen    As String
    Result      As String
    Reference   As String
    Remark      As String
    RsltDate    As String
    IOFlag      As String
End Type

Public XMLOutData As typeXMLOutData


Sub Main()
    'If App.PrevInstance Then
        frmInterface.Show
    'Else
        
    'End If
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

    '== 장비 관련 설정  ==============================================================================
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
    
    '== 통신 관련 설정  ==============================================================================
    db_tmp = ""
    Call GetPrivateProfileString("config", "gPort", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gPort = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gSpeed", "", db_tmp, 100, App.Path & "\Interf ace.ini")
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

    '-- osw 추가
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
    
    '==  Winsock 관련 설정    ==============================================================================
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "ServerIP", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDRDB_Parm.ServerIP = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "ServerPort", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDRDB_Parm.ServerPort = Trim(frmInterface.txtTemp)
        
    '== DB Table 관련 설정    ==============================================================================
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
        
    '== DB Table Column 관련 설정    =======================================================================
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
    gDBCOLUMN_Parm.Result = Trim(frmInterface.txtTemp)

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
    Call GetPrivateProfileString("Order", "MACHCD", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.MACHCD = Trim(frmInterface.txtTemp)
  
    db_tmp = ""
    Call GetPrivateProfileString("Order", "USER", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDBCOLUMN_Parm.USER = Trim(frmInterface.txtTemp)
  
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "PART", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gPart = Trim(frmInterface.txtTemp)
      
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


Public Function f_subSet_XMLWorkList(ByVal strDate As String, ByVal strDate1 As String, Optional ByVal strTime As String) As Variant
    Dim strPath   As String
    Dim strBuffer As String
    Dim i         As Long
    Dim lngBufLen As Long
    Dim BufChar   As String
    Dim strTmp As String
    Dim intIDX As Integer
    
    
On Error GoTo ErrorTrap
    
    Screen.MousePointer = 11
    
    '-- 오더파일명과 경로를 지정한다.
    strPath = "C:\UBCare\SINAI\IF\ExamIF_In.xml"

    
    '1라인씩 가져오기 MSDN내용
    Dim TextLine
    Open strPath For Input As #1 ' 파일을 엽니다.
    
    Do While Not EOF(1) ' 파일의 끝을 만날 때까지 반복합니다.
        Line Input #1, TextLine ' 변수로 데이터 행을 읽어들입니다.
        strBuffer = strBuffer & TextLine
    Loop
    
    Close #1 ' 파일을 닫습니다
 
    intIDX = 0
    lngBufLen = Len(strBuffer)
        
    For i = 1 To lngBufLen
        If intIDX = 0 Then
            BufChar = Mid$(strBuffer, i, 4)
        Else
            BufChar = Mid$(strBuffer, i + 3)
        End If
        
        If BufChar = "<검사>" Then
            intIDX = 1
            strTmp = BufChar
        Else
            strTmp = strTmp & BufChar
            If intIDX = 1 Then Exit For
        End If
    
    Next
    
'    f_subSet_XMLWorkList = Split(strTmp, "</검사>")
    strTmp = Replace(strTmp, "<검사>", ""): strTmp = Replace(strTmp, "</검사>", "|")
    strTmp = Replace(strTmp, "<업체>", ""): strTmp = Replace(strTmp, "</업체>", ",")
    strTmp = Replace(strTmp, "<요양기관번호>", ""): strTmp = Replace(strTmp, "</요양기관번호>", ",")
    strTmp = Replace(strTmp, "<차트번호>", ""): strTmp = Replace(strTmp, "</차트번호>", ",")
    strTmp = Replace(strTmp, "<수진자명>", ""): strTmp = Replace(strTmp, "</수진자명>", ",")
    strTmp = Replace(strTmp, "<주민등록번호>", ""): strTmp = Replace(strTmp, "</주민등록번호>", ",")
    strTmp = Replace(strTmp, "<내원번호>", ""): strTmp = Replace(strTmp, "</내원번호>", ",")
    strTmp = Replace(strTmp, "<의뢰일>", ""): strTmp = Replace(strTmp, "</의뢰일>", ",")
    strTmp = Replace(strTmp, "<검사번호>", ""): strTmp = Replace(strTmp, "</검사번호>", ",")
    strTmp = Replace(strTmp, "<검사ID>", ""): strTmp = Replace(strTmp, "</검사ID>", ",")
    strTmp = Replace(strTmp, "<업체검사ID>", ""): strTmp = Replace(strTmp, "</업체검사ID>", ",")
    strTmp = Replace(strTmp, "<검체>", ""): strTmp = Replace(strTmp, "</검체>", ",")
    strTmp = Replace(strTmp, "<결과치>", ""): strTmp = Replace(strTmp, "</결과치>", ",")
    strTmp = Replace(strTmp, "<참조치>", ""): strTmp = Replace(strTmp, "</참조치>", ",")
    strTmp = Replace(strTmp, "<소견>", ""): strTmp = Replace(strTmp, "</소견>", ",")
    strTmp = Replace(strTmp, "<결과일>", ""): strTmp = Replace(strTmp, "</결과일>", ",")
    strTmp = Replace(strTmp, "<업체>", ""): strTmp = Replace(strTmp, "</업체>", ",")
    strTmp = Replace(strTmp, "<입원외래구분>", ""): strTmp = Replace(strTmp, "</입원외래구분>", ",")
    
    f_subSet_XMLWorkList = Split(strTmp, "|")
    blnSameRecord = True
    
    'Kill strPath
    
    Screen.MousePointer = 0

    
    Exit Function
        
ErrorTrap:
    
    blnSameRecord = False
    Screen.MousePointer = 0
    
    
End Function


Public Function SeqSearch_New(ByVal brspread As Object, ByVal brSeq As String, ByVal brSeq2 As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqSearch_New = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Trim(.Text) = brSeq Then
                .Row = sCnt
                .Col = 3
                If Trim(.Text) = brSeq2 Then
                    SeqSearch_New = sCnt
                    .Action = ActionActiveCell
                    .Refresh
                    Exit For
                End If
            End If
        Next sCnt
    End With

End Function


Public Function SeqSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqSearch = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Trim(.Text) = brSeq Then
                .Row = sCnt
                .Col = 5
                SeqSearch = .Row
                Exit For
            End If
        Next sCnt
    End With

End Function

Public Function SeqNullSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqNullSearch = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Trim(.Text) = "" Then
                SeqNullSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function
