Attribute VB_Name = "modCommon"
Option Explicit

Public gERP         As String
Public gMACH        As String

Public gMACHCOUNT   As Integer
Public gMACHS()     As String
                    
Public gCOLWIDTH    As String
Public gCOLHEADER   As String
Public gCOLVIEW     As String
Public gCOLSIZE     As String

Public gWORKPOS     As String
Public gWORKTEST    As String

Public strSetup     As String * 100
Public strSetUp1    As String

Public gSORT        As Integer


'==== 인쇄관련 상수
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_PAINT = &HF
Public Const WM_PRINT = &H317


'
'Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
'Public Declare Function GetDefaultCommConfig Lib "kernel32" Alias "GetDefaultCommConfigA" (ByVal lpszName As String, lpCC As COMMCONFIG, lpdwSize As Long) As Long
'Private Const CHUNK_SIZE& = 4096&
'Private Const CP_UTF8 As Long = 65001
'Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
'Private Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cbMultiByte As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
'Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
'                    (ByVal lpApplicationName As String _
'                   , ByVal lpKeyName As Any _
'                   , ByVal lpString As Any _
'                   , ByVal lplFileName As String) As Long
'
'Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
'                    (ByVal lpApplicationName As String _
'                   , ByVal lpKeyName As Any _
'                   , ByVal lpDefault As String _
'                   , ByVal lpReturnedString As String _
'                   , ByVal nSize As Long _
'                   , ByVal lpFileName As String) As Long

                    
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
    USERGRD     As String
    
    LOGINYN     As String
    SAVEPW      As String
    BARUSE      As String
    SAVEAUTO    As String
    SAVELIS     As String
    RSTTYPE     As String
    QCPATH      As String
    LOGWRITE    As String
    SAVEDAY     As String
    BARLEN      As Integer
    DBCONCHK    As String
    APIURL      As String
    USEURL      As String
    STDURL      As String
    DEVURL      As String
    ORDCODE     As String
    EMR         As String
    COMPNM      As String
    TITLE       As String
End Type

Public gKUKDO        As HospParameter
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
    BARPOS      As String
End Type

Public gComm        As ComParameter
'########## 통신정보 관련 ############################

'########## 폼 관련 ############################
Type FormParameter
    MAXYN       As String
    TOP         As String
    LEFT        As String
    WIDTH       As String
    HEIGHT      As String
End Type

Public gForm        As FormParameter
'########## 폼 관련 ############################

'########## 시리얼포트 찾기 ############################
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
'########## 시리얼포트 찾기 ############################

'##########  사용자 정보 (DB 전달용)
Type UserParam
    ID      As String
    NAME    As String
    PW      As String
    DEPT    As String
    COMP    As String
    YN      As String
    REGID   As String
    REGDT   As String
    MDYID   As String
    MDYDT   As String
End Type

Public gUSER    As UserParam
'##########  사용자 정보 (DB 전달용)

'##########  고객사 정보 (DB 전달용)
Type CompParam
    CD      As String
    NAME    As String
    LINE    As String
    VIEW    As String
    DISNO   As String
    YN      As String
    REGID   As String
    REGDT   As String
    MDYID   As String
    MDYDT   As String
End Type

Public gComp    As CompParam
'##########  고객사 정보 (DB 전달용)

'##########  포장코드 정보 (DB 전달용)
Type PackParam
    CD      As String
    NAME    As String
    CORE    As String
    DIA     As String
    CWID    As String
    PWID    As String
    pLen    As String
    CGBN    As String
    DISNO   As String
    YN      As String
    REGID   As String
    REGDT   As String
    MDYID   As String
    MDYDT   As String
End Type

Public gPack    As PackParam
'##########  포장코드 정보 (DB 전달용)

'##########  포장코드 정보 (DB 전달용)
Type ProdParam
    CD      As String
    NAME    As String
    PRTNAME As String
    COMPCD  As String
    LEN     As String
    METCD   As String
    MONTH   As String
    TEMP    As String
    SIZE    As String
    CHPN    As String
    VDCD    As String
    LINEFA  As String
    SLITFA  As String
    CTYN    As String
    PCNNO   As String
    BAR     As String
    YN      As String
    REGID   As String
    REGDT   As String
    MDYID   As String
    MDYDT   As String
End Type

Public gProd    As ProdParam
'##########  포장코드 정보 (DB 전달용)

'##########  자재코드 정보 (DB 전달용)
Type MatParam
    CD      As String
    NAME    As String
    DISNO   As String
    YN      As String
    REGID   As String
    REGDT   As String
    MDYID   As String
    MDYDT   As String
End Type

Public gMAT    As MatParam
'##########  자재코드 정보 (DB 전달용)


'##########  제품라벨 정보(Header) (DB 전달용)
Type LabelMasterParam
    LABELCD         As String   'key
    PRODCD          As String
    COMPCD          As String
    LBLTYPE         As String
    LBLPRTNO        As String   '박스당 릴 기본수량
    LBLPRTSIDE      As String
    LBLBARSIDE1     As String
    LBLBARSIDE2     As String
    PRODMAXTOT      As String
    BARMAXLEN       As String
    BARPNOPOS       As String
    BARPNOLEN       As String
    BARPADDYN       As String
    BARSNOPOS       As String
    BARSNOLEN       As String
    BARSADDYN       As String
    YN              As String
    REGID           As String
    REGDT           As String
    MDYID           As String
    MDYDT           As String
End Type

Public gLblMaster    As LabelMasterParam
'##########  제품라벨 정보(Header) (DB 전달용)

'##########  제품라벨 정보(Master) (DB 전달용)
Type LabelDetailParam
    LABELCD         As String   'key
    LBLITEM_NO()    As String   'key
    LBLITEM_SEQ()   As String
    LBLITEM_NAME()  As String
    LBLITEM_NMPRT() As String
    LBLITEM_BARGU() As String
    LBLITEM_BARCD() As String
    LBLITEM_X()     As String
    LBLITEM_Y()     As String
    LBLITEM_FONT()  As String
    LBLITEM_ROT()   As String
    YN()            As String
    REGID           As String
    REGDT           As String
    MDYID           As String
    MDYDT           As String
End Type

Public gLblDetail    As LabelDetailParam
'##########  제품라벨 정보(Master) (DB 전달용)

'##########  바코드마스터 정보(Header) (DB 전달용)
Type BarMasterParam
    BARCD           As String   'key
    PRODCD          As String
    COMPCD          As String   '바코드타입 1,2(QR)
    BARTYPE         As String
    BARGU           As String   '라벨구분 R,P,I
    TEMPGU          As String   'TEMP MASTER 구분코드
    YN              As String
    REGID           As String
    REGDT           As String
    MDYID           As String
    MDYDT           As String
End Type

Public gBarMaster    As BarMasterParam
'##########  바코드마스터 정보(Header) (DB 전달용)

'##########  바코드마스터 정보(Detail) (DB 전달용)
Type BarDetailParam
    BARCD           As String   'key
    BARITEM_NO()     As String
    BARITEM_SEQ()    As String
    BARITEM_NAME()   As String
    BARCHRNUM()     As String
    LBLITEMTYPE()   As String
    YN()            As String
    REGID           As String
    REGDT           As String
    MDYID           As String
    MDYDT           As String
End Type

Public gBarDetail    As BarDetailParam
'##########  바코드마스터 정보(Detail) (DB 전달용)

'##########  TEMP 마스터 정보 (DB 전달용)
Type TempParam
    GUBUN   As String
    Seq     As String
    CODE1   As String
    CODE2   As String
    CODE3   As String
    CDVAL1  As String
    CDVAL2  As String
    CDVAL3  As String
    DESC    As String
End Type

Public gTemp    As TempParam
'##########  TEMP 마스터 정보 (DB 전달용)

'##########  작업지시서 정보(Header) (DB 전달용)
Type OrderParam
    ORDDATE     As String   'key
    PRODPOSNO   As String   'key
    PRODCD      As String   'key
    SLITINGNO   As String   'key
    'NO          As String
    COMPCD      As String
    PRODNAME    As String
    PACKCD      As String
    REELQTY     As String
    'JOBINFO     As String
'    ROLLINFO    As String
    ORDERMEMO   As String
    LOTNO       As String
    CLOSEYN     As String
    YN          As String
    REGID       As String
    REGDT       As String
    MDYID       As String
    MDYDT       As String
End Type

Public gOrder    As OrderParam
'##########  작업지시서 정보(Header) (DB 전달용)

'##########  작업지시서 정보(Detail) (DB 전달용)
Type OrderParamDetail
    ORDDATE     As String   'key
    PRODPOSNO   As String   'key
    PRODCD      As String   'key
    SLITINGNO   As String   'key
    NO()        As String   'key
    SLTINFO()   As String
    PFROMNO()   As String
    PTONO()     As String
End Type

Public gOrderDetail    As OrderParamDetail
'##########  작업지시서 정보(Detail) (DB 전달용)


'##########  제품 TRACKING 정보() (DB 전달용)
Type PackTrackParam
    ORDERDT     As String   'key
    PRODCD      As String   'key
    REELBAR     As String   'key
    PPBAR       As String
    ICEBAR      As String
    PPBARIN     As String   '?
    ICEBARIN    As String   '?
    LOTNO       As String
    REELPRTID   As String
    REELPRTDT   As String
    PPPRTID     As String
    PPPRTDT     As String
    ICEPRTID    As String
    ICEPRTDT    As String
    REELVAL     As String
    PPVAL       As String
    ICEVAL      As String
End Type

Public gPackTrack    As PackTrackParam
'##########  제품 TRACKING 정보() (DB 전달용)


Sub Main()
        
    '두번 실행 하지 않음
    If App.PrevInstance Then
       MsgBox "프로그램이 이미 실행중입니다.!", vbExclamation
       End
    End If
    
    '-- INI 정보
    Call GetSetup
    
    If gDBCONN = "1" Then
        '-- 로컬 DB 접속
        If Not DbConnect_Local Then
            If vbYes = MsgBox("로컬 데이터베이스가 없습니다. 찾으시겠습니까? ", vbCritical + vbYesNo) Then
                frmDB_Local.Show vbModal
            Else
                End
            End If
        Else
            cn_Server_Flag = True
        End If
    Else
        '-- MSSQL DB 접속
        If Not DbConnect_SQL Then
            If vbYes = MsgBox("MS-SQL 연결정보가 없습니다. 등록하시겠습니까? ", vbCritical + vbYesNo) Then
                frmDB_MSSQL.Show vbModal
            Else
                End
            End If
        Else
            cn_Server_Flag = True
        End If
    End If

    If cn_Server_Flag = True Then
        Call frmLogin.Show
    End If
    
End Sub


Public Sub CtlInitializing()
                 
    RcvBuffer = ""
    Erase strRecvData
    
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 1
    intFrameNo = 1

End Sub

Public Sub SetMenu()
    
'    With frmMain
'        '-- 바코드사용
'        If gKUKDO.BARUSE = "Y" Then
'            .mnuBarcode.Checked = True
'            .mnuSeqno.Checked = False
''            .optBarSeq(0).Value = True
'        Else
'            .mnuBarcode.Checked = False
'            If gKUKDO.RSTTYPE = "1" Then
'                .mnuSeqno.Checked = True
''                .optBarSeq(1).Value = True
'            ElseIf gKUKDO.RSTTYPE = "2" Then
'                .mnuRackPos.Checked = True
''                .optBarSeq(2).Value = True
'            ElseIf gKUKDO.RSTTYPE = "3" Then
'                .mnuCheckBox.Checked = True
''                .optBarSeq(3).Value = True
'            End If
'        End If
'
'        '-- 결과전송
'        If gKUKDO.SAVEAUTO = "Y" Then
'            .mnuSaveAuto.Checked = True
'            .mnuSaveManual.Checked = False
''            .optTrans(0).Value = True
'        Else
'            .mnuSaveAuto.Checked = False
'            .mnuSaveManual.Checked = True
''            .optTrans(1).Value = True
'        End If
'
'        '-- 적용결과
'        If gKUKDO.SAVELIS = "Y" Then
'            .mnuLisResult.Checked = True
'            .mnuEqpResult.Checked = False
''            .optSaveResult(1).Value = True
'        Else
'            .mnuLisResult.Checked = False
'            .mnuEqpResult.Checked = True
''            .optSaveResult(0).Value = True
'        End If
'
'
'    End With
    
    
End Sub

Public Sub SetCommStatus(ByVal pSRflag As String, ByVal pBarNo As String, ByVal SPD As Object)
    
'    With SPD
'        .MaxRows = .MaxRows + 1
'        If pSRflag = "S" Then
'            Call SetText(SPD, "Send", .MaxRows, 1)
'            Call SetText(SPD, pBarNo, .MaxRows, 2)
'            Call SetText(SPD, "오더전송", .MaxRows, 3)
'
'        ElseIf pSRflag = "Q" Then
'            Call SetText(SPD, "Recv", .MaxRows, 1)
'            Call SetText(SPD, pBarNo, .MaxRows, 2)
'            Call SetText(SPD, "오더요청", .MaxRows, 3)
'
'        ElseIf pSRflag = "R" Then
'            Call SetText(SPD, "Recv", .MaxRows, 1)
'            Call SetText(SPD, pBarNo, .MaxRows, 2)
'            Call SetText(SPD, "결과수신", .MaxRows, 3)
'        End If
'
'        .Row = .MaxRows
'        .Col = 1
'        .Action = ActionActiveCell
'
'        If .MaxRows > 100 Then
'            Call DeleteRow(SPD, 1, 1)
'            .MaxRows = .MaxRows - 1
'        End If
'
'    End With
    
    With SPD
        '.MaxRows = .MaxRows + 1
        If pSRflag = "S" Then
            .AddItem "Send" & vbTab & pBarNo & vbTab & "장비전송"
            'Call SetText(SPD, "Send", .MaxRows, 1)
            'Call SetText(SPD, pBarNo, .MaxRows, 2)
            'Call SetText(SPD, "오더전송", .MaxRows, 3)
            
        ElseIf pSRflag = "Q" Then
            'Call SetText(SPD, "Recv", .MaxRows, 1)
            'Call SetText(SPD, pBarNo, .MaxRows, 2)
            'Call SetText(SPD, "오더요청", .MaxRows, 3)
            
            .AddItem "Recv" & vbTab & pBarNo & vbTab & "오더요청"
        
        ElseIf pSRflag = "R" Then
            'Call SetText(SPD, "Recv", .MaxRows, 1)
            'Call SetText(SPD, pBarNo, .MaxRows, 2)
            'Call SetText(SPD, "결과수신", .MaxRows, 3)
            .AddItem "Recv" & vbTab & pBarNo & vbTab & "결과수신"
        End If
            
        '.Row = .MaxRows
        '.Col = 1
        '.Action = ActionActiveCell
        
        'If .MaxRows > 100 Then
        '    Call DeleteRow(SPD, 1, 1)
        '    .MaxRows = .MaxRows - 1
        'End If
        
    End With
    
    
End Sub

Public Sub SetColumnView(ByVal SPD As Object)
    Dim i       As Integer
    Dim varSize As Variant

    varSize = Split(gCOLSIZE, "|")

    For i = 0 To UBound(varSize) - 1
        SPD.Col = i + 1
        If Mid(gCOLVIEW, i + 1, 1) = 1 Then
            SPD.ColHidden = False
            
            If varSize(i) <> "" Then
                SPD.ColWidth(i + 1) = varSize(i)
            End If
        
        Else
            SPD.ColHidden = True
        End If
    Next

End Sub

Public Sub SetColumnHeader(ByVal SPD As Object)
    Dim i       As Integer
    Dim varHeader As Variant

    varHeader = Split(gCOLHEADER, "|")

    For i = 0 To UBound(varHeader) - 1
        Call SetText(SPD, varHeader(i), 0, i + 1)
        'SPD.Alignment = 2
        SPD.Font = "맑은 고딕"
        SPD.FontSize = 10
    Next
        
    SPD.RowHeight(-1) = 15

    
    
End Sub

'--  INFO GET
Public Sub GetExeVersion()
    Dim i As Integer
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DB", "DBCONN", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gDBCONN = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DB", "DBTYPE", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gDBTYPE = Trim(strSetUp1)
    
    
'    strSetup = "":    strSetUp1 = ""
'    Call GetPrivateProfileString("EXE", "MACH", "", strSetup, 100, App.PATH & "\KDBAR.ini")
'    strSetUp1 = Trim(strSetup)
'    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'    gMACH = Trim(strSetUp1)
'
'    strSetup = "":    strSetUp1 = ""
'    Call GetPrivateProfileString("EXE", "DBTYPE", "", strSetup, 100, App.PATH & "\KDBAR.ini")
'    strSetUp1 = Trim(strSetup)
'    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'    gDBTYPE = Trim(strSetUp1)
'
'    strSetup = "":    strSetUp1 = ""
'    Call GetPrivateProfileString("EXE", "MACHCOUNT", "", strSetup, 100, App.PATH & "\KDBAR.ini")
'    strSetUp1 = Trim(strSetup)
'    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'    gMACHCOUNT = Trim(strSetUp1)
'
'    If IsNumeric(gMACHCOUNT) Then
'        ReDim Preserve gMACHS(gMACHCOUNT) As String
'        For i = 1 To gMACHCOUNT
'            strSetup = "":    strSetUp1 = ""
'            Call GetPrivateProfileString("EXE", "MACH" & CStr(i), "", strSetup, 100, App.PATH & "\KDBAR.ini")
'            strSetUp1 = Trim(strSetup)
'            strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'            gMACHS(i) = Trim(strSetUp1)
'        Next
'    End If
    
End Sub

Public Sub GetSetup()
    
    '-- DB 연결 관련
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DB", "DBCONN", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gDBCONN = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DB", "DBTYPE", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gDBTYPE = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DB", "ERP", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gERP = Trim(strSetUp1)
    
    '-- 사용자 정보
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("USER", "USERID", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gKUKDO.USERID = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("USER", "USERNM", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gKUKDO.USERNM = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("USER", "SAVEPW", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gKUKDO.SAVEPW = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("USER", "COMPNM", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gKUKDO.COMPNM = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("USER", "TITLE", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gKUKDO.TITLE = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("USER", "LOGWRITE", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gKUKDO.LOGWRITE = Trim(strSetUp1)
    
    
    '-- FORM INFO GET (폼사이즈 기억)
    Call GetPrivateProfileString("FORM", "MAXYN", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gForm.MAXYN = Trim(strSetUp1)
    
    Call GetPrivateProfileString("FORM", "TOP", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gForm.TOP = Trim(strSetUp1)
    
    Call GetPrivateProfileString("FORM", "LEFT", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gForm.LEFT = Trim(strSetUp1)
    
    Call GetPrivateProfileString("FORM", "WIDTH", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gForm.WIDTH = Trim(strSetUp1)
    
    Call GetPrivateProfileString("FORM", "HEIGHT", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gForm.HEIGHT = Trim(strSetUp1)
    

    
    '-- LOCAL DB GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "LOCALPATH", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gLocalDB.PATH = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "LOCALUID", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gLocalDB.UID = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "LOCALPWD", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gLocalDB.PWD = Trim(strSetUp1)

    
    '-- MSSQL DB GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DB", "MSSQLIP", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB.IP = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DB", "MSSQLDB", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB.DB = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DB", "MSSQLUID", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB.UID = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DB", "MSSQLPWD", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB.PWD = Trim(strSetUp1)


    '-- 시리얼 통신
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "COMPORT", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.COMPORT = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "SPEED", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.SPEED = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "PARITY", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.Parity = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "DATABIT", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.DATABIT = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "STARTBIT", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.STARTBIT = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "STOPBIT", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.STOPBIT = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "RTSEnable", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.RTSEnable = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "DTREnable", "", strSetup, 100, App.PATH & "\KDBAR.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.DTREnable = Trim(strSetUp1)
    
End Sub


Public Sub SaveExcel(Filename As String, argSpread As Object)

On Error Resume Next

' Excel Object Library 와 연결합니다.
Dim xlapp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim iRow    As Integer
Dim iCol    As Integer
Dim i       As Integer

    Set xlapp = CreateObject("Excel.Application")
    
    xlapp.DisplayAlerts = False
    
    Set xlBook = xlapp.Workbooks.Add
    
    Set xlSheet = xlBook.Worksheets(1)
        
    i = 0
    For iRow = 0 To argSpread.DataRowCnt
        For iCol = 1 To argSpread.DataColCnt
            If iCol = colEXAMDATE Or iCol = colEXAMTIME Or iCol = colSAVESEQ Or iCol = colHOSPDATE Or iCol = colBARCODE Or iCol > colSTATE Then
                i = i + 1
                argSpread.Row = iRow
                argSpread.Col = iCol
                'xlSheet.Cells(iRow + 1, iCol) = argSpread.Text
                If iCol = colBARCODE Then
                    xlSheet.Cells(iRow + 1, i) = EB & argSpread.Text
                Else
                    xlSheet.Cells(iRow + 1, i) = argSpread.Text
                End If
            End If
        Next iCol
        i = 0
    Next iRow
    
    xlBook.SaveAs (Filename)
    xlapp.Quit


End Sub

'-- 스프레드 정렬
'Public Sub SetSpreadSort(SP As Object)
Public Sub SetSpreadSort(SP As Object, Optional ByVal iSortoption As Integer = 0)
'    Dim i As Integer
'
'    If gSORT = 0 Then
'        gSORT = 1
'    Else
'        gSORT = 0
'    End If
'
'    '## Setting Sort Indicate
'    For i = 1 To SP.MaxCols
'        SP.ColUserSortIndicator(i) = gSORT
'    Next
'
'    SP.UserColAction = gSORT
    
    Dim i As Integer
    
    '## Setting Sort Indicate
    For i = 1 To SP.MaxCols
        SP.ColUserSortIndicator(i) = iSortoption
    Next
    
    SP.UserColAction = iSortoption 'UserColActionSort
       
   
End Sub

Public Sub frmShow(frm As Form)
    
    frmMDI.lblFrmInfo.Caption = "현재화면 : " & frm.Caption & " : " & frm.Tag
    frm.Show
    frm.ZOrder 0
    
End Sub

