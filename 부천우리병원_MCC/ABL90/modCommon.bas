Attribute VB_Name = "modCommon"
Option Explicit

Public gEMR         As String
Public gMACH        As String
Public gMACHCOUNT   As Integer
Public gMACHS()     As String

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String _
                                                                                            , ByVal lpKeyName As Any _
                                                                                            , ByVal lpString As Any _
                                                                                            , ByVal lplFileName As String) As Long
    
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String _
                                                                                        , ByVal lpKeyName As Any _
                                                                                        , ByVal lpDefault As String _
                                                                                        , ByVal lpReturnedString As String _
                                                                                        , ByVal nSize As Long _
                                                                                        , ByVal lpFileName As String) As Long
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
    LOQWRITE    As String
    SAVEDAY     As String
    BARLEN      As Integer
    DBCONCHK    As String
    MENULOCK    As String
    NEG         As String
    POS         As String
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
    RSTPATH     As String
End Type

Public gComm        As ComParameter
'########## 통신정보 관련 ############################

'########## 의사랑[UBCARE] 처방 XML ##################
Type XMLInData
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

Public XMLInData As XMLInData
'########## 의사랑[UBCARE] 처방 XML ##################

'########## 의사랑[UBCARE] 결과 XML ##################
Type XMLOutData
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

Public XMLOutData As XMLOutData
'########## 의사랑[UBCARE] 결과 XML ##################

'########## 폼 관련 ############################
Type FormParameter
    MAXYN       As String
    TOP         As String
    LEFT        As String
    WIDTH       As String
    HEIGHT      As String
End Type

Public gForm        As FormParameter
'########## 병원정보 관련 ############################

'########## JSON 관련 ############################
Type JsonParameter
    LOGIN       As String
    WORKLIST    As String
    PATLIST     As String
    PATSAVE     As String
End Type

Public gURL        As JsonParameter
'########## JSON 관련 ############################

'########## 보건소 관련 ############################
Type HealthParameter
    INITURL     As String
    'WORKLIST    As String
    'PATLIST     As String
    'PATSAVE     As String
End Type

Public gHEALTH      As HealthParameter
'########## 보건소 관련 ############################

'########## FTP 관련 ############################
Type FTPParameter
    SERVER      As String
    port        As Long
    UID         As String
    PWD         As String
End Type

Public gFTP      As FTPParameter
'########## FTP 관련 ############################


Public strSetup     As String * 100
Public strSetUp1    As String

Public gArrEQP()    As String
Public gArrEQPNm()  As String   '-- 인터페이스에 등록된 전체검사명
Public gAllTestCd   As String   '-- 인터페이스에 등록된 전체검사코드
Public gAllTestCd_F As String   '-- 인터페이스에 등록된 전체검사코드
Public gAllOrdCd    As String   '-- 인터페이스에 등록된 전체오더코드
Public gPatOrdCd    As String   '-- 검체별 검사코드
Public gPatOrdNm    As String   '-- 검체별 검사명
Public gPatTest()   As String   '-- 환자처방된 전체검사코드
Public gRow         As Long     '-- 작업중 Row

Public gCENXPCD     As String
Public gADV18CD     As String

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
Public gCOLHEADER        As String
Public gCOLVIEW          As String
Public gCOLSIZE          As String
Public gDETAILVIEW       As String
Public gROWHEIGHT        As String
Public gCOLVIEW_R        As String
Public gCOLSIZE_R        As String
Public gWORKPOS          As String
Public gWORKTEST         As String

Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Dim mmftp           As New clsFTP     'FTP관련
Dim hOpen           As Long
Dim dwType          As Long
Dim hConnection     As Long

Const ASCII         As Long = FTP_TRANSFER_TYPE_ASCII
Const BINARY        As Long = FTP_TRANSFER_TYPE_BINARY

Public Const SPCYYLEN As Long = 2
Public Const SPCNOLEN As Long = 9

Sub Main()
    
On Error GoTo onError
    
    
    '-- 두번 실행 하지 않음
    If App.PrevInstance Then
       MsgBox "프로그램이 이미 실행중입니다.!", vbExclamation
       End
    End If
    
    frmSplash.Show
    
    
    '-- SANSOFT.INI 정보
    Call GetExeVersion
    
    '-- INI 정보
    'Call GetSetup
    
    frmSplash.labMsg.Caption = "프로그램 로딩중입니다."
    DoEvents
    
    '-- 레지스트리 주소 설정
    REG_MACH = REG_POSITION & "\" & gHOSP.MACHNM & "\" & gMACH
    
    frmSplash.labMsg.Caption = "인터페이스 설정 백업중입니다."
    DoEvents
    
    '-- 레지스트리 정보 읽어오기
    Call GetRegSetup
    
'''    '-- 레지스트리 백업
'''    Call Shell(App.PATH & "\RegBackup.bat", vbNormalFocus)
'''
'''    frmSplash.labMsg.Caption = "검사코드 백업중입니다."
'''    DoEvents
'''
'''    '############################################ FTP 백업 Begin
'''    '
'''    '폴더구조
'''    '/
'''    '   home
'''    '   photo
'''    '   san
'''    '       interface
'''    '           #병원이름
'''    '               #장비이름
'''
'''    '세션 연결처리
'''    'hOpen = 0
'''    'hConnection = 0
'''    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
'''    hConnection = InternetConnect(hOpen, gFTP.SERVER, gFTP.port, gFTP.UID, gFTP.PWD, INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
'''
'''    '-- FTP서버에 경로만들기
'''    Call SetBackup
'''
'''    '-- 레지스트리 백업
'''    If UpLoad("C:\SANSOFT.REG", "/san/interface/" & gHOSP.HOSPNM & "/" & gHOSP.MACHNM & "/" & "SANSOFT.REG", ASCII) Then
'''        'success
'''    Else
'''        frmSplash.labMsg.Caption = "인터페이스 설정 백업실패."
'''        DoEvents
'''    End If
'''
'''    '-- MDB 백업
'''    If UpLoad(App.PATH & "\Database\" & gHOSP.MACHNM & ".mdb", "/san/interface/" & gHOSP.HOSPNM & "/" & gHOSP.MACHNM & "/" & gHOSP.MACHNM & ".mdb", BINARY) Then
'''        'success
'''    Else
'''        frmSplash.labMsg.Caption = "검사코드 백업실패."
'''        DoEvents
'''    End If
'''
'''    '-- EXE 백업
'''    If UpLoad(App.PATH & "\IF_" & gHOSP.MACHNM & ".exe", "/san/interface/" & gHOSP.HOSPNM & "/" & gHOSP.MACHNM & "/IF_" & gHOSP.MACHNM & ".exe", BINARY) Then
'''        'success
'''    Else
'''        frmSplash.labMsg.Caption = "인터페이스 백업실패."
'''        DoEvents
'''    End If
    '############################################ FTP 백업 Finish
    
    frmSplash.labMsg.Caption = "인터페이스 프로그램을 로딩합니다."
    DoEvents
    
    'Unload frmSplash
    'frmSplash.Timer1.Interval = 3000
    'frmSplash.Timer1.Enabled = True
    
    '-- 사용장비
    If gMACH = "" Then
        frmEMRInfo.Show vbModal
    End If
    
    '-- EMR
    If gEMR = "" Then
        frmEMRInfo.Show vbModal
    End If
    
    '-- 병원코드
    If Len(gHOSP.HOSPCD) = 0 Then
        frmHospInfo.Show vbModal
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
            If vbYes = MsgBox("오라클 연결정보가 없습니다. 등록하시겠습니까? ", vbCritical + vbYesNo) Then
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
            If vbYes = MsgBox("MS-SQL 연결정보가 없습니다. 등록하시겠습니까? ", vbCritical + vbYesNo) Then
                frmDB_MSSQL.Show vbModal
            Else
                End
            End If
        Else
            cn_Server_Flag = True
        End If
    ElseIf gDBTYPE = "3" Then
        '-- PostGresSQL 연결
        If Not DbConnect_PostGres Then
            If vbYes = MsgBox("Postgres SQL 연결정보가 없습니다. 등록하시겠습니까? ", vbCritical + vbYesNo) Then
                frmDB_PGSQL.Show vbModal
            Else
                End
            End If
        Else
            cn_Server_Flag = True
        End If
    ElseIf gDBTYPE = "99" Then
        cn_Server_Flag = False 'True
    Else
        MsgBox "데이터베이스 연결설정을 확인하세요.", vbCritical, "데이터베이스 설정"
        End
    End If
    
    '-- 컨트롤초기화
    Call CtlInitializing
    
    '-- 로그인 사용자
    If gHOSP.LOGINYN = "Y" Then
        Call frmLogin.Show
    Else
        Call MDIIF.Show
    End If
    
    Unload frmSplash
Exit Sub

onError:
    
    frmSplash.labErrMsg = Err.Description
    frmSplash.cmdExit.Visible = True
    
End Sub


Function UpLoad(szFileLocal As String, szFileRemote As String, dwType As Long) As Boolean
    'Dim bRet As Boolean
    UpLoad = FtpPutFile(hConnection, szFileLocal, szFileRemote, dwType, 0)
End Function

Function DownLoad(szFileRemote As String, szFileLocal As String, dwType As Long) As Boolean
    'Dim bRet As Boolean
    DownLoad = FtpGetFile(hConnection, szFileRemote, szFileLocal, False, INTERNET_FLAG_RELOAD, dwType, 0)
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 폼을 띄운다.
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal frmThis As Form, ByVal strFrmNm As String)
    
    Screen.MousePointer = vbHourglass
    
    If frmThis.MDIChild = True Then
        MDIIF.lblMenuInfo.Caption = strFrmNm
        frmThis.Show
        frmThis.ZOrder 0
        
        MDIIF.cmdNode.Caption = "▶"
        MDIIF.TreeView1.Visible = False
        MDIIF.picNode.WIDTH = 400 '300
        MDIIF.cmdNode.LEFT = 0
        MDIIF.cmdNode.HEIGHT = MDIIF.ScaleHeight - MDIIF.picHeader.HEIGHT
        
        'Call FrmMove
        
    Else
        'frmThis.Show , frmThis
        frmThis.Show ', frmThis
        frmThis.ZOrder 0
        DoEvents
    End If
    
    Screen.MousePointer = vbDefault

End Sub

Public Sub frmShow(frm As Form)
    
    'MDIIF.lblFrmInfo.Caption = "현재화면 : " & frm.Caption & " : " & frm.Tag
    
'    frm.WindowState = 2
    
    frm.Show
    
    frm.ZOrder 0
    
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
    
    With MDIIF
        '-- 바코드사용
        If gHOSP.BARUSE = "Y" Then
            .mnuBarcode.Checked = True
            .mnuSeqno.Checked = False
'            .optBarSeq(0).Value = True
        Else
            .mnuBarcode.Checked = False
            If gHOSP.RSTTYPE = "1" Then
                .mnuSeqno.Checked = True
'                .optBarSeq(1).Value = True
            ElseIf gHOSP.RSTTYPE = "2" Then
                .mnuRackPos.Checked = True
'                .optBarSeq(2).Value = True
            ElseIf gHOSP.RSTTYPE = "3" Then
                .mnuCheckBox.Checked = True
'                .optBarSeq(3).Value = True
            End If
        End If
        
        '-- 결과전송
        If gHOSP.SAVEAUTO = "Y" Then
            .mnuSaveAuto.Checked = True
            .mnuSaveManual.Checked = False
'            .optTrans(0).Value = True
        Else
            .mnuSaveAuto.Checked = False
            .mnuSaveManual.Checked = True
'            .optTrans(1).Value = True
        End If
        
        '-- 적용결과
        If gHOSP.SAVELIS = "Y" Then
            .mnuLisResult.Checked = True
            .mnuEqpResult.Checked = False
'            .optSaveResult(1).Value = True
        Else
            .mnuLisResult.Checked = False
            .mnuEqpResult.Checked = True
'            .optSaveResult(0).Value = True
        End If
        
        
    End With
    
    
End Sub

Public Sub SetCommStatus(ByVal pSRflag As String, ByVal pBarno As String, ByVal SPD As Object)
    
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
            .AddItem "Send" & vbTab & pBarno & vbTab & "장비전송"
            'Call SetText(SPD, "Send", .MaxRows, 1)
            'Call SetText(SPD, pBarNo, .MaxRows, 2)
            'Call SetText(SPD, "오더전송", .MaxRows, 3)
            
        ElseIf pSRflag = "Q" Then
            'Call SetText(SPD, "Recv", .MaxRows, 1)
            'Call SetText(SPD, pBarNo, .MaxRows, 2)
            'Call SetText(SPD, "오더요청", .MaxRows, 3)
            
            .AddItem "Recv" & vbTab & pBarno & vbTab & "오더요청"
        
        ElseIf pSRflag = "R" Then
            'Call SetText(SPD, "Recv", .MaxRows, 1)
            'Call SetText(SPD, pBarNo, .MaxRows, 2)
            'Call SetText(SPD, "결과수신", .MaxRows, 3)
            .AddItem "Recv" & vbTab & pBarno & vbTab & "결과수신"
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
            
            SPD.ColWidth(i + 1) = 10
            
            If varSize(i) <> "" Then
                SPD.ColWidth(i + 1) = varSize(i)
            End If
        
        Else
            SPD.ColHidden = True
        End If
    Next

End Sub

Public Sub SetColumnViewResult(ByVal SPD As Object)
    Dim i       As Integer
    Dim varSize As Variant
    Dim intViewSizeSum  As Integer
    
    intViewSizeSum = 0
    varSize = Split(gCOLSIZE_R, "|")

    For i = 0 To UBound(varSize) - 1
        SPD.Col = i + 1
        If Mid(gCOLVIEW_R, i + 1, 1) = 1 Then
            SPD.ColHidden = False
            
            SPD.ColWidth(i + 1) = 10
            
            If varSize(i) <> "" Then
                SPD.ColWidth(i + 1) = varSize(i)
                intViewSizeSum = intViewSizeSum + varSize(i)
            End If
        
        Else
            SPD.ColHidden = True
        End If
    Next

    SPD.WIDTH = intViewSizeSum * 150
    
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

Public Sub GetExeVersion()
    Dim i As Integer
    
    '-- HOSPITAl INFO GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("EXE", "HOSP", "", strSetup, 100, App.PATH & "\SANSOFT.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.MACHNM = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("EXE", "EMR", "", strSetup, 100, App.PATH & "\SANSOFT.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gEMR = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("EXE", "MACH", "", strSetup, 100, App.PATH & "\SANSOFT.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gMACH = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("EXE", "DBTYPE", "", strSetup, 100, App.PATH & "\SANSOFT.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gDBTYPE = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("EXE", "MACHCOUNT", "", strSetup, 100, App.PATH & "\SANSOFT.ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gMACHCOUNT = IIf(Trim(strSetUp1) = "", 0, Trim(strSetUp1))
    
    If IsNumeric(gMACHCOUNT) Then
        ReDim Preserve gMACHS(gMACHCOUNT) As String
        For i = 1 To gMACHCOUNT
            strSetup = "":    strSetUp1 = ""
            Call GetPrivateProfileString("EXE", "MACH" & CStr(i), "", strSetup, 100, App.PATH & "\SANSOFT.ini")
            strSetUp1 = Trim(strSetup)
            strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
            gMACHS(i) = Trim(strSetUp1)
        Next
    End If
    
End Sub

Public Sub GetSetup()
    Dim intVal As Integer
    
'''    '-- FORM INFO GET
'''    Call GetPrivateProfileString("FORM", "MAXYN", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gForm.MAXYN = Mid(strSetUp1, 1, InStr(strSetUp1, Chr(0)) - 1) 'Trim(strSetUp1)
'''
'''    Call GetPrivateProfileString("FORM", "TOP", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gForm.TOP = Mid(strSetUp1, 1, InStr(strSetUp1, Chr(0)) - 1) 'Trim(strSetUp1)
'''
'''    Call GetPrivateProfileString("FORM", "LEFT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gForm.LEFT = Mid(strSetUp1, 1, InStr(strSetUp1, Chr(0)) - 1) 'Trim(strSetUp1)
'''
'''    Call GetPrivateProfileString("FORM", "WIDTH", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gForm.WIDTH = Mid(strSetUp1, 1, InStr(strSetUp1, Chr(0)) - 1) 'Trim(strSetUp1)
'''
'''    Call GetPrivateProfileString("FORM", "HEIGHT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gForm.HEIGHT = Mid(strSetUp1, 1, InStr(strSetUp1, Chr(0)) - 1) 'Trim(strSetUp1)
'''
'''    '-- HOSPITAl INFO GET
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "HOSPCD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.HOSPCD = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "HOSPNM", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.HOSPNM = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "LABCD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.LABCD = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "LABNM", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.LABNM = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "PARTCD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.PARTCD = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "PARTNM", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.PARTNM = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "MACHCD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.MACHCD = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "MACHNM", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.MACHNM = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "USERID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.USERID = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "USERPW", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.USERPW = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "USERNM", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.USERNM = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "LOGINYN", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.LOGINYN = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "SAVEPW", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.SAVEPW = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "BARUSE", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.BARUSE = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "SAVELIS", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.SAVELIS = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "SAVEAUTO", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.SAVEAUTO = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "MENULOCK", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.MENULOCK = Trim(strSetUp1)
'''
'''    '-- 바코드 미사용시 결과받는 형태
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "RSTTYPE", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.RSTTYPE = Trim(strSetUp1)
'''
'''    '-- QC결과 저장경로
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "QCPATH", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.QCPATH = Trim(strSetUp1)
'''
'''    '-- ADVIA1800-2 장비코드
''''    strSetup = "":    strSetUp1 = ""
''''    Call GetPrivateProfileString("HOSP", "ADVIA1800", "", strSetup, 100, App.PATH & "\INI\" & gmach & ".ini")
''''    strSetUp1 = Trim(strSetup)
''''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
''''    gADV18CD = Trim(strSetUp1)
''''
''''    '-- CENTAURXP 장비코드
''''    strSetup = "":    strSetUp1 = ""
''''    Call GetPrivateProfileString("HOSP", "CENTAURXP", "", strSetup, 100, App.PATH & "\INI\" & gmach & ".ini")
''''    strSetUp1 = Trim(strSetup)
''''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
''''    gCENXPCD = Trim(strSetUp1)
'''
'''    '-- LOG 기록여부
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "LOGWRITE", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.LOQWRITE = Trim(strSetUp1)
'''
'''    '-- 워크리스트 조회화면
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "WORKTEST", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gWORKTEST = Trim(strSetUp1)
'''
'''    '-- 로컬저장기간
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "SAVEDAY", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.SAVEDAY = Trim(strSetUp1)
'''
'''    '-- 바코드길이
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "BARLEN", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.BARLEN = strSetUp1
'''
'''    '-- DB연결체크
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "DBCONCHK", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.DBCONCHK = strSetUp1
'''
'''    '-- Negative 표현
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "NEG", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.NEG = strSetUp1
'''
'''    '-- Positive 표현
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("HOSP", "POS", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHOSP.POS = strSetUp1
'''
'''
'''    '-- HOSPITAl INFO GET END
'''
'''    '-- OCS
''''    strSetup = "":    strSetUp1 = ""
''''    Call GetPrivateProfileString("HOSP", "OCS", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
''''    strSetUp1 = Trim(strSetup)
''''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
''''    gEMR = Trim(strSetUp1)
'''
'''    '-- DB TYPE GET
''''    strSetup = "":    strSetUp1 = ""
''''    Call GetPrivateProfileString("DATABASE", "DBTYPE", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
''''    strSetUp1 = Trim(strSetup)
''''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
''''    gDBTYPE = Trim(strSetUp1)
'''
'''    '-- LOCAL DB GET
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "LOCALPATH", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gLocalDB.PATH = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "LOCALUID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gLocalDB.UID = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "LOCALPWD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gLocalDB.PWD = Trim(strSetUp1)
'''
'''    '-- ORACLE DB GET
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "ORACLESID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gORADB.SID = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "ORACLEUID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gORADB.UID = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "ORACLEPWD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gORADB.PWD = Trim(strSetUp1)
'''
'''    '-- MSSQL DB GET
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "MSSQLIP", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gSQLDB.IP = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "MSSQLDB", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gSQLDB.DB = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "MSSQLUID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gSQLDB.UID = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "MSSQLPWD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gSQLDB.PWD = Trim(strSetUp1)
'''
'''    '-- PostGresSQL DB GET
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "PGSQLIP", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gPGSQLDB.IP = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "PGSQLDB", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gPGSQLDB.DB = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "PGSQLUID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gPGSQLDB.UID = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "PGSQLPWD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gPGSQLDB.PWD = Trim(strSetUp1)
'''
'''
'''    '-- MSSQL QC DB GET
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "MSSQLIP_QC", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gSQLDB_QC.IP = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "MSSQLDB_QC", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gSQLDB_QC.DB = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "MSSQLUID_QC", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gSQLDB_QC.UID = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("DATABASE", "MSSQLPWD_QC", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gSQLDB_QC.PWD = Trim(strSetUp1)
'''    '-- MSSQL QC DB GET END
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("VIEW", "COLWIDTH", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gCOLWIDTH = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("VIEW", "WORKPOS", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gWORKPOS = Trim(strSetUp1)
'''
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("VIEW", "SPDHEADER", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gCOLHEADER = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("VIEW", "SPDVIEW", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gCOLVIEW = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("VIEW", "SPDVIEW_R", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gCOLVIEW_R = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("VIEW", "SPDSIZE", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gCOLSIZE = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("VIEW", "SPDSIZE_R", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gCOLSIZE_R = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("CODE", "WBCM", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gUrinMic.WBCM = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("CODE", "RBCM", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gUrinMic.RBCM = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("CODE", "EPIC", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gUrinMic.EPIC = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("CODE", "BACT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gUrinMic.BACT = Trim(strSetUp1)
'''
'''
'''    '-- COMM INFO GET
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("COMM", "COMTYPE", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gComm.COMTYPE = Trim(strSetUp1)
'''
'''    If gComm.COMTYPE <> "" Then
'''        strSetup = "":    strSetUp1 = ""
'''        Call GetPrivateProfileString("COMM", "COMPORT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''        strSetUp1 = Trim(strSetup)
'''        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''        gComm.COMPORT = Trim(strSetUp1)
'''
'''        strSetup = "":    strSetUp1 = ""
'''        Call GetPrivateProfileString("COMM", "SPEED", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''        strSetUp1 = Trim(strSetup)
'''        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''        gComm.SPEED = Trim(strSetUp1)
'''
'''        strSetup = "":    strSetUp1 = ""
'''        Call GetPrivateProfileString("COMM", "PARITY", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''        strSetUp1 = Trim(strSetup)
'''        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''        gComm.Parity = Trim(strSetUp1)
'''
'''        strSetup = "":    strSetUp1 = ""
'''        Call GetPrivateProfileString("COMM", "DATABIT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''        strSetUp1 = Trim(strSetup)
'''        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''        gComm.DATABIT = Trim(strSetUp1)
'''
'''        strSetup = "":    strSetUp1 = ""
'''        Call GetPrivateProfileString("COMM", "STARTBIT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''        strSetUp1 = Trim(strSetup)
'''        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''        gComm.STARTBIT = Trim(strSetUp1)
'''
'''        strSetup = "":    strSetUp1 = ""
'''        Call GetPrivateProfileString("COMM", "STOPBIT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''        strSetUp1 = Trim(strSetup)
'''        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''        gComm.STOPBIT = Trim(strSetUp1)
'''
'''        strSetup = "":    strSetUp1 = ""
'''        Call GetPrivateProfileString("COMM", "RTSEnable", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''        strSetUp1 = Trim(strSetup)
'''        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''        gComm.RTSEnable = Trim(strSetUp1)
'''
'''        strSetup = "":    strSetUp1 = ""
'''        Call GetPrivateProfileString("COMM", "DTREnable", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''        strSetUp1 = Trim(strSetup)
'''        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''        gComm.DTREnable = Trim(strSetUp1)
'''
'''        strSetup = "":    strSetUp1 = ""
'''        Call GetPrivateProfileString("COMM", "TCPTYPE", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''        strSetUp1 = Trim(strSetup)
'''        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''        gComm.TCPTYPE = Trim(strSetUp1)
'''
'''        strSetup = "":    strSetUp1 = ""
'''        Call GetPrivateProfileString("COMM", "TCPIP", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''        strSetUp1 = Trim(strSetup)
'''        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''        gComm.TCPIP = Trim(strSetUp1)
'''
'''        strSetup = "":    strSetUp1 = ""
'''        Call GetPrivateProfileString("COMM", "TCPPORT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''        strSetUp1 = Trim(strSetup)
'''        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''        gComm.TCPPORT = Trim(strSetUp1)
'''
'''        strSetup = "":    strSetUp1 = ""
'''        Call GetPrivateProfileString("COMM", "RSTPATH", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''        strSetUp1 = Trim(strSetup)
'''        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''        gComm.RSTPATH = Trim(strSetUp1)
'''
'''
'''    End If
'''
'''    '-- JSON
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("URL", "WORKLIST", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gURL.WORKLIST = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("URL", "PATLIST", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gURL.PATLIST = Trim(strSetUp1)
'''
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("URL", "PATSAVE", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gURL.PATSAVE = Trim(strSetUp1)
'''
'''    '-- 보건소
'''    strSetup = "":    strSetUp1 = ""
'''    Call GetPrivateProfileString("URL", "INITURL", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'''    strSetUp1 = Trim(strSetup)
'''    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'''    gHEALTH.INITURL = Trim(strSetUp1)
    
    '-- FORM INFO
    gForm.MAXYN = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "MAXYN")
    gForm.TOP = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "TOP")
    gForm.LEFT = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "LEFT")
    gForm.WIDTH = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "WIDTH")
    gForm.HEIGHT = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "HEIGHT")

    '-- HOSPITAl INFO
    gHOSP.HOSPCD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "HOSPCD")
    gHOSP.HOSPNM = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "HOSPNM")
    gHOSP.LABCD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "LABCD")
    gHOSP.LABNM = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "LABNM")
    gHOSP.PARTCD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "PARTCD")
    gHOSP.PARTNM = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "PARTNM")
    gHOSP.MACHCD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "MACHCD")
    gHOSP.MACHNM = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "MACHNM")
    gHOSP.USERID = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "USERID")
    gHOSP.USERPW = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "USERPW")
    gHOSP.USERNM = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "USERNM")
    gHOSP.LOGINYN = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "LOGINYN")
    gHOSP.SAVEPW = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVEPW")
    gHOSP.BARUSE = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "BARUSE")
    gHOSP.SAVELIS = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVELIS")
    gHOSP.SAVEAUTO = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVEAUTO")
    gHOSP.MENULOCK = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "MENULOCK")
    gHOSP.RSTTYPE = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "RSTTYPE")
    gHOSP.LOQWRITE = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "LOQWRITE")
    gHOSP.QCPATH = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "QCPATH")        '-- QC결과 저장경로
    gHOSP.SAVEDAY = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVEDAY")      '-- 로컬저장기간
    gHOSP.BARLEN = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "BARLEN")        '-- 바코드길이
    gHOSP.DBCONCHK = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "DBCONCHK")    '-- DB연결체크
    
    gWORKTEST = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "WORKTEST")         '-- 워크리스트 조회화면
    gWORKPOS = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "WORKPOS")

    gHOSP.NEG = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "NEG")              '-- Negative 표현
    gHOSP.POS = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "POS")              '-- Positive 표현
    
    '-- LOCAL DB GET
    gLocalDB.PATH = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MDBPATH")
    gLocalDB.UID = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MDBUID")
    gLocalDB.PWD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MDBPWD")
    '-- ORACLE DB GET
    gORADB.SID = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "ORACLESID")
    gORADB.UID = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "ORACLEUID")
    gORADB.PWD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "ORACLEPWD")
    '-- MSSQL DB GET
    gSQLDB.IP = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MSSQLIP")
    gSQLDB.DB = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MSSQLDB")
    gSQLDB.UID = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MSSQLUID")
    gSQLDB.PWD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MSSQLPWD")
    '-- PostGresSQL DB GET
    gPGSQLDB.IP = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "PGSQLIP")
    gPGSQLDB.DB = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "PGSQLDB")
    gPGSQLDB.UID = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "PGSQLUID")
    gPGSQLDB.PWD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "PGSQLPWD")
    '-- MSSQL QC DB GET
    gSQLDB_QC.IP = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MSSQLIP_QC")
    gSQLDB_QC.DB = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MSSQLDB_QC")
    gSQLDB_QC.UID = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MSSQLUID_QC")
    gSQLDB_QC.PWD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MSSQLPWD_QC")

    '-- VIEW
'    gWORKPOS = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "WORKPOS")
    gCOLWIDTH = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "COLWIDTH")
    gCOLHEADER = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "SPDHEADER")
    gCOLVIEW = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "SPDVIEW")
    gCOLVIEW_R = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "SPDVIEW_R")
    gCOLSIZE = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "SPDSIZE")
    gCOLSIZE_R = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "SPDSIZE_R")
    gROWHEIGHT = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "ROWHEIGHT")
    
    '-- COMM INFO GET
    gComm.COMTYPE = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "COMTYPE")
    gComm.COMPORT = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "COMPORT")
    gComm.SPEED = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "SPEED")
    gComm.Parity = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "PARITY")
    gComm.DATABIT = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "DATABIT")
    gComm.STARTBIT = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "STARTBIT")
    gComm.STOPBIT = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "STOPBIT")
    gComm.RTSEnable = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "RTSEnable")
    gComm.DTREnable = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "DTREnable")
    gComm.TCPTYPE = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "TCPTYPE")
    gComm.TCPIP = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "TCPIP")
    gComm.TCPPORT = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "TCPPORT")
    gComm.RSTPATH = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "RSTPATH")
    
    '-- URL (JSON)
    strSetup = "":    strSetUp1 = ""
    gURL.WORKLIST = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "URL", "WORKLIST")
    gURL.PATLIST = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "URL", "PATLIST")
    gURL.PATSAVE = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "URL", "PATSAVE")
    '-- URL (보건소)
    gHEALTH.INITURL = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "URL", "INITURL")
    
    
End Sub

'FTP 전송하기 ============================================================
Public Sub SetBackup() '(ByVal pFilePath As String, ByVal pFileName As String)
    Dim okConn      As Boolean
    Dim okFTrans    As Boolean
    
    '-- FTP OPEN
    okConn = mmftp.OpenConnection(gFTP.SERVER, gFTP.port, gFTP.UID, gFTP.PWD)

    '-- 최상위 이동
    okConn = mmftp.SetFTPDirectory("/")
    
    '-- 해당 폴더 이동
    If mmftp.SetFTPDirectory("/san/interface/" & gHOSP.HOSPNM) = False Then
        '폴더가 없으면 만든다.
        If mmftp.CreateFTPDirectory("/san/interface/" & gHOSP.HOSPNM) Then
            Call mmftp.SetFTPDirectory("/san/interface/" & gHOSP.HOSPNM)
        End If
    End If
    
    
    If mmftp.SetFTPDirectory("/san/interface/" & gHOSP.HOSPNM & "/" & gHOSP.MACHNM) = False Then
        '폴더가 없으면 만든다.
        If mmftp.CreateFTPDirectory("/san/interface/" & gHOSP.HOSPNM & "/" & gHOSP.MACHNM) Then
            Call mmftp.SetFTPDirectory("/san/interface/" & gHOSP.HOSPNM & "/" & gHOSP.MACHNM)
        End If
    End If
    
    '-- FTP WRITE(백업)
    'okFTrans = mmftp.RenameFTPFile("/san/interface/" & gHOSP.HOSPNM & "/" & gHOSP.MACHNM & "/" & gHOSP.MACHNM & ".mdb", "/san/interface/" & gHOSP.HOSPNM & "/" & gHOSP.MACHNM & "/" & gHOSP.MACHNM & ".bak")
    
    '-- FTP SEND
    'okFTrans = mmftp.FTPUploadFile(pFilePath & pFileName, pFileName)

    '-- CLOSE
    mmftp.CloseConnection
    '=========================================================================

End Sub

':: 레지스트리에서 인터페이스 설정을 읽어온다
Public Sub GetRegSetup()
    
    '-- FTP 설정 [백업용]
    gFTP.SERVER = GetString(HKEY_CURRENT_USER, REG_POSITION, "FTP_SERVER")
    gFTP.port = IIf(GetString(HKEY_CURRENT_USER, REG_POSITION, "FTP_PORT") = "", 0, GetString(HKEY_CURRENT_USER, REG_POSITION, "FTP_PORT"))
    gFTP.UID = GetString(HKEY_CURRENT_USER, REG_POSITION, "FTP_UID")
    gFTP.PWD = GetString(HKEY_CURRENT_USER, REG_POSITION, "FTP_PWD")
    
    If gFTP.SERVER = "" Then gFTP.SERVER = "san.i234.me"
    If gFTP.port = 0 Then gFTP.port = 20021
    If gFTP.UID = "" Then gFTP.UID = "test1"
    If gFTP.PWD = "" Then gFTP.PWD = "123456"

    '-- FORM INFO
    gForm.MAXYN = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "MAXYN")
    gForm.TOP = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "TOP")
    gForm.LEFT = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "LEFT")
    gForm.WIDTH = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "WIDTH")
    gForm.HEIGHT = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "HEIGHT")

    '-- HOSPITAl INFO
    gHOSP.HOSPCD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "HOSPCD")
    gHOSP.HOSPNM = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "HOSPNM")
    gHOSP.LABCD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "LABCD")
    gHOSP.LABNM = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "LABNM")
    gHOSP.PARTCD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "PARTCD")
    gHOSP.PARTNM = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "PARTNM")
    gHOSP.MACHCD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "MACHCD")
    gHOSP.MACHNM = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "MACHNM")
    gHOSP.USERID = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "USERID")
    gHOSP.USERPW = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "USERPW")
    gHOSP.USERNM = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "USERNM")
    gHOSP.LOGINYN = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "LOGINYN")
    gHOSP.SAVEPW = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVEPW")
    gHOSP.BARUSE = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "BARUSE")
    gHOSP.SAVELIS = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVELIS")
    gHOSP.SAVEAUTO = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVEAUTO")
    gHOSP.MENULOCK = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "MENULOCK")
    gHOSP.RSTTYPE = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "RSTTYPE")
    gHOSP.LOQWRITE = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "LOQWRITE")
    gHOSP.QCPATH = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "QCPATH")        '-- QC결과 저장경로
    gHOSP.SAVEDAY = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVEDAY")      '-- 로컬저장기간
    gHOSP.BARLEN = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "BARLEN")        '-- 바코드길이
    gHOSP.DBCONCHK = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "DBCONCHK")    '-- DB연결체크

    gWORKTEST = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "WORKTEST")         '-- 워크리스트 조회화면
    gWORKPOS = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "WORKPOS")
    
    gHOSP.NEG = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "NEG")              '-- Negative 표현
    gHOSP.POS = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "POS")              '-- Positive 표현
    
    '-- LOCAL DB GET
    gLocalDB.PATH = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MDBPATH")
    gLocalDB.UID = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MDBUID")
    gLocalDB.PWD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MDBPWD")
    '-- ORACLE DB GET
    gORADB.SID = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "ORACLESID")
    gORADB.UID = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "ORACLEUID")
    gORADB.PWD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "ORACLEPWD")
    '-- MSSQL DB GET
    gSQLDB.IP = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MSSQLIP")
    gSQLDB.DB = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MSSQLDB")
    gSQLDB.UID = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MSSQLUID")
    gSQLDB.PWD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MSSQLPWD")
    '-- PostGresSQL DB GET
    gPGSQLDB.IP = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "PGSQLIP")
    gPGSQLDB.DB = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "PGSQLDB")
    gPGSQLDB.UID = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "PGSQLUID")
    gPGSQLDB.PWD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "PGSQLPWD")
    '-- MSSQL QC DB GET
    gSQLDB_QC.IP = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MSSQLIP_QC")
    gSQLDB_QC.DB = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MSSQLDB_QC")
    gSQLDB_QC.UID = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MSSQLUID_QC")
    gSQLDB_QC.PWD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MSSQLPWD_QC")

    '-- VIEW
    'gWORKPOS = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "WORKPOS")
    gCOLWIDTH = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "COLWIDTH")
    gCOLHEADER = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "SPDHEADER")
    gCOLVIEW = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "SPDVIEW")
    gCOLVIEW_R = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "SPDVIEW_R")
    gCOLSIZE = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "SPDSIZE")
    gCOLSIZE_R = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "SPDSIZE_R")
    gROWHEIGHT = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "ROWHEIGHT")
    gDETAILVIEW = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "DETAILVIEW")

    '-- COMM INFO GET
    gComm.COMTYPE = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "COMTYPE")
    gComm.COMPORT = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "COMPORT")
    gComm.SPEED = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "SPEED")
    gComm.Parity = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "PARITY")
    gComm.DATABIT = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "DATABIT")
    gComm.STARTBIT = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "STARTBIT")
    gComm.STOPBIT = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "STOPBIT")
    gComm.RTSEnable = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "RTSEnable")
    gComm.DTREnable = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "DTREnable")
    gComm.TCPTYPE = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "TCPTYPE")
    gComm.TCPIP = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "TCPIP")
    gComm.TCPPORT = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "TCPPORT")
    gComm.RSTPATH = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "COMM", "RSTPATH")
    
    '-- URL (JSON)
    gURL.WORKLIST = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "URL", "WORKLIST")
    gURL.PATLIST = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "URL", "PATLIST")
    gURL.PATSAVE = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "URL", "PATSAVE")
    '-- URL (보건소)
    gHEALTH.INITURL = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "URL", "INITURL")
    
End Sub

Public Sub SetExamCode(ByVal SPD As Object)
    Dim i As Integer
    
    With SPD
        .MaxCols = colSTATE + UBound(gArrEQP)
        For i = 0 To UBound(gArrEQP) - 1
            .Col = colSTATE + (i + 1)
            .Row = -1
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            .ColWidth(colSTATE + (i + 1)) = gCOLWIDTH
            .FontBold = False
            
            '5 : 약어명으로 표시한다.
            Call SetText(SPD, Trim(gArrEQP(i + 1, 6)), 0, colSTATE + (i + 1))
            
            '계산식 여부
            If gArrEQP((i + 1), 14) = "" Or gArrEQP((i + 1), 14) = "0" Then
                .FontBold = False
            Else
                .FontBold = True    '계산식 검사일 경우 굵게 표현한다.
            End If
        Next
    End With
    
End Sub

Public Sub SaveExcel(Filename As String, argSpread As Object)
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
    
    SP.UserColAction = iSortoption 'UserColActionSort
   
End Sub



Public Function KillProcess(PNAME As String) As Boolean

    On Error GoTo onError
    
    Dim wmi As Object
    Dim Processes, Process
    Dim sQuery As String
    
    Set wmi = GetObject("winmgmts:")
    sQuery = "select * from win32_process where namae = '" & PNAME & "'"
    Set Processes = wmi.execquery(sQuery)
    
    For Each Process In Processes
        Process.Terminate
    Next
    
    Set wmi = Nothing
    
    KillProcess = True
    Exit Function
    
Exit Function
onError:
    KillProcess = False
End Function


Public Sub spdPopUpDel(ByVal spdObj As Object, ByVal pCol As Long, ByVal pRow As Long, ByVal pIdx As Integer)
    Dim oMenu       As clsPopUp
    Dim lMenuChosen As Long
    Dim strVal1     As String
    Dim strVal2     As String
    Dim strVal3     As String
    Dim strVal4     As String
    Dim strPName    As String
    Dim intRow      As Integer
    Dim intCnt      As Integer
    Dim blnCheck    As Boolean
    
    intCnt = 0
    blnCheck = False
    With frmInterface.spdOrder
        For intRow = 1 To .DataRowCnt
            .Row = intRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                intCnt = intCnt + 1
                blnCheck = True
                'Exit For
            End If
        Next
    End With
    
    If intCnt >= 2 Then
        MsgBox "우측 리스트에서 한개의 검체만 선택하세요", vbOKOnly + vbCritical, "검체 선택"
        Exit Sub
    End If
    
    If blnCheck = False Then
        MsgBox "우측 리스트에서 한개의 검체를 선택하세요", vbOKOnly + vbCritical, "검체 선택"
        Exit Sub
    End If
    
    Set oMenu = New clsPopUp
    
    strPName = GetText(spdObj, pRow, colPNAME)
    
    lMenuChosen = oMenu.Popup("▒ " & strPName & " 검사매치")

    Select Case lMenuChosen
        Case 1
            If MsgBox("선택하신 검체를 검사내역검체에 매치하시겠습니까?", vbYesNo + vbInformation, "매치") = vbYes Then
                Exit Sub
                With spdObj
                    .Row = pRow
                    Select Case pIdx
                        Case 1:  .Col = 1: strVal1 = Trim(.Text)
                        Case 2:  .Col = 1: strVal1 = Trim(.Text)
                        Case 3:  .Col = 1: strVal1 = Trim(.Text)
                                 .Col = 2: strVal2 = Trim(.Text)
                        Case 4:  .Col = 1: strVal1 = Trim(.Text)
                                 .Col = 2: strVal2 = Trim(.Text)
                                 .Col = 3: strVal3 = Trim(.Text)
                        Case 5:  .Col = 1: strVal1 = Trim(.Text)
                                 .Col = 2: strVal2 = Trim(.Text)
                        Case 6:  .Col = 1: strVal1 = Trim(.Text)
                                 .Col = 2: strVal2 = Trim(.Text)
                        Case 8:  .Col = 1: strVal1 = Trim(.Text)
                                 .Col = 2: strVal2 = Trim(.Text)
                                 .Col = 4: strVal3 = Trim(.Text)
                                 strVal3 = mGetP(mGetP(strVal3, 1, "]"), 2, "[")
                                 .Col = 5: strVal4 = Trim(.Text)
                        Case 9:  .Col = 1: strVal1 = Trim(.Text)
                        Case 10: .Col = 1: strVal1 = Trim(.Text)
                        Case 11: .Col = 1: strVal1 = Trim(.Text)
                                 .Col = 2: strVal2 = Trim(.Text)
                                 .Col = 4: strVal3 = Trim(.Text)
                                 .Col = 6: strVal4 = Trim(.Text)
                        Case 12: .Col = 3: strVal1 = Trim(.Text)
                                 .Col = 4: strVal2 = Trim(.Text)
                                 .Col = 5: strVal3 = Trim(.Text)
                                 .Col = 7: strVal4 = Trim(.Text)
                    End Select
                    
                    'If cMstDelete(pIdx, strVal1, strVal2, strVal3, strVal4) Then
                        .Action = ActionDeleteRow
                        .MaxRows = .MaxRows - 1
                    'End If
                End With
            End If
    End Select

End Sub

