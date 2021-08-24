Attribute VB_Name = "modCommon"
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                    (ByVal lpApplicationName As String _
                   , ByVal lpKeyName As Any _
                   , ByVal lpString As Any _
                   , ByVal lplFileName As String) As Long
    
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                    (ByVal lpApplicationName As String _
                   , ByVal lpKeyName As Any _
                   , ByVal lpDefault As String _
                   , ByVal lpReturnedString As String _
                   , ByVal nSize As Long _
                   , ByVal lpFileName As String) As Long
                    
                    
Public gEMR         As String
Public gMACH        As String
Public gMACHCOUNT   As Integer
Public gMACHS()     As String
                    
                    
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


Public strSetup     As String * 100
Public strSetUp1    As String

Public gArrEQP()    As String
Public gArrEQPNm()  As String   '-- 인터페이스에 등록된 전체검사명
Public gAllTestCd   As String   '-- 인터페이스에 등록된 전체검사코드
Public gAllOrdCd    As String   '-- 인터페이스에 등록된 전체오더코드
Public gPatOrdCd    As String   '-- 검체별 검사코드
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

Public gCOLVIEW_R        As String
Public gCOLSIZE_R        As String

Public gWORKPOS          As String
Public gWORKTEST         As String

Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long


Sub Main()
        
    '두번 실행 하지 않음
    If App.PrevInstance Then
       MsgBox "프로그램이 이미 실행중입니다.!", vbExclamation
       End
    End If
    
    '-- INI 정보
    Call GetExeVersion
    
    '-- 사용장비
    If gMACH = "" Then
        frmEMRInfo.Show vbModal
    End If
    
    '-- EMR
    If gEMR = "" Then
        frmEMRInfo.Show vbModal
    End If
    
    '-- INI 정보
    Call GetSetup
    
    '-- 병원코드
    If Len(gHOSP.HOSPCD) = 0 Then
        frmHospInfo.Show vbModal
    End If
    
    
'''    If Len(gLocalDB.PATH) = 0 Then
'''        frmDB_Local.Show vbModal
'''    End If
'''
'''    If gDBTYPE = "1" Then
'''        If Len(gORADB.SID) = 0 Then
'''            frmDB_Oracle.Show vbModal
'''        End If
'''    ElseIf gDBTYPE = "2" Then
'''        If Len(gSQLDB.IP) = 0 Then
'''            frmDB_MSSQL.Show vbModal
'''        End If
'''    ElseIf gDBTYPE = "3" Then
'''        If Len(gPGSQLDB.IP) = 0 Then
'''            frmDB_MSSQL.Show vbModal
'''        End If
'''    ElseIf gDBTYPE = "99" Then
'''
'''    Else
'''        MsgBox App.PATH & "\SANSOFT.ini 파일에서" & vbNewLine & vbNewLine & "DBTYPE을 먼저 설정하세요 ", vbOKOnly + vbInformation, "DB TYPE 설정"
'''        End
'''    End If
    
    
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


''        If Not DbConnect_SQL_QC Then
''            MsgBox "QC 데이터베이스가 없습니다.", vbCritical, "QC"
''        End If
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
    
End Sub

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
    frm.WindowState = 2
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
    gMACHCOUNT = Trim(strSetUp1)
    
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
    
    '-- FORM INFO GET
    Call GetPrivateProfileString("FORM", "MAXYN", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gForm.MAXYN = Trim(strSetUp1)
    
    Call GetPrivateProfileString("FORM", "TOP", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gForm.TOP = Trim(strSetUp1)
    
    Call GetPrivateProfileString("FORM", "LEFT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gForm.LEFT = Trim(strSetUp1)
    
    Call GetPrivateProfileString("FORM", "WIDTH", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gForm.WIDTH = Trim(strSetUp1)
    
    Call GetPrivateProfileString("FORM", "HEIGHT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gForm.HEIGHT = Trim(strSetUp1)
    
    '-- HOSPITAl INFO GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "HOSPCD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.HOSPCD = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "HOSPNM", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.HOSPNM = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "LABCD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.LABCD = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "LABNM", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.LABNM = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "PARTCD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.PARTCD = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "PARTNM", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.PARTNM = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "MACHCD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.MACHCD = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "MACHNM", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.MACHNM = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "USERID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.USERID = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "USERPW", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.USERPW = Trim(strSetUp1)
        
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "USERNM", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.USERNM = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "LOGINYN", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.LOGINYN = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "SAVEPW", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.SAVEPW = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "BARUSE", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.BARUSE = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "SAVELIS", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.SAVELIS = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "SAVEAUTO", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.SAVEAUTO = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "MENULOCK", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.MENULOCK = Trim(strSetUp1)
    
    '-- 바코드 미사용시 결과받는 형태
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "RSTTYPE", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.RSTTYPE = Trim(strSetUp1)
    
    '-- QC결과 저장경로
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "QCPATH", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.QCPATH = Trim(strSetUp1)
    
    '-- ADVIA1800-2 장비코드
'    strSetup = "":    strSetUp1 = ""
'    Call GetPrivateProfileString("HOSP", "ADVIA1800", "", strSetup, 100, App.PATH & "\INI\" & gmach & ".ini")
'    strSetUp1 = Trim(strSetup)
'    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'    gADV18CD = Trim(strSetUp1)
'
'    '-- CENTAURXP 장비코드
'    strSetup = "":    strSetUp1 = ""
'    Call GetPrivateProfileString("HOSP", "CENTAURXP", "", strSetup, 100, App.PATH & "\INI\" & gmach & ".ini")
'    strSetUp1 = Trim(strSetup)
'    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'    gCENXPCD = Trim(strSetUp1)
    
    '-- LOG 기록여부
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "LOGWRITE", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.LOQWRITE = Trim(strSetUp1)
    
    '-- 워크리스트 조회화면
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "WORKTEST", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gWORKTEST = Trim(strSetUp1)
    
    '-- 로컬저장기간
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "SAVEDAY", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.SAVEDAY = Trim(strSetUp1)
    
    '-- 바코드길이
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "BARLEN", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.BARLEN = strSetUp1
    
    '-- DB연결체크
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("HOSP", "DBCONCHK", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gHOSP.DBCONCHK = strSetUp1


    
    '-- HOSPITAl INFO GET END
    
    '-- OCS
'    strSetup = "":    strSetUp1 = ""
'    Call GetPrivateProfileString("HOSP", "OCS", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'    strSetUp1 = Trim(strSetup)
'    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'    gEMR = Trim(strSetUp1)
    
    '-- DB TYPE GET
'    strSetup = "":    strSetUp1 = ""
'    Call GetPrivateProfileString("DATABASE", "DBTYPE", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
'    strSetUp1 = Trim(strSetup)
'    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
'    gDBTYPE = Trim(strSetUp1)
    
    '-- LOCAL DB GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "LOCALPATH", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gLocalDB.PATH = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "LOCALUID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gLocalDB.UID = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "LOCALPWD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gLocalDB.PWD = Trim(strSetUp1)

    '-- ORACLE DB GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "ORACLESID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gORADB.SID = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "ORACLEUID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gORADB.UID = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "ORACLEPWD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gORADB.PWD = Trim(strSetUp1)

    '-- MSSQL DB GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "MSSQLIP", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB.IP = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "MSSQLDB", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB.DB = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "MSSQLUID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB.UID = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "MSSQLPWD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB.PWD = Trim(strSetUp1)

    '-- PostGresSQL DB GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "PGSQLIP", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gPGSQLDB.IP = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "PGSQLDB", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gPGSQLDB.DB = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "PGSQLUID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gPGSQLDB.UID = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "PGSQLPWD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gPGSQLDB.PWD = Trim(strSetUp1)


    '-- MSSQL QC DB GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "MSSQLIP_QC", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB_QC.IP = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "MSSQLDB_QC", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB_QC.DB = Trim(strSetUp1)

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "MSSQLUID_QC", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB_QC.UID = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("DATABASE", "MSSQLPWD_QC", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gSQLDB_QC.PWD = Trim(strSetUp1)
    '-- MSSQL QC DB GET END

    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("VIEW", "COLWIDTH", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gCOLWIDTH = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("VIEW", "WORKPOS", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gWORKPOS = Trim(strSetUp1)
    
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("VIEW", "SPDHEADER", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gCOLHEADER = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("VIEW", "SPDVIEW", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gCOLVIEW = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("VIEW", "SPDVIEW_R", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gCOLVIEW_R = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("VIEW", "SPDSIZE", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gCOLSIZE = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("VIEW", "SPDSIZE_R", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gCOLSIZE_R = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("CODE", "WBCM", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gUrinMic.WBCM = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("CODE", "RBCM", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gUrinMic.RBCM = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("CODE", "EPIC", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gUrinMic.EPIC = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("CODE", "BACT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gUrinMic.BACT = Trim(strSetUp1)


    '-- COMM INFO GET
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("COMM", "COMTYPE", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gComm.COMTYPE = Trim(strSetUp1)
    
    If gComm.COMTYPE <> "" Then
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("COMM", "COMPORT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gComm.COMPORT = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("COMM", "SPEED", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gComm.SPEED = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("COMM", "PARITY", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gComm.Parity = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("COMM", "DATABIT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gComm.DATABIT = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("COMM", "STARTBIT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gComm.STARTBIT = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("COMM", "STOPBIT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gComm.STOPBIT = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("COMM", "RTSEnable", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gComm.RTSEnable = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("COMM", "DTREnable", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gComm.DTREnable = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("COMM", "TCPTYPE", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gComm.TCPTYPE = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("COMM", "TCPIP", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gComm.TCPIP = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("COMM", "TCPPORT", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gComm.TCPPORT = Trim(strSetUp1)
        
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("COMM", "RSTPATH", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gComm.RSTPATH = Trim(strSetUp1)
    

    End If
    
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("URL", "WORKLIST", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gURL.WORKLIST = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("URL", "PATLIST", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gURL.PATLIST = Trim(strSetUp1)
    
    strSetup = "":    strSetUp1 = ""
    Call GetPrivateProfileString("URL", "PATSAVE", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
    strSetUp1 = Trim(strSetup)
    strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
    gURL.PATSAVE = Trim(strSetUp1)
    
    
    
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
            Call SetText(SPD, Trim(gArrEQP(i + 1, 6)), 0, colSTATE + (i + 1))    '-- 5 : 약어명
            .ColWidth(colSTATE + (i + 1)) = gCOLWIDTH
            .FontBold = False 'True
            
            If gArrEQP((i + 1), 14) = "" Or gArrEQP((i + 1), 14) = "0" Then
                .FontBold = False
                '.ForeColor = vbYellow
            Else
                .FontBold = True
                '.ForeColor = vbRed
            End If
        Next
    End With
    
'    With SPD
'        .MaxCols = colSTATE + UBound(gArrEQPNm)
'        For i = 0 To UBound(gArrEQPNm) - 1
'            .Col = colSTATE + (i + 1)
'            .Row = -1
'            .CellType = CellTypeStaticText
'            .TypeHAlign = TypeHAlignCenter
'            .TypeVAlign = TypeVAlignCenter
'            Call SetText(SPD, Trim(gArrEQPNm(i + 1, 6)), 0, colSTATE + (i + 1))   '-- 5 : 약어명
'            .ColWidth(colSTATE + (i + 1)) = gCOLWIDTH
'        Next
'    End With
    
    
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
