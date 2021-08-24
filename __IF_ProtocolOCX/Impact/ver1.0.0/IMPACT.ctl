VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl IMPACT 
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   LockControls    =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   3330
   Begin VB.CommandButton cmdTest 
      Caption         =   "TEST"
      Height          =   375
      Left            =   210
      TabIndex        =   1
      Top             =   1725
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   1395
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   0
      Top             =   135
      Width           =   1365
   End
   Begin MSCommLib.MSComm msComm 
      Left            =   255
      Top             =   2370
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "IMPACT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_EqName = "0"
Const m_def_bUseBarcode = 0
Const m_def_iPhase = 0
Const m_def_iSendPhase = 0
Const m_def_sTestMode = "0"
Const m_def_iFrameN = 0
Const m_def_p_sID = "0"
Const m_def_p_sSeq = "0"
Const m_def_p_sRack = "0"
Const m_def_p_sPos = "0"
Const m_def_p_iOrdCnt = 0
Const m_def_p_sTIFCd = "0"
Const m_def_PortOpen = 0
Const m_def_OpenPW = "0"
Const m_def_EditPW = "0"
'속성 변수:
Dim m_EqName As String
Dim m_bUseBarcode As Boolean
Dim m_iPhase As Integer
Dim m_iSendPhase As Integer
Dim m_sTestMode As String
Dim m_iFrameN As Integer
Dim m_p_sID As String
Dim m_p_sSeq As String
Dim m_p_sRack As String
Dim m_p_sPos As String
Dim m_p_iOrdCnt As Integer
Dim m_p_sTIFCd As String
Dim m_PortOpen As Boolean
Dim m_OpenPW As String
Dim m_EditPW As String
'이벤트 선언:
Event RequestCurOrder(sID$, sSeq$, sRack$, sPos$)
Event SendOrderOK(sID$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
'Event RequestCurOrder(sID$, sRack$, sPos$)
Event DispMsg(sMsg$)
Event RequestNextOrder()
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)


'===== User Define
'인터페이스에서 사용
Dim RcvBuffer   As String
Dim wkBuf   As String
Dim sState  As String
Dim sReqStatusCd    As String

'구조체 지정
Private pSampleInfo As SAMPLE_INFO
Private pResultInfo As RESULT_INFO

'기타
Dim iSpaceCnt   As Integer

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=msComm,msComm,-1,CommPort
Public Property Get CommPort() As Integer
Attribute CommPort.VB_Description = "통신 포트 번호를 반환하거나 설정합니다."
    CommPort = msComm.CommPort
End Property

Public Property Let CommPort(ByVal New_CommPort As Integer)
    msComm.CommPort() = New_CommPort
    PropertyChanged "CommPort"
End Property

Private Sub PhaseCfg_Protocol()

    '--- 사용자 확인
    If m_EditPW <> pEditPW Then
        MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
        Exit Sub
    End If
    '---------------
    
    If m_EqName = "0" Or m_EqName = "" Then
        RaiseEvent DispMsg("검사장비명을 지정해 주십시오.!!!")
        Exit Sub
    End If
    
    Select Case UCase(m_EqName)
        Case "IMPACT"
            Call PhaseCfg_Protocol_IMPACT
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_Elecsys1010()
'    On Error GoTo ErrRtn
'
'    Dim RecType As String   'Record Type
'    Dim i       As Integer
'    Dim tmpData()   As String
'    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$
'    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpRef$, tmpFlag$
'
'    RecType = Mid$(RcvBuffer, 2, 1)
'
'    Select Case RecType
'        Case "H"        'Header Record
'            sState = ""
'
'        Case "M"
'        Case "P"        'Patient Record
'            Call Init_pResultInfo
'
'        Case "Q"        'Order Request Record
'            tmpData() = Split(RcvBuffer, "|")
'            sReqStatusCd = Left(tmpData(12), 1)     'Order Request Status Code
'            tmpData() = Split(tmpData(2), "^")
'
'            tmpBarCd = Trim(tmpData(1))
'            tmpSeqNo = Trim(tmpData(2))
'            tmpRack = Trim(tmpData(3))
'            tmpPos = Trim(tmpData(4))
'
'            If tmpBarCd = "" Then           'BarCode ID가 잘 넘어왔는지 검사
'                sState = ""
'                sReqStatusCd = ""
'                pSampleInfo.ID = ""
'            End If
'
'            sState = "Q"
'            pSampleInfo.ID = tmpBarCd        'BarCode
'            pSampleInfo.SEQNO = tmpSeqNo
'            pSampleInfo.RACK = tmpRack
'            pSampleInfo.POS = tmpPos
'
'        Case "O"
'            tmpData() = Split(RcvBuffer, "|")
'            tmpBarCd = Trim(tmpData(2))
'            tmpData() = Split(tmpData(3), "^")
'            tmpSeqNo = Trim(tmpData(0))
'            tmpRack = Trim(tmpData(1))
'            tmpPos = Trim(tmpData(2))
'
'            pSampleInfo.ID = tmpBarCd
'            pSampleInfo.RACK = tmpRack
'            pSampleInfo.POS = tmpPos
'
'        Case "R"        'Result Record
'            '--- 결과데이타 편집
'            'tmpData(2): TESTCD
'            '    "  (3): RESULT
'            '    "  (4): UNIT
'            '    "  (5): 참고치 범위
'            '    "  (6): 참고치(N:Neg/H:Pos)
'            tmpData() = Split(RcvBuffer, "|")
'            If Trim(tmpData(2)) <> "" Then
'                tmpData(2) = Mid(tmpData(2), 4)
'            End If
'            i = InStr(tmpData(2), "^")
'            If i <> 0 Then
'                tmpData(2) = Mid(tmpData(2), 1, i - 1)
'            End If
'
'            tmpIFCd = Mid(tmpData(2), 1, Len(tmpData(2)) - 1) & "0"     '검사코드(시약버전부분은 '0'으로 편집)
'            tmpRst = Trim(tmpData(3))
'            If Left$(tmpRst, 1) = "." Then
'                tmpRst = "0" & tmpRst
'            End If
'            tmpUnit = Trim(tmpData(4))
'            tmpFlag = Trim(tmpData(6))
'            '단위가 'COI'인 경우 결과 편집
'            If tmpUnit = "COI" Then
'                i = InStr(tmpRst, "^")
'                If i <> 0 Then
'                    tmpRef = Mid(tmpRst, 1, i - 1)
'                    tmpRst = Mid(tmpRst, i + 1)
'                End If
'            End If
'
'            '결과정보 구조체에 저장
'            With pResultInfo
'                .ID = pSampleInfo.ID
'                .SEQNO = pSampleInfo.SEQNO
'                .RACK = pSampleInfo.RACK
'                .POS = pSampleInfo.POS
'
'                '결과값 누적
'                .RSTCNT = .RSTCNT + 1
'                .IFCD = .IFCD & tmpIFCd & Chr(124)
'                .RST1 = .RST1 & tmpRst & Chr(124)
'                .RST2 = .RST2 & tmpRef & Chr(124)
'                .UNIT = .UNIT & tmpUnit & Chr(124)
'                .FLAG = .FLAG & tmpFlag & Chr(124)
'            End With
'
'        Case "C"        'Comment Record
'
'        Case "L"
'            '결과값 등록/화면 표시 처리...
'            With pResultInfo
'                If .RSTCNT > 0 Then
'                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
'                End If
'            End With
'
'            Call Init_pResultInfo
'
'    End Select
'
'ErrRtn:
'    If Err <> 0 Then
'        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
'    End If
End Sub
Private Sub DataEditResponse_IMPACT()
    On Error GoTo ErrRtn

    Dim RecType As String   'Record Type
    Dim i       As Integer
    Dim tmpData()   As String
    Dim tmpField()  As String
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpRef$, tmpFlag$

    Dim tmpOneRow() As String
    Dim ii%
    Dim tmpUserID$
    
    tmpOneRow() = Split(RcvBuffer, Chr(10))

    For ii = 0 To UBound(tmpOneRow())
        RecType = Mid$(tmpOneRow(ii), 2, 1)
    
        If RecType = "" Then
            Exit For
        End If
        
        Select Case RecType
            Case "H"        'Header Record
                Call Init_pResultInfo
                
            Case "P"        'Patient Record
                tmpData() = Split(tmpOneRow(ii), Chr(124))
                tmpBarCd = Trim(tmpData(2))
                
                pSampleInfo.ID = tmpBarCd
                
            Case "O"
    
            Case "R"        'Result Record
                '--- 결과데이타 편집
                '7R|1|^^^10^Glu|97.000|mmol/L||||F||234^IMPACT||20050224150525|15989
                '0R|2|^^^200^pH|7.360|||||F||234^IMPACT||20050224150525|15989
                '1R|3|^^^201^pCO2|77.000|mmHg||||X||234^IMPACT||20050224150525|15989
                '2C|1|L|107^Outside Panic Range|I
                '3R|4|^^^202^pO2|83.000|mmHg||||X||234^IMPACT||20050224150525|15989
                '4C|1|L|106^Outside Normal Range|I
                '5R|5|^^^210^Hct|23.000|%||||X||234^IMPACT||20050224150525|15989
                '6C|1|L|106^Outside Normal Range|I
                '7R|6|^^^250^Temp|37.000|||||F||234^IMPACT||20050224150525|15989
                '0R|7|^^^260^HCO3|43.500|mmol/L||||X||234^IMPACT||20050224150525|15989
                '1C|1|L|107^Outside Panic Range|I
                '2R|8|^^^261^BEb|16.200|mmol/L||||F||234^IMPACT||20050224150525|15989
                '3R|9|^^^264^%sO2c|96.000|%||||F||234^IMPACT||20050224150525|15989
                '4R|10|^^^30^Na+|140.000|mmol/L||||F||234^IMPACT||20050224150525|15989
                '5R|11|^^^31^K+|4.800|mmol/L||||F||234^IMPACT||20050224150525|15989
                '6R|12|^^^44^Ca++|1.100|mmol/L||||F||234^IMPACT||20050224150525|15989
                '7R|13|^^^9^Lac|1.500|||||F||234^IMPACT||20050224150525|15989
                tmpData() = Split(tmpOneRow(ii), "|")
                
                tmpField() = Split(tmpData(10), "^")
                tmpUserID = Trim(tmpField(0))
    
                tmpField() = Split(tmpData(2), "^")
                tmpIFCd = Trim(tmpField(3))
                
                tmpRst = Trim(tmpData(3))
                
                '결과정보 구조체에 저장
                With pResultInfo
                    .ID = pSampleInfo.ID
                    .SEQNO = ""
                    .RACK = ""
                    .POS = ""
    
                    '결과값 누적
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst & Chr(124)
                    .RST2 = .RST2 & Chr(124)
                    .UNIT = .UNIT & Chr(124)
                    .FLAG = .FLAG & Chr(124)
                End With
    
            Case "C"        'Comment Record
    
            Case "L"
                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                    End If
                End With
    
                Call Init_pResultInfo
    
        End Select

    Next ii

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub PhaseCfg_Protocol_IMPACT()

    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5      'ENQ
                        RcvBuffer = ""
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                    Case Else
                        m_iPhase = 1
                End Select

            Case 2
                Select Case Asc(wkDat)
                    Case 2      'STX

                    Case 4      'EOT        '10     '<LF>
                        Call DataEditResponse_IMPACT

                    Case 5      'ENQ
                        msComm.Output = Chr(6)   'Send ACK

                    Case 10     'LF
                        RcvBuffer = RcvBuffer & wkDat
                        msComm.Output = Chr(6)
                    
                    Case 21     'NAK
                        msComm.Output = Chr(5)   'Send ENQ
                        m_iPhase = 1

                    Case Is < 0
                    
                    Case Else
                        RcvBuffer = RcvBuffer & wkDat

                End Select

        End Select
    Next ix1
    
End Sub

Private Sub PhaseCfg_Protocol_IMPACT_kang()
'
'    Dim wkDat   As String
'    Dim ix1     As Integer
'
'    For ix1 = 1 To Len(wkBuf)
'        wkDat = Mid$(wkBuf, ix1, 1)
'
'        Select Case Phase
'            'ENQ 대기 상태
'            Case 1
'                Select Case Asc(wkDat)
'                    'ENQ
'                    Case 5
'                        sRcvState = ""
'                        sSndState = ""
'                        RcvBuffer = ""
'
'                        'ACK 전송
'                        Comm1.Output = Chr(6)
'
'                        If giTestMode = 2 Then
'                        Print #2, "<ACK>";
'                        End If
'
'                        Phase = 2
'                    Case Else
'                        sRcvState = ""
'                        sSndState = ""
'                        Phase = 1
'                End Select
'
'            'Packet 모음, Packet 분석(Edit_Data)
'            Case 2
'                Select Case Asc(wkDat)
'                    'STX
'                    Case 2
'                    'EOT
'                    Case 4
'                        Call Edit_Data
'                    'ENQ
'                    Case 5
'                        'ACK 전송
'                        Comm1.Output = Chr(6)
'
'                        If giTestMode = 2 Then
'                        Print #2, "<ACK>";
'                        End If
'                    'LF
'                    Case 10
'                        RcvBuffer = RcvBuffer & wkDat
'                        'ACK 전송
'                        Comm1.Output = Chr(6)
'                        If giTestMode = 2 Then
'                        Print #2, "<ACK>";
'                        End If
'                    'NAK
'                    Case 21
'                        'ENQ 전송
'                        Comm1.Output = Chr(5)
'                        If giTestMode = 2 Then
'                        Print #2, "<ENQ_NAK_P2>";
'                        End If
'                    Case Is < 0
'
'                    Case Else
'                        RcvBuffer = RcvBuffer & wkDat
'                End Select
'
'                    'ENQ
'                    Case 5
'                        'ACK 전송
'                        Comm1.Output = Chr(6)
'                        If giTestMode = 2 Then
'                        Print #2, "<ACK>";
'                        End If
'
'                        RcvBuffer = ""
'                        Phase = 2
'                End Select
'        End Select
'    Next
'
'    Exit Sub
'
'ErrHandler:
'    ViewMsg "PhaseCfg_Protocol 오류 - (" & Err.Description & ")"
End Sub


Private Sub Edit_Data()
'    On Error GoTo ErrHandler
'
''<---- COBAS 장비에서 주로 사용 S --->
'    Dim sBC         As String
'    Dim sLC         As String
'    Dim iBCpos      As Integer
'    Dim iLCpos      As Integer
'
'    Dim iErrCode    As Integer
'    Dim sGeneralErrCode As String
''<---- COBAS 장비에서 주로 사용 E --->
'
'    Dim sJDate      As String
'    Dim sJGbn       As String
'    Dim sJNo        As String
'    Dim sIFSpcCd    As String   '인터페이스시 검체코드
'    Dim sIFRstCd    As String   '인터페이스시 검사항목코드
'    Dim sRxData     As String
'
'    Dim sSampNo     As String
'    Dim sRack       As String
'    Dim sPos        As String
'
'    Dim sRst        As String
'    Dim sRst2       As String
'
'    Dim sTestCd     As String
'    Dim sTestNm     As String
'
'    Dim sBarCd      As String
'    Dim i           As Integer
'
'    Dim sTmp        As String
'    Dim sDat        As String
'    Dim iPos        As Integer
'    Dim iETBpos     As Integer
'    Dim sRecType    As String
'    Dim sBuf        As String
'    Dim sRstGbn     As String
'
'    Dim sUserID     As String
'    Dim sMachine    As String
'
'    Dim sTmpData()  As String
'    Dim sTmpInfo()  As String
'
'    '### Rack Or Tray 방식과 Conflict 방지
'    Call ProtectConflict("Y")
'
'    sRxData = ""
'    sRxData = RcvBuffer
'
'   'sRecType 초기화
'    sRecType = "S"
'
'    Do While sRecType <> ""
'        sTmp = GetByOneUserSymbol(sRxData, sRxData, vbLf)
'        sTmp = GetByOneUserSymbol(sTmp, sTmp, vbCr)
'
'        sRecType = Mid(sTmp, 2, 1)
'
'        If sRecType = "" Then
'           Exit Do
'        End If
'
'        If sRecType = "H" Then
'            '1H|@^\|||IMPACT^Automation Lab^Blood Gas Lab|||||HOST||P|1|20050224151755
'            miRstCnt = 0
'            msTotIFRstCd = ""
'            msTotRst = ""
'            msTotRst2 = ""
'            msBarCd = ""
'            msLocation = ""
'            msRcvDTTM = ""
'            msOperID = ""
'
'        ElseIf sRecType = "P" Then
'            '2P|1|1234|0000000000000006||^^|||U
'            sTmpData = Split(sTmp, "|")
'
'            'PATID(8), BARCODE(12)
'            msBarCd = Trim(sTmpData(2))
'            If Len(msBarCd) = 12 Then
'                msBarCd = Mid(msBarCd, 2, 11)
'            End If
'
'        ElseIf sRecType = "O" Then
'            '3O|1||14|||20050224151755|||||||||Arterial|^^^||||||20050224151755|||F
'            sTmpData = Split(sTmp, "|")
'
'            msRcvDTTM = Trim(sTmpData(7))
'            msOperID = Trim(sTmpData(10))
'
'            'QC 결과 전송시 Exit
'
'        ElseIf sRecType = "R" Then
'            '7R|1|^^^10^Glu|97.000|mmol/L||||F||234^IMPACT||20050224150525|15989
'            '0R|2|^^^200^pH|7.360|||||F||234^IMPACT||20050224150525|15989
'            '1R|3|^^^201^pCO2|77.000|mmHg||||X||234^IMPACT||20050224150525|15989
'            '2C|1|L|107^Outside Panic Range|I
'            '3R|4|^^^202^pO2|83.000|mmHg||||X||234^IMPACT||20050224150525|15989
'            '4C|1|L|106^Outside Normal Range|I
'            '5R|5|^^^210^Hct|23.000|%||||X||234^IMPACT||20050224150525|15989
'            '6C|1|L|106^Outside Normal Range|I
'            '7R|6|^^^250^Temp|37.000|||||F||234^IMPACT||20050224150525|15989
'            '0R|7|^^^260^HCO3|43.500|mmol/L||||X||234^IMPACT||20050224150525|15989
'            '1C|1|L|107^Outside Panic Range|I
'            '2R|8|^^^261^BEb|16.200|mmol/L||||F||234^IMPACT||20050224150525|15989
'            '3R|9|^^^264^%sO2c|96.000|%||||F||234^IMPACT||20050224150525|15989
'            '4R|10|^^^30^Na+|140.000|mmol/L||||F||234^IMPACT||20050224150525|15989
'            '5R|11|^^^31^K+|4.800|mmol/L||||F||234^IMPACT||20050224150525|15989
'            '6R|12|^^^44^Ca++|1.100|mmol/L||||F||234^IMPACT||20050224150525|15989
'            '7R|13|^^^9^Lac|1.500|||||F||234^IMPACT||20050224150525|15989
'            sRcvState = "R"
'            sTmpData = Split(sTmp & "|", "|")
'
'            sBuf = Trim(sTmpData(10)) & "^^"
'            sTmpInfo = Split(sBuf, "^")
'            sUserID = Trim(sTmpInfo(0))
'
'            sMachine = Trim(sTmpData(13))
'            sMachine = sMachine & Space(7)
'
'            Select Case Mid(sMachine, 7, 1)
'                Case "1"
'                    sMachine = "GEMPCS"
'                Case "2"
'                    sMachine = "GEMC4"
'                Case Else
'                    sMachine = "IMPACT"
'            End Select
'
'            sIFRstCd = Trim(sTmpData(2)) & "^^^^"
'            sRst = Trim(sTmpData(3))
'            sTmpData = Split(sIFRstCd, "^")
'            sIFRstCd = Trim(sTmpData(3))
'
'            sRst2 = ""
'            sRst = ConvertResult1("", "", sRst, sIFRstCd)
'            sRst = JudgeResult1(sIFRstCd, sRst, sRst2)
'
'            miRstCnt = miRstCnt + 1
'            msTotIFRstCd = msTotIFRstCd & sIFRstCd & "|"
'            msTotRst = msTotRst & sRst & "|"
'            msTotRst2 = msTotRst2 & sRst2 & "|"
'
'            Select Case sIFRstCd
'                Case "44" 'iCa(mmol/L)
'                    '--- iCa(mg/dL) 추가
'                    If IsNumeric(sRst) Then
'                        sRst = Trim(CStr(Val(sRst) * 4#))
'                    Else
'                        sRst = ""
'                    End If
'
'                    sRst2 = ""
'                    sRst = ConvertResult1("", "", sRst, "1001")
'                    Call JudgeResult1("1001", sRst, sRst2)
'
'                    miRstCnt = miRstCnt + 1
'                    msTotIFRstCd = msTotIFRstCd & "1001" & "|"
'                    msTotRst = msTotRst & sRst & "|"
'                    msTotRst2 = msTotRst2 & sRst2 & "|"
'                    '--- iMg 추가
'                    miRstCnt = miRstCnt + 1
'                    msTotIFRstCd = msTotIFRstCd & "1002" & "|"
'                    msTotRst = msTotRst & "" & "|"
'                    msTotRst2 = msTotRst2 & "" & "|"
'                Case Else
'            End Select
'
'        ElseIf sRecType = "C" Then
'
'        ElseIf sRecType = "L" Then
'            'USERID apply
'            miRstCnt = miRstCnt + 1
'            msTotIFRstCd = msTotIFRstCd & "USER" & "|"
'            msTotRst = msTotRst & sUserID & "|"
'            msTotRst2 = msTotRst2 & "*" & "|"
'
'            'MACHINE apply
'            miRstCnt = miRstCnt + 1
'            msTotIFRstCd = msTotIFRstCd & "MACH" & "|"
'            msTotRst = msTotRst & sMachine & "|"
'            msTotRst2 = msTotRst2 & "*" & "|"
'
'            'IFSEQ로 SORTING (화면IFSEQ순서로 보이기 및 결과등록로직의 용이성을위해)
'            Call Rst_Sorting(miRstCnt, msTotIFRstCd, msTotRst, msTotRst2)
'
'            Call DisplayResultOK(3, Format(Now, "YYYYMMDD"), "", _
'                                        "", "", msBarCd, "", "", "", "", "", "", "", "", _
'                                        miRstCnt, msTotIFRstCd, msTotRst, msTotRst2, _
'                                        "", "")
'
'        Else
'        End If
'    Loop
'
'    If sRcvState = "R" Then
'        If (sSndState = "E") Or (sSndState = "H") Or (sSndState = "P") _
'                Or (sSndState = "O") Or (sSndState = "L") Then
'            'ENQ 전송
'            Comm1.Output = Chr(5)
'            If giTestMode = 2 Then
'            Print #2, "<ENQ_R2>";
'            End If
'
'            Phase = 3
'        Else
'            Phase = 1
'        End If
'    End If
'
'    sRcvState = ""
'    Call ProtectConflict("N")
'
'    Exit Sub
'
'ErrHandler:
'    sRcvState = ""
'    Call ProtectConflict("N")
'    ViewMsg "Edit_Data 에러 발생" & "(" & CStr(Err.Number) & " : " & Err.Description & ")"
End Sub



Private Sub Get_OrderString()

    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .ORDCNT = 0
        End With
        Exit Sub
    End If
    
    ReDim tmpData(m_p_iOrdCnt) As String
    tmpData() = Split(m_p_sTIFCd, Chr(124))
    
    With pSampleInfo
        .ID = m_p_sID
        .SEQNO = m_p_sSeq
        .RACK = m_p_sRack
        .POS = m_p_sPos
        .ORDCNT = m_p_iOrdCnt
        
        ReDim .IFCD(.ORDCNT)
        iCnt = 0
        For ii = 1 To .ORDCNT
            If Trim(tmpData(ii - 1)) <> "" Then
                iCnt = iCnt + 1
                .IFCD(iCnt) = tmpData(ii - 1)
            End If
        Next ii
        .ORDCNT = iCnt      '실제 검사 가능한 항목 갯수
    End With
        
End Sub


'
'   결과정보 구조체 초기화
'
Private Sub Init_pResultInfo()
    
    With pResultInfo
        .ID = ""
        .SEQNO = ""
        .RACK = ""
        .POS = ""
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
    End With
    
End Sub
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=msComm,msComm,-1,RTSEnable
Public Property Get RTSEnable() As Boolean
Attribute RTSEnable.VB_Description = "전송 요청 줄이 가능한지의 여부를 결정합니다."
    RTSEnable = msComm.RTSEnable
End Property

Public Property Let RTSEnable(ByVal New_RTSEnable As Boolean)
    msComm.RTSEnable() = New_RTSEnable
    PropertyChanged "RTSEnable"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=msComm,msComm,-1,RThreshold
Public Property Get RThreshold() As Integer
Attribute RThreshold.VB_Description = "수신할 문자의 수를 반환하거나 설정합니다."
    RThreshold = msComm.RThreshold
End Property

Public Property Let RThreshold(ByVal New_RThreshold As Integer)
    msComm.RThreshold() = New_RThreshold
    PropertyChanged "RThreshold"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=msComm,msComm,-1,Settings
Public Property Get Settings() As String
Attribute Settings.VB_Description = "전송 속도, 패리티, 데이터 비트, 중단 비트 매개 변수를 반환하거나 설정합니다."
    Settings = msComm.Settings
End Property

Public Property Let Settings(ByVal New_Settings As String)
    msComm.Settings() = New_Settings
    PropertyChanged "Settings"
End Property




Private Sub cmdTest_Click()

    wkBuf = Text1
    Call PhaseCfg_Protocol

End Sub

Private Sub msComm_OnComm()
        
    Select Case msComm.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                                ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
            wkBuf = msComm.Input
            
            If sTestMode = "77" Then
                RaiseEvent PrintRcvLog(wkBuf)
            End If
                                
            If iSpaceCnt = 30 Then
                iSpaceCnt = 0
            End If
            iSpaceCnt = iSpaceCnt + 2
            
            RaiseEvent DispMsg(Space(iSpaceCnt) & "장비와 Interface 작업 중...")
            
            Call PhaseCfg_Protocol
            
        Case MSCOMM_EV_CTS      'j
        Case MSCOMM_EV_DSR      ' Change in the DSR line.
        Case MSCOMM_EV_CD       ' Change in the CD line.
        Case MSCOMM_EV_RING     ' Change in the Ring Indicator.
        ' Errors
        Case MSCOMM_ER_BREAK    ' A Break was received.
        ' Code to handle a BREAK goes here, and so on.
        Case MSCOMM_ER_CTSTO    ' CTS Timeout.
        Case MSCOMM_ER_DSRTO    ' DSR Timeout.
        Case MSCOMM_ER_FRAME    ' Framing Error.
        Case MSCOMM_ER_OVERRUN  ' Data Lost.
        Case MSCOMM_ER_CDTO     ' CD (RLSD) Timeout.
        Case MSCOMM_ER_RXOVER   ' Receive buffer overflow.
        Case MSCOMM_ER_RXPARITY ' Parity Error.
        Case MSCOMM_ER_TXFULL   ' Transmit buffer full.
    End Select
    
End Sub
'저장소에서 속성값을 로드합니다.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    msComm.CommPort = PropBag.ReadProperty("CommPort", 1)
    msComm.RTSEnable = PropBag.ReadProperty("RTSEnable", False)
    msComm.RThreshold = PropBag.ReadProperty("RThreshold", 0)
    msComm.Settings = PropBag.ReadProperty("Settings", "9600,n,8,1")
    m_PortOpen = PropBag.ReadProperty("PortOpen", m_def_PortOpen)
    m_OpenPW = PropBag.ReadProperty("OpenPW", m_def_OpenPW)
    m_EditPW = PropBag.ReadProperty("EditPW", m_def_EditPW)
    m_EqName = PropBag.ReadProperty("EqName", m_def_EqName)
    m_bUseBarcode = PropBag.ReadProperty("bUseBarcode", m_def_bUseBarcode)
    m_iPhase = PropBag.ReadProperty("iPhase", m_def_iPhase)
    m_iSendPhase = PropBag.ReadProperty("iSendPhase", m_def_iSendPhase)
    m_sTestMode = PropBag.ReadProperty("sTestMode", m_def_sTestMode)
    m_iFrameN = PropBag.ReadProperty("iFrameN", m_def_iFrameN)
    m_p_sID = PropBag.ReadProperty("p_sID", m_def_p_sID)
    m_p_sSeq = PropBag.ReadProperty("p_sSeq", m_def_p_sSeq)
    m_p_sRack = PropBag.ReadProperty("p_sRack", m_def_p_sRack)
    m_p_sPos = PropBag.ReadProperty("p_sPos", m_def_p_sPos)
    m_p_iOrdCnt = PropBag.ReadProperty("p_iOrdCnt", m_def_p_iOrdCnt)
    m_p_sTIFCd = PropBag.ReadProperty("p_sTIFCd", m_def_p_sTIFCd)
End Sub

'속성값을 저장소에 기록합니다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("CommPort", msComm.CommPort, 1)
    Call PropBag.WriteProperty("RTSEnable", msComm.RTSEnable, False)
    Call PropBag.WriteProperty("RThreshold", msComm.RThreshold, 0)
    Call PropBag.WriteProperty("Settings", msComm.Settings, "9600,n,8,1")
    Call PropBag.WriteProperty("PortOpen", m_PortOpen, m_def_PortOpen)
    Call PropBag.WriteProperty("OpenPW", m_OpenPW, m_def_OpenPW)
    Call PropBag.WriteProperty("EditPW", m_EditPW, m_def_EditPW)
    Call PropBag.WriteProperty("EqName", m_EqName, m_def_EqName)
    Call PropBag.WriteProperty("bUseBarcode", m_bUseBarcode, m_def_bUseBarcode)
    Call PropBag.WriteProperty("iPhase", m_iPhase, m_def_iPhase)
    Call PropBag.WriteProperty("iSendPhase", m_iSendPhase, m_def_iSendPhase)
    Call PropBag.WriteProperty("sTestMode", m_sTestMode, m_def_sTestMode)
    Call PropBag.WriteProperty("iFrameN", m_iFrameN, m_def_iFrameN)
    Call PropBag.WriteProperty("p_sID", m_p_sID, m_def_p_sID)
    Call PropBag.WriteProperty("p_sSeq", m_p_sSeq, m_def_p_sSeq)
    Call PropBag.WriteProperty("p_sRack", m_p_sRack, m_def_p_sRack)
    Call PropBag.WriteProperty("p_sPos", m_p_sPos, m_def_p_sPos)
    Call PropBag.WriteProperty("p_iOrdCnt", m_p_iOrdCnt, m_def_p_iOrdCnt)
    Call PropBag.WriteProperty("p_sTIFCd", m_p_sTIFCd, m_def_p_sTIFCd)
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0,0,0,0
Public Property Get PortOpen() As Boolean
    PortOpen = m_PortOpen
End Property

Public Property Let PortOpen(ByVal New_PortOpen As Boolean)
    m_PortOpen = New_PortOpen
    PropertyChanged "PortOpen"
    
    '--- PortOpen시 암호 확인
    If m_OpenPW <> pOpenPW Then
        MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
        Exit Property
    End If
    '-----------------------
    
    On Error GoTo ErrPortOpen
    If m_PortOpen = True Then
        msComm.PortOpen = True
    End If
    On Error GoTo 0
ErrPortOpen:
    If Err <> 0 Then
        MsgBox "PortOpen Error!!! " & Err.Description, vbCritical
        RaiseEvent DispMsg(Err.Description)
    End If
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get OpenPW() As String
    OpenPW = m_OpenPW
End Property

Public Property Let OpenPW(ByVal New_OpenPW As String)
    m_OpenPW = New_OpenPW
    PropertyChanged "OpenPW"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get EditPW() As String
    EditPW = m_EditPW
End Property

Public Property Let EditPW(ByVal New_EditPW As String)
    m_EditPW = New_EditPW
    PropertyChanged "EditPW"
End Property

'사용자 정의 컨트롤에 대한 속성을 초기화합니다.
Private Sub UserControl_InitProperties()
    m_PortOpen = m_def_PortOpen
    m_OpenPW = m_def_OpenPW
    m_EditPW = m_def_EditPW
    m_EqName = m_def_EqName
    m_bUseBarcode = m_def_bUseBarcode
    m_iPhase = m_def_iPhase
    m_iSendPhase = m_def_iSendPhase
    m_sTestMode = m_def_sTestMode
    m_iFrameN = m_def_iFrameN
    m_p_sID = m_def_p_sID
    m_p_sSeq = m_def_p_sSeq
    m_p_sRack = m_def_p_sRack
    m_p_sPos = m_def_p_sPos
    m_p_iOrdCnt = m_def_p_iOrdCnt
    m_p_sTIFCd = m_def_p_sTIFCd
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get EqName() As String
    EqName = m_EqName
End Property

Public Property Let EqName(ByVal New_EqName As String)
    m_EqName = New_EqName
    PropertyChanged "EqName"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0,0,0,0
Public Property Get bUseBarcode() As Boolean
    bUseBarcode = m_bUseBarcode
End Property

Public Property Let bUseBarcode(ByVal New_bUseBarcode As Boolean)
    m_bUseBarcode = New_bUseBarcode
    PropertyChanged "bUseBarcode"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iPhase() As Integer
    iPhase = m_iPhase
End Property

Public Property Let iPhase(ByVal New_iPhase As Integer)
    m_iPhase = New_iPhase
    PropertyChanged "iPhase"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iSendPhase() As Integer
    iSendPhase = m_iSendPhase
End Property

Public Property Let iSendPhase(ByVal New_iSendPhase As Integer)
    m_iSendPhase = New_iSendPhase
    PropertyChanged "iSendPhase"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get sTestMode() As String
    sTestMode = m_sTestMode
End Property

Public Property Let sTestMode(ByVal New_sTestMode As String)
    m_sTestMode = New_sTestMode
    PropertyChanged "sTestMode"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iFrameN() As Integer
    iFrameN = m_iFrameN
End Property

Public Property Let iFrameN(ByVal New_iFrameN As Integer)
    m_iFrameN = New_iFrameN
    PropertyChanged "iFrameN"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sID() As String
    p_sID = m_p_sID
End Property

Public Property Let p_sID(ByVal New_p_sID As String)
    m_p_sID = New_p_sID
    PropertyChanged "p_sID"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sSeq() As String
    p_sSeq = m_p_sSeq
End Property

Public Property Let p_sSeq(ByVal New_p_sSeq As String)
    m_p_sSeq = New_p_sSeq
    PropertyChanged "p_sSeq"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sRack() As String
    p_sRack = m_p_sRack
End Property

Public Property Let p_sRack(ByVal New_p_sRack As String)
    m_p_sRack = New_p_sRack
    PropertyChanged "p_sRack"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sPos() As String
    p_sPos = m_p_sPos
End Property

Public Property Let p_sPos(ByVal New_p_sPos As String)
    m_p_sPos = New_p_sPos
    PropertyChanged "p_sPos"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get p_iOrdCnt() As Integer
    p_iOrdCnt = m_p_iOrdCnt
End Property

Public Property Let p_iOrdCnt(ByVal New_p_iOrdCnt As Integer)
    m_p_iOrdCnt = New_p_iOrdCnt
    PropertyChanged "p_iOrdCnt"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sTIFCd() As String
    p_sTIFCd = m_p_sTIFCd
End Property

Public Property Let p_sTIFCd(ByVal New_p_sTIFCd As String)
    m_p_sTIFCd = New_p_sTIFCd
    PropertyChanged "p_sTIFCd"
End Property
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function Send_Chr(iChr%) As Variant
    On Error GoTo ErrComm
    msComm.Output = Chr(iChr)
    On Error GoTo 0
ErrComm:
    If Err <> 0 Then
        RaiseEvent DispMsg("Send_Chr 에러 - " & Err.Description)
    End If
End Function

