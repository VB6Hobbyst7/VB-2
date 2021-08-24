VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl COBAS8000 
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2415
   LockControls    =   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   2415
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   195
      Top             =   2130
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "TEST"
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   0
      Top             =   165
      Width           =   1425
   End
End
Attribute VB_Name = "COBAS8000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'기본 속성 값:
Const m_def_pPatBirth = ""
Const m_def_pPatSex = ""
Const m_def_pPatName = ""
Const m_def_EqName = 0
Const m_def_iPhase = 0
Const m_def_iSendPhase = 0
Const m_def_p_sID = ""
Const m_def_p_sSeq = ""
Const m_def_p_sRack = ""
Const m_def_p_sPos = ""
Const m_def_p_iOrdCnt = 0
Const m_def_p_sTIFCd = ""
Const m_def_SocketPort = "0"
Const m_def_pType = "HL7"
Const m_def_pCmt1 = ""
Const m_def_pSpcCd = 0
Const m_def_pTSVol = 0
Const m_def_pRerunGbn = ""
Const m_def_pSIndex = 0
Const m_def_pUseBarCd = 0
Const m_def_pTestMode = "0"
Const m_def_pOpenPW = ""
Const m_def_pEditPW = ""
'속성 변수:
Dim m_pPatBirth As String
Dim m_pPatSex As String
Dim m_pPatName As String
Dim m_EqName As Variant
Dim m_iPhase As Integer
Dim m_iSendPhase As Integer
Dim m_p_sID As String
Dim m_p_sSeq As String
Dim m_p_sRack As String
Dim m_p_sPos As String
Dim m_p_iOrdCnt As Integer
Dim m_p_sTIFCd As String
Dim m_SocketPort As String
Dim m_pType As String
Dim m_pCmt1 As String
Dim m_pSpcCd As Variant
Dim m_pTSVol As Variant
Dim m_pRerunGbn As String
Dim m_pSIndex As Boolean
Dim m_pUseBarCd As Boolean
Dim m_pTestMode As String
Dim m_pOpenPW As String
Dim m_pEditPW As String
'이벤트 선언:
Event DispMsg(sMsg$)
Event DispMsgComm(sMsg$)
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sInstID$, sAlarmCd$, sKind$, sTRstDT$, sOther1$)
Event RequestCurOrder(sID$, sRack$, sPos$, sKind$)
Event SendOrderOK(sID$, sSeqNo$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event ClientOpen()
Event ClientClose()

'===== User Define
'인터페이스에서 사용
Dim RcvBuffer   As String
Dim wkBuf   As String
Dim msState As String

'구조체 지정
Private pSampleInfo As SAMPLE_INFO
Private pResultInfo As RESULT_INFO

'기타
Dim sOpenPW$, sEditPW$
Dim iSpaceCnt   As Integer

'for HL7
Dim msMsgType   As String
Dim miSckMax    As Integer

Dim msRackType  As String
Dim msCupType   As String
Dim msStatGbn   As String

Dim mlCtlID As Long
Dim msCtlID As String


Private Sub PhaseCfg_Protocol()

    '--- 사용자 확인
    If m_pEditPW <> pEditPW Then
        MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
        Exit Sub
    End If
    '---------------
    
    If m_EqName = "0" Or m_EqName = "" Then
        RaiseEvent DispMsg("검사장비명을 지정해 주십시오.!!!")
        Exit Sub
    End If
    
    Select Case UCase(m_EqName)
        Case "COBAS8000"
            If m_pType = "ASTM" Then
            Else
                Call PhaseCfg_Protocol_Cobas8000_HL7
            End If
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub

Private Sub PhaseCfg_Protocol_Cobas8000_HL7()
    On Error GoTo ErrPhase
    
    Dim wkDat   As String
    Dim ix1     As Integer
    Dim ix2     As Integer
    Dim tmpOneMsg() As String
    
    tmpOneMsg = Split(wkBuf, Chr(28))
    
    For ix1 = 0 To UBound(tmpOneMsg) - 1
        If Trim(tmpOneMsg(ix1)) = "" Then Exit For
        
        RcvBuffer = ""
        
        For ix2 = 1 To Len(tmpOneMsg(ix1))
            wkDat = Mid(tmpOneMsg(ix1), ix2, 1)
    
            Select Case Asc(wkDat)
                Case 11     'VT
'                    RcvBuffer = wkDat
                    
                Case 10     'LF
                                    
                Case 13     'CR
                   Call DataEdit_Cobas8000
                   RcvBuffer = ""
                    
                Case Else
                    RcvBuffer = RcvBuffer & wkDat
                    
            End Select
        Next ix2
        
        '결과값 등록/화면 표시 처리...
        With pResultInfo
            If .RSTCNT > 0 Then
                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, .ALARMCD, .KIND, .RSTDT, .OTHER)
            End If
        End With

        Call Init_pResultInfo
        
        If msState = "Q" Then
            Call SendOrder_Cobas8000
        Else
            Call Send_ACK
        End If
    Next ix1

ErrPhase:
    If Err <> 0 Then
        RaiseEvent DispMsg(Err.Description)
    End If
End Sub

Private Sub DataEdit_Cobas8000()
    On Error GoTo ErrRtn

    Dim sHeader     As String   'Record Type
    Dim sMsgType    As String
    Dim ii          As Integer
    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    Dim tmpKind     As String
    Dim tmpSampType As String
    Dim tmpContType As String
    
    Dim aField()    As String
    Dim aData()     As String
    
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$, tmpAlarmCd$, tmpInstID$
    Dim tmpRstDT$, tmpCmt$, tmpSrcCd$
    
    aField = Split(RcvBuffer, Chr(124))
    
    sHeader = Trim(aField(0))
    
    Select Case sHeader
        Case "MSH"          'Message heading
            'MSH|^~\&|cobas 8000||host||20090402173655||OUL^R22|13007||2.5||||ER||UNICODE UTF-8|
            msMsgType = Trim(aField(8))
            msCtlID = Trim(aField(9))
            
            Select Case msMsgType
                Case "OUL^R22^PCUPL", "OUL^R22^ICUPL", "OUL^R22^ECUPL"
                    pSampleInfo.KIND = "CAL"
                Case "OUL^R22^REAL", "OUL^R22^BATCH"
                    pSampleInfo.KIND = "QC"
                Case Else
                    pSampleInfo.KIND = ""
            End Select
                    
        Case "MSA"          'Message Acknowledgment
            'MSA|AE|38764|ORA-20001: Validation error|
            'MSA|AA|38764||
            If msMsgType = "ACK" Then
                If Trim(aField(1)) = "AA" Then
                    RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
                Else
                    RaiseEvent SendOrderOK("", "", "", "")
                End If
            End If
        
        Case "PID"          'Patient Identification Segment - PID
            Call Init_pResultInfo
        
        Case "SPM"          'Specimen Segment - SPM for patient and quality control results
            'SPM||110005||S1||not|||||P|||^^^^|||20100429161525||||||||||SC|[CR]
            tmpBarCd = Trim(aField(2))
            tmpSeqNo = ""
            tmpKind = Trim(aField(11))
            
            pResultInfo.ID = tmpBarCd
            pResultInfo.SEQNO = tmpSeqNo
            If tmpKind = "Q" Then
                pResultInfo.KIND = "QC"
            Else
                pResultInfo.KIND = pSampleInfo.KIND
            End If
            
        Case "SAC"          'Specimen Container Detail Segment - SAC
            'SAC||||||||||50042|2|
            tmpRack = Trim(aField(10))
            tmpPos = Trim(aField(11))
            
            pResultInfo.RACK = tmpRack
            pResultInfo.POS = tmpPos
                
        Case "OBR"          'Observation Request Segment - OBR
            'OBR|1|||989|[CR]
            
        Case "TQ1"          'Timing Quantity Segment - TQ1
            
        Case "OBX"          'Observation Result Segment - OBX
            'OBX|1||989||1.1|mmol/L|^TECH\^NORM\^CRIT\^USER|N|||F|||20091218164600|bmserv^SYSTEM||28|ISE^2^MU1#ISE#1#2^4|20100430102029|[CR]
            tmpIFCd = Trim(aField(3))
            tmpRst = Trim(aField(5))
            tmpUnit = Trim(aField(6))
            tmpFlag = Trim(aField(8))

            If Left$(tmpRst, 1) = "." Then
                tmpRst = "0" & tmpRst
            End If
            
            tmpRstDT = Trim(aField(19))
            
            '결과정보 구조체에 저장
            With pResultInfo
                '결과값 누적
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
                .RSTDT = .RSTDT & tmpRstDT & Chr(124)
            End With
        
        Case "TCD"          'Test Code Detail Segment - TCD
            'TCD|989|Dec|[CR]
                    
        Case "NTE"          'Comment Segment - NTE
            'NTE|1|I|23^ISE·Sample·range·over|I|[CR]
            tmpSrcCd = Trim(aField(2))
            aData() = Split(aField(3), "^")
            If tmpSrcCd = "I" Then
                If Trim(aData(0)) = "0" Then
                Else
                    tmpAlarmCd = Trim(aData(1))
                End If
            ElseIf tmpSrcCd = "L" Then
                tmpAlarmCd = Trim(aData(0))
            End If
            pResultInfo.ALARMCD = pResultInfo.ALARMCD & tmpAlarmCd & Chr(124)
        
        Case "QPD"          'Query Parameter Segment - QPD (for a test selection inquiry)
            'QPD|TSREQ|12896|000137||50042|2||||S1|SC|R1|R|
            'QPD|TSREQ|12897|**********************||50045|1||||S1|SC|R1|R|
            sMsgType = Trim(aField(1))
            tmpBarCd = Trim(aField(3))
            tmpSeqNo = ""
            tmpRack = Trim(aField(5))
            tmpPos = Trim(aField(6))
            tmpSampType = Trim(aField(10))  'rack type
            tmpContType = Trim(aField(11))  'container type
            tmpKind = Trim(aField(12))      '1st/rerun
            msStatGbn = Trim(aField(13))    'sample priority
            
            If sMsgType = "TSREQ" And tmpBarCd <> "" Then
                msState = "Q"
                pSampleInfo.ID = tmpBarCd
            Else
                msState = ""
                pSampleInfo.ID = ""
            End If
            
            pSampleInfo.SEQNO = tmpSeqNo
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos
            If tmpKind = "R2" Then
                pSampleInfo.KIND = tmpKind
            End If
            pSampleInfo.SPCCD = tmpSampType
            pSampleInfo.CONTAINER = tmpContType
        
        Case "RCP"          'Response Control Parameter Segment - RCP

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit_Cobas8000 - " & Err.Description)
    End If
End Sub

Private Sub EqOutput_Socket(ByVal sMsg As String)
    sckServer(miSckMax).SendData (sMsg)
    
    If pTestMode = "77" Then
        RaiseEvent PrintSendLog(sMsg)
    End If
End Sub

Private Sub SendOrder_Cobas8000()
    On Error GoTo ErrSendOrder
    
    Dim sSndBuf As String: sSndBuf = ""
    Dim iCnt As Integer
    Dim aDilInfo()  As String
    Dim sDilInfo    As String
    
    '----- 검사항목 조회
    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS, pSampleInfo.KIND)
    
    Call Get_OrderString
    
    mlCtlID = mlCtlID + 1   'control id 증가
    
    'MSH                                                                                                                          'ER
'    sSndBuf = sSndBuf & "MSH|^~\&|host||cobas 8000||" & Format(Now(), "YYYYMMDDhhmmss") & "||OML^O33|" & Trim(mlCtlID) & "||2.5||||NE||UNICODE UTF-8|" & Chr(13)
    sSndBuf = sSndBuf & "MSH|^~\&|host||cobas 8000||" & Format(Now(), "YYYYMMDDhhmmss") & "||OML^O33|" & Trim(mlCtlID) & "||2.5||||SU||UNICODE UTF-8|" & Chr(13)    '2012/4/26 yk
    
'    'PID(optional)
    sSndBuf = sSndBuf & "PID|1|" & Trim(pSampleInfo.CMT1) & " |||^" & Trim(pSampleInfo.OTHER) & "||" & Trim(pSampleInfo.BIRTH) & "|" & Trim(pSampleInfo.SEX) & "|" & Chr(13)
'    sSndBuf = sSndBuf & "PID|1|" & Trim(pSampleInfo.CMT1) & " |||^" & "" & "||" & Trim(pSampleInfo.BIRTH) & "|" & Trim(pSampleInfo.SEX) & "|" & Chr(13)
    
    'SPM
    sSndBuf = sSndBuf & "SPM||" & pSampleInfo.ID & "||" & pSampleInfo.SPCCD & "||||||||||" & Trim(pSampleInfo.CMT1) & "^^^^|||||||||||||" & Trim(pSampleInfo.CONTAINER) & "|" & Chr(13)
    
    'SAC
    sSndBuf = sSndBuf & "SAC||||||||||" & pSampleInfo.RACK & "|" & pSampleInfo.POS & "|" & Chr(13)
    
    If pSampleInfo.ORDCNT = 0 Then
        sSndBuf = Chr(11) & sSndBuf & Chr(28) & Chr(13)
        EqOutput_Socket (sSndBuf)
    
        msState = ""
        Exit Sub
    End If
    
    For iCnt = 1 To pSampleInfo.ORDCNT
        'TQ1
        sSndBuf = sSndBuf & "TQ1|1||||||||" & IIf(msStatGbn = "", "R", msStatGbn) & "|" & Chr(13)
        'OBR
        If InStr(pSampleInfo.IFCD(iCnt), "^") = 0 Then
            '일반항목
            'OBR|1|||990^|||||||A[CR]
            sSndBuf = sSndBuf & "OBR|" & Trim(iCnt) & "|||" & Trim(pSampleInfo.IFCD(iCnt)) & "^|||||||A|" & Chr(13)
        Else
            'Dilution 정보 추가
            sDilInfo = ""
            
            Erase aDilInfo()
            aDilInfo() = Split(Trim(pSampleInfo.IFCD(iCnt)), "^")
            
            Select Case UCase(Trim(aDilInfo(1)))
                Case "INC"
                    sDilInfo = "Inc"
                Case "DEC"
                    sDilInfo = "Dec"
                Case "1", "2", "3", "5", "10", "20", "50", "100", "400"
                    sDilInfo = Trim(aDilInfo(1))
                Case Else
            End Select
            'OBR|4|||8717^Inc|||||||A[CR]
            sSndBuf = sSndBuf & "OBR|" & Trim(iCnt) & "|||" & Trim(aDilInfo(0)) & "^" & sDilInfo & "|||||||A|" & Chr(13)
        End If
    Next iCnt
   
    'SendData 편집
    sSndBuf = Chr(11) & sSndBuf & Chr(28) & Chr(13)
    
'    '<S--- UTF-8 Encode...2012/4/27 yk
'    sSndBuf = EncodeUTF8(sSndBuf)
''    Dim UTF8Bytes() As Byte
''    Call EncodeUTF8_Byte(sSndBuf, UTF8Bytes)
''
''    sckServer(miSckMax).SendData UTF8Bytes
    
    Dim UTF8Bytes() As Byte
    UTF8Bytes = EncodeUTF8_ADOStream(sSndBuf)
    
    sckServer(miSckMax).SendData UTF8Bytes
'    '>E----------------
    
'    EqOutput_Socket (sSndBuf)
    
    msState = ""
   
ErrSendOrder:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder_Cobas8000 - " & Err.Description)
    End If
End Sub

Private Sub Send_ACK()
    
    Dim sSndBuf As String
    
    mlCtlID = mlCtlID + 1
    
    'MSH
    sSndBuf = "MSH|^~\&|host||cobas 8000||" & Format(Now(), "YYYYMMDDhhmmss") & "||ACK|" & Trim(mlCtlID) & "||2.5||||ER||UNICODE UTF-8" & Chr(13)
'    sSndBuf = "MSH|^~\&|host||cobas 8000||" & Format(Now(), "YYYYMMDDhhmmss") & "||ACK|" & Trim(mlCtlID) & "||2.5||||SU||UNICODE UTF-8" & Chr(13)       '2012/4/26 yk
    'MSA
    sSndBuf = sSndBuf & "MSA|AA|" & msCtlID & "||" & Chr(13)

    
    sSndBuf = Chr(11) & sSndBuf & Chr(28) & Chr(13)
    
    EqOutput_Socket (sSndBuf)

End Sub

Private Sub Get_OrderString()
        
    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .SEQNO = m_p_sSeq
            .RACK = m_p_sRack
            .POS = m_p_sPos
            .ORDCNT = 0
            .SINDEX = False
            .CMT1 = ""
            .OTHER = ""         '환자명
            .BIRTH = ""         '생년월일
            .SEX = ""           '성별
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
        .SINDEX = m_pSIndex
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
        
        .CMT1 = m_pCmt1
        .OTHER = m_pPatName     '환자명
        .BIRTH = m_pPatBirth    '생년월일
        .SEX = m_pPatSex        '성별
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
        .QCGBN = ""
        .KIND = ""
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
        .INSTID = ""
        .ALARMCD = ""
        .RSTDT = ""
        .OTHER = ""
    End With
    
    pSampleInfo.CMT1 = ""
    
End Sub

Private Function EncodeUTF8_ADOStream(strText As String) As Byte()
    On Error GoTo ErrEncode
    
    Dim oStream As New ADODB.Stream
    Dim Data()  As Byte
    
    With oStream
        .Charset = "UTF-8"
        .Mode = adModeReadWrite
        .Type = adTypeText
        .Open
        
        .WriteText strText
        .Flush
        
        .Position = 0
        .Type = adTypeBinary
        .Read 3
        Data = .Read()
        
        .Close: Set oStream = Nothing
    End With
    
    EncodeUTF8_ADOStream = Data

ErrEncode:
    If Err <> 0 Then
        RaiseEvent DispMsg("EncodeUTF8_ADOStream - " & Err.Description)
    End If
End Function
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get EqName() As Variant
    EqName = m_EqName
End Property

Public Property Let EqName(ByVal New_EqName As Variant)
    m_EqName = New_EqName
    PropertyChanged "EqName"
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
'MemberInfo=13,0,0,
Public Property Get p_sID() As String
    p_sID = m_p_sID
End Property

Public Property Let p_sID(ByVal New_p_sID As String)
    m_p_sID = New_p_sID
    PropertyChanged "p_sID"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,
Public Property Get p_sSeq() As String
    p_sSeq = m_p_sSeq
End Property

Public Property Let p_sSeq(ByVal New_p_sSeq As String)
    m_p_sSeq = New_p_sSeq
    PropertyChanged "p_sSeq"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,
Public Property Get p_sRack() As String
    p_sRack = m_p_sRack
End Property

Public Property Let p_sRack(ByVal New_p_sRack As String)
    m_p_sRack = New_p_sRack
    PropertyChanged "p_sRack"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,
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
'MemberInfo=13,0,0,
Public Property Get p_sTIFCd() As String
    p_sTIFCd = m_p_sTIFCd
End Property

Public Property Let p_sTIFCd(ByVal New_p_sTIFCd As String)
    m_p_sTIFCd = New_p_sTIFCd
    PropertyChanged "p_sTIFCd"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get SocketPort() As String
    SocketPort = m_SocketPort
End Property

Public Property Let SocketPort(ByVal New_SocketPort As String)
    m_SocketPort = New_SocketPort
    PropertyChanged "SocketPort"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,HL7
Public Property Get pType() As String
    pType = m_pType
End Property

Public Property Let pType(ByVal New_pType As String)
    m_pType = New_pType
    PropertyChanged "pType"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,
Public Property Get pCmt1() As String
    pCmt1 = m_pCmt1
End Property

Public Property Let pCmt1(ByVal New_pCmt1 As String)
    m_pCmt1 = New_pCmt1
    PropertyChanged "pCmt1"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get pSpcCd() As Variant
    pSpcCd = m_pSpcCd
End Property

Public Property Let pSpcCd(ByVal New_pSpcCd As Variant)
    m_pSpcCd = New_pSpcCd
    PropertyChanged "pSpcCd"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get pTSVol() As Variant
    pTSVol = m_pTSVol
End Property

Public Property Let pTSVol(ByVal New_pTSVol As Variant)
    m_pTSVol = New_pTSVol
    PropertyChanged "pTSVol"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,
Public Property Get pRerunGbn() As String
    pRerunGbn = m_pRerunGbn
End Property

Public Property Let pRerunGbn(ByVal New_pRerunGbn As String)
    m_pRerunGbn = New_pRerunGbn
    PropertyChanged "pRerunGbn"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0,0,0,0
Public Property Get pSIndex() As Boolean
    pSIndex = m_pSIndex
End Property

Public Property Let pSIndex(ByVal New_pSIndex As Boolean)
    m_pSIndex = New_pSIndex
    PropertyChanged "pSIndex"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0,0,0,0
Public Property Get pUseBarCd() As Boolean
    pUseBarCd = m_pUseBarCd
End Property

Public Property Let pUseBarCd(ByVal New_pUseBarCd As Boolean)
    m_pUseBarCd = New_pUseBarCd
    PropertyChanged "pUseBarCd"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get pTestMode() As String
    pTestMode = m_pTestMode
End Property

Public Property Let pTestMode(ByVal New_pTestMode As String)
    m_pTestMode = New_pTestMode
    PropertyChanged "pTestMode"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,
Public Property Get pOpenPW() As String
    pOpenPW = m_pOpenPW
End Property

Public Property Let pOpenPW(ByVal New_pOpenPW As String)
    m_pOpenPW = New_pOpenPW
    PropertyChanged "pOpenPW"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,
Public Property Get pEditPW() As String
    pEditPW = m_pEditPW
End Property

Public Property Let pEditPW(ByVal New_pEditPW As String)
    m_pEditPW = New_pEditPW
    PropertyChanged "pEditPW"
End Property

Private Sub cmdTest_Click()

    wkBuf = Text1
    Call PhaseCfg_Protocol
    
End Sub

Private Sub sckServer_Close(Index As Integer)
    On Error GoTo ErrSckClose

    RaiseEvent ClientClose
        
    sckServer(Index).Close
    Unload sckServer(Index)
    
ErrSckClose:
    If Err <> 0 Then
        RaiseEvent DispMsg("sckServer_Close - " & Err.Description)
    End If
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    If Index = 0 Then
        miSckMax = miSckMax + 1
        
        Load sckServer(miSckMax)
        sckServer(miSckMax).LocalPort = 0
        sckServer(miSckMax).Accept requestID
        
        RaiseEvent ClientOpen
    End If
    
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    sckServer(Index).GetData wkBuf
    
    If pTestMode = "77" Then
        RaiseEvent PrintRcvLog(wkBuf)
    End If
    
    If iSpaceCnt = 30 Then
        iSpaceCnt = 0
    End If
    iSpaceCnt = iSpaceCnt + 2
    
    RaiseEvent DispMsgComm(Space(iSpaceCnt) & "장비와 Interface 작업 중...")
    
    Call PhaseCfg_Protocol
    
End Sub

Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    RaiseEvent DispMsg("sckServer_Error(" & Trim(Index) & ") " & Description)
    
End Sub

'사용자 정의 컨트롤에 대한 속성을 초기화합니다.
Private Sub UserControl_InitProperties()
    m_EqName = m_def_EqName
    m_iPhase = m_def_iPhase
    m_iSendPhase = m_def_iSendPhase
    m_p_sID = m_def_p_sID
    m_p_sSeq = m_def_p_sSeq
    m_p_sRack = m_def_p_sRack
    m_p_sPos = m_def_p_sPos
    m_p_iOrdCnt = m_def_p_iOrdCnt
    m_p_sTIFCd = m_def_p_sTIFCd
    m_SocketPort = m_def_SocketPort
    m_pType = m_def_pType
    m_pCmt1 = m_def_pCmt1
    m_pSpcCd = m_def_pSpcCd
    m_pTSVol = m_def_pTSVol
    m_pRerunGbn = m_def_pRerunGbn
    m_pSIndex = m_def_pSIndex
    m_pUseBarCd = m_def_pUseBarCd
    m_pTestMode = m_def_pTestMode
    m_pOpenPW = m_def_pOpenPW
    m_pEditPW = m_def_pEditPW
    m_pPatName = m_def_pPatName
    m_pPatBirth = m_def_pPatBirth
    m_pPatSex = m_def_pPatSex
End Sub

'저장소에서 속성값을 로드합니다.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_EqName = PropBag.ReadProperty("EqName", m_def_EqName)
    m_iPhase = PropBag.ReadProperty("iPhase", m_def_iPhase)
    m_iSendPhase = PropBag.ReadProperty("iSendPhase", m_def_iSendPhase)
    m_p_sID = PropBag.ReadProperty("p_sID", m_def_p_sID)
    m_p_sSeq = PropBag.ReadProperty("p_sSeq", m_def_p_sSeq)
    m_p_sRack = PropBag.ReadProperty("p_sRack", m_def_p_sRack)
    m_p_sPos = PropBag.ReadProperty("p_sPos", m_def_p_sPos)
    m_p_iOrdCnt = PropBag.ReadProperty("p_iOrdCnt", m_def_p_iOrdCnt)
    m_p_sTIFCd = PropBag.ReadProperty("p_sTIFCd", m_def_p_sTIFCd)
    m_SocketPort = PropBag.ReadProperty("SocketPort", m_def_SocketPort)
    m_pType = PropBag.ReadProperty("pType", m_def_pType)
    m_pCmt1 = PropBag.ReadProperty("pCmt1", m_def_pCmt1)
    m_pSpcCd = PropBag.ReadProperty("pSpcCd", m_def_pSpcCd)
    m_pTSVol = PropBag.ReadProperty("pTSVol", m_def_pTSVol)
    m_pRerunGbn = PropBag.ReadProperty("pRerunGbn", m_def_pRerunGbn)
    m_pSIndex = PropBag.ReadProperty("pSIndex", m_def_pSIndex)
    m_pUseBarCd = PropBag.ReadProperty("pUseBarCd", m_def_pUseBarCd)
    m_pTestMode = PropBag.ReadProperty("pTestMode", m_def_pTestMode)
    m_pOpenPW = PropBag.ReadProperty("pOpenPW", m_def_pOpenPW)
    m_pEditPW = PropBag.ReadProperty("pEditPW", m_def_pEditPW)
    m_pPatName = PropBag.ReadProperty("pPatName", m_def_pPatName)
    m_pPatBirth = PropBag.ReadProperty("pPatBirth", m_def_pPatBirth)
    m_pPatSex = PropBag.ReadProperty("pPatSex", m_def_pPatSex)
End Sub

'속성값을 저장소에 기록합니다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("EqName", m_EqName, m_def_EqName)
    Call PropBag.WriteProperty("iPhase", m_iPhase, m_def_iPhase)
    Call PropBag.WriteProperty("iSendPhase", m_iSendPhase, m_def_iSendPhase)
    Call PropBag.WriteProperty("p_sID", m_p_sID, m_def_p_sID)
    Call PropBag.WriteProperty("p_sSeq", m_p_sSeq, m_def_p_sSeq)
    Call PropBag.WriteProperty("p_sRack", m_p_sRack, m_def_p_sRack)
    Call PropBag.WriteProperty("p_sPos", m_p_sPos, m_def_p_sPos)
    Call PropBag.WriteProperty("p_iOrdCnt", m_p_iOrdCnt, m_def_p_iOrdCnt)
    Call PropBag.WriteProperty("p_sTIFCd", m_p_sTIFCd, m_def_p_sTIFCd)
    Call PropBag.WriteProperty("SocketPort", m_SocketPort, m_def_SocketPort)
    Call PropBag.WriteProperty("pType", m_pType, m_def_pType)
    Call PropBag.WriteProperty("pCmt1", m_pCmt1, m_def_pCmt1)
    Call PropBag.WriteProperty("pSpcCd", m_pSpcCd, m_def_pSpcCd)
    Call PropBag.WriteProperty("pTSVol", m_pTSVol, m_def_pTSVol)
    Call PropBag.WriteProperty("pRerunGbn", m_pRerunGbn, m_def_pRerunGbn)
    Call PropBag.WriteProperty("pSIndex", m_pSIndex, m_def_pSIndex)
    Call PropBag.WriteProperty("pUseBarCd", m_pUseBarCd, m_def_pUseBarCd)
    Call PropBag.WriteProperty("pTestMode", m_pTestMode, m_def_pTestMode)
    Call PropBag.WriteProperty("pOpenPW", m_pOpenPW, m_def_pOpenPW)
    Call PropBag.WriteProperty("pEditPW", m_pEditPW, m_def_pEditPW)
    Call PropBag.WriteProperty("pPatName", m_pPatName, m_def_pPatName)
    Call PropBag.WriteProperty("pPatBirth", m_pPatBirth, m_def_pPatBirth)
    Call PropBag.WriteProperty("pPatSex", m_pPatSex, m_def_pPatSex)
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function mSckServerOpen() As Variant
    On Error GoTo ErrSckOpen
    
    miSckMax = 0
    With sckServer(0)
        .LocalPort = Val(m_SocketPort)
        .Listen
    End With
    
ErrSckOpen:
    If Err <> 0 Then
        RaiseEvent RaiseError("mSckServerOpen - " & Err.Description)
    End If
End Function

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,
Public Property Get pPatName() As String
    pPatName = m_pPatName
End Property

Public Property Let pPatName(ByVal New_pPatName As String)
    m_pPatName = New_pPatName
    PropertyChanged "pPatName"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,
Public Property Get pPatBirth() As String
    pPatBirth = m_pPatBirth
End Property

Public Property Let pPatBirth(ByVal New_pPatBirth As String)
    m_pPatBirth = New_pPatBirth
    PropertyChanged "pPatBirth"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,
Public Property Get pPatSex() As String
    pPatSex = m_pPatSex
End Property

Public Property Let pPatSex(ByVal New_pPatSex As String)
    m_pPatSex = New_pPatSex
    PropertyChanged "pPatSex"
End Property

