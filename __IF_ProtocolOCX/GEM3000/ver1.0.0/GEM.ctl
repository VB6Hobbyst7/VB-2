VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl GEM 
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   LockControls    =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   3330
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   840
      Top             =   2235
   End
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
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   300
      Top             =   2205
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "GEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_Port = 0
Const m_def_IPAddress = 0
Const m_def_p_sCmt1 = ""
Const m_def_p_sSpcCd = 0
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
Const m_def_OpenPW = "0"
Const m_def_EditPW = "0"
'속성 변수:
Dim m_Port As Variant
Dim m_IPAddress As Variant
Dim m_p_sCmt1 As String
Dim m_p_sSpcCd As Variant
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
Dim m_OpenPW As String
Dim m_EditPW As String
'이벤트 선언:
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTRstDT$, sTAlarmCd$, sKind$, sSpcDesc$, sOperID$, sTInstID$, sTInstNm$, sOther1$)
'Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTRstDT$, sTAlarmCd$, sKind$, sSpcDesc$, sTInstID$, sTInstNm$, sOther1$)
Event RequestCurOrder(sID$, sRack$, sPos$, sKind$)
Event SendOrderOK(sID$, sSeqNo$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event DispMsg(sMsg$)

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

'For E-170/Hitachi7600
Dim bEndChk As Boolean
Dim bSTXChk As Boolean
Dim sNextSend   As String
Dim RstEnd      As String


Private Function ConvertDataAlarmCode(ByVal sEqNm As String, ByVal Scode As String) As String
    
    Dim sTmp    As String
    
    ConvertDataAlarmCode = "": sTmp = ""
    
    Select Case UCase(sEqNm)
        Case "HITACHI7600"
            Select Case Trim(Scode)
                Case "0": sTmp = ""
                Case "1": sTmp = "ADC?"
                Case "2": sTmp = "Cell?"
                Case "3": sTmp = "Sampl"
                Case "4": sTmp = "Reagn"
                Case "5": sTmp = "ABS?"
                Case "6": sTmp = "Prozon"
                Case "7": sTmp = "Limt0"
                Case "8": sTmp = "Limt1"
                Case "9": sTmp = "Limt2"
                Case "10": sTmp = "Lin."
                Case "11": sTmp = "Lin8."
                Case "12": sTmp = "S1Abs?"
                Case "13": sTmp = "Dup"
                Case "14": sTmp = "Std?"
                Case "15": sTmp = "Sens"
                Case "16": sTmp = "Calib"
                Case "17": sTmp = "SDI"
                Case "18": sTmp = "Noise"
                Case "19": sTmp = "Level"
                Case "20": sTmp = "Slope?"
                Case "21": sTmp = "Margin"
                Case "22": sTmp = "I.Std"
                Case "23": sTmp = "R.Over"
                Case "24": sTmp = "Cmp.T"
                Case "25": sTmp = "Cmp.TI"
                Case "26": sTmp = "LIMTH"
                Case "27": sTmp = "LIMTL"
                Case "28": sTmp = "Random"
                Case "29": sTmp = "Systm1"
                Case "30": sTmp = "Systm2"
                Case "31": sTmp = "Systm3"
                Case "32": sTmp = "Systm4"
                Case "33": sTmp = "Systm5"
                Case "34": sTmp = "Systm6"
                Case "35": sTmp = "QCErr1"
                Case "36": sTmp = "QCErr2"
                Case "37": sTmp = "Calc?"
                Case "38": sTmp = "Over"
                Case "39": sTmp = "???"
                Case "42": sTmp = "Edited"
                Case "44": sTmp = "ReptH"
                Case "45": sTmp = "ReptL"
                Case "51": sTmp = "Resp1"
                Case "52": sTmp = "Resp2"
                Case "53": sTmp = "Condi"
            End Select
        
        Case Else
        
    End Select
    
    ConvertDataAlarmCode = Trim(sTmp)
    
End Function
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
        Case "GEM3000"
            Call PhaseCfg_Protocol_GEM
        
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_DPE_Batch()
'    On Error GoTo ErrRtn
'
'    Dim wkDat   As String
'    Dim ix1 As Integer
'    Dim i   As Integer
'
'    For ix1 = 1 To Len(wkBuf)
'        wkDat = Mid$(wkBuf, ix1, 1)
'
'        Select Case m_iPhase
'            Case 1
'                Select Case Asc(wkDat)
'                    Case 5      'ENQ
'                        m_iPhase = 2
'                        RstEnd = "Y"
'                        bEndChk = True: bSTXChk = False
'
'                        msComm.Output = Chr(6)
'
'                    Case Else
'                        m_iPhase = 1
'                End Select
'
'            Case 2
'                Select Case Asc(wkDat)
'                    Case 2      'STX
'                        If bEndChk = True Then
'                            RcvBuffer = ""
'                        Else
'                            bSTXChk = True
'                        End If
'                        bEndChk = True
'
'                    Case 10     '<LF>
'                        If bEndChk = True Then
'                            Call DataEditResponse_DPE
'                            RcvBuffer = ""
'                        End If
'                        msComm.Output = Chr(6)
'
'                    Case 13     'CR
'                        If bEndChk = True Then
'                            Call DataEditResponse_DPE
'                            RcvBuffer = ""
'                        End If
'
'                    Case 4      'EOT
'                        If sState = "Q" Then
'                            msComm.Output = Chr(5)
'                            m_iSendPhase = 1
'                            sState = ""
'                        End If
''                        m_iPhase = 3
'                        m_iPhase = 1
'
'                    Case 5      'ENQ
'                        bEndChk = True: bSTXChk = True
'                        msComm.Output = Chr(6)   'Send ACK
'
'                    Case 21     'NAK
'                        Call DataEditResponse_DPE
'
'                        m_iSendPhase = 1
'                        m_iFrameN = 1
'
'                        msComm.Output = Chr(5)   'Send ENQ
'
'                    Case 23     ' ETB
'                        bEndChk = False
'
'                    Case Else
'                        If bEndChk = True Then
'                            If bSTXChk = True Then
'                                bSTXChk = False
'                            Else
'                                RcvBuffer = RcvBuffer & wkDat
'                            End If
'                        End If
'
'                End Select
'
'            Case 3
'                Select Case Asc(wkDat)
'                    Case 6      'ACK
'                        Call SendOrder_DPE_Batch
'
'                    Case 5      'ENQ
'                        bEndChk = True: bSTXChk = False
'                        msComm.Output = Chr(6)
'                        m_iPhase = 2
'
'                    Case 21     'NAK
'                        m_iSendPhase = 1
'                        m_iFrameN = 1
'                        msComm.Output = Chr(5)
'                        m_iPhase = 3
'
'                    Case 4      'EOT
'                        m_iPhase = 1
'
'                End Select
'
''            Case 4
''                Select Case Asc(wkDat)
''                    Case 4      'EOT
''                        msComm.Output = Chr(5)
''                        m_iPhase = 3
''                        RcvBuffer = ""
''
''                    Case 5      'ENQ
''                        msComm.Output = Chr(6)
''                        m_iPhase = 2
''
''                    Case 10
''                        msComm.Output = Chr(6)
''                End Select
'
'        End Select
'    Next ix1
'
'ErrRtn:
'    If Err <> 0 Then
'        RaiseEvent DispMsg(Err.Description)
'    End If
End Sub
Private Sub PhaseCfg_Protocol_GEM()
    On Error GoTo ErrRtn
    
    Dim wkDat   As String
    Dim ix1 As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5      'ENQ
                        m_iPhase = 2
                        RstEnd = "Y"
                        bEndChk = True: bSTXChk = False

'                        msComm.Output = Chr(6)
                        Call SendSckData(Chr(6))

                    Case Else
                        m_iPhase = 1
                End Select

            Case 2
                Select Case Asc(wkDat)
                    Case 2      'STX
                        If bEndChk = True Then
                            RcvBuffer = ""
                        Else
                            bSTXChk = True
                        End If
                        bEndChk = True

                    Case 10     '<LF>
                        If bEndChk = True Then
                            Call DataEditResponse_GEM
                            RcvBuffer = ""
                        End If
'                        msComm.Output = Chr(6)
                        Call SendSckData(Chr(6))

                    Case 13     'CR
                        If bEndChk = True Then
                            Call DataEditResponse_GEM
                            RcvBuffer = ""
                        End If

                    Case 4      'EOT
                        If sState = "Q" Then
'                            msComm.Output = Chr(5)
                            Call SendSckData(Chr(5))
                            m_iSendPhase = 1
                        End If
                        m_iPhase = 3

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = True
'                        msComm.Output = Chr(6)   'Send ACK
                        Call SendSckData(Chr(6))

                    Case 21     'NAK
                        Call DataEditResponse_GEM

                        m_iSendPhase = 1
                        m_iFrameN = 1

'                        msComm.Output = Chr(5)   'Send ENQ
                        Call SendSckData(Chr(5))

                    Case 23     ' ETB
                        bEndChk = False

                    Case Else
                        If bEndChk = True Then
                            If bSTXChk = True Then
                                bSTXChk = False
                            Else
                                RcvBuffer = RcvBuffer & wkDat
                            End If
                        End If

                End Select

            Case 3
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        If sState = "Q" Then
'                            Call SendOrder_DPE
                        End If

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
'                        msComm.Output = Chr(6)
                        Call SendSckData(Chr(6))
                        m_iPhase = 2

                    Case 21     'NAK
                        m_iSendPhase = 1
                        m_iFrameN = 1
'                        msComm.Output = Chr(5)
                        Call SendSckData(Chr(5))
                        m_iPhase = 3

                    Case 4      'EOT
                        m_iPhase = 1

                End Select

'            Case 4
'                Select Case Asc(wkDat)
'                    Case 4      'EOT
'                        msComm.Output = Chr(5)
'                        m_iPhase = 3
'                        RcvBuffer = ""
'
'                    Case 5      'ENQ
'                        msComm.Output = Chr(6)
'                        m_iPhase = 2
'
'                    Case 10
'                        msComm.Output = Chr(6)
'                End Select

        End Select
    Next ix1

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg(Err.Description)
    End If
End Sub
Private Sub SendSckData(ByVal sData As String)
    On Error GoTo ErrSck
    
    tcpClient.SendData sData
            
ErrSck:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendSckData - " & Err.Description & "(State:" & tcpClient.State & ")")
        
        Select Case tcpClient.State
            Case 8, 9
                Timer1.Enabled = True
            Case Else
        End Select
    End If
End Sub
' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_GEM()
    On Error GoTo ErrRtn

    Dim RecType As String   'Record Type
    Dim ii      As Integer
    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    Dim tmpKind     As String
    Dim tmpSampType As String
    Dim tmpField()  As String
    Dim tmpData()   As String
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$, tmpAlarmCd$, tmpInstID$, tmpInstNm$
    Dim tmpRstDT$, tmpSpcDesc$, tmpOperID$


    ii = InStr(1, RcvBuffer, "|")
    If ii <> 0 Then
        RecType = Mid$(RcvBuffer, ii - 1, 1)
    Else
        Exit Sub
    End If

    Select Case RecType
        Case "H"        'Header Record
            Call Init_pResultInfo

            tmpField() = Split(RcvBuffer & "||||", "|")
            tmpData() = Split(tmpField(4) & "^^^", "^")
            
            tmpInstID = Trim(tmpData(2))
            tmpInstNm = Trim(tmpData(3))

            pSampleInfo.INSTNM = tmpInstNm
            pSampleInfo.INSTID = tmpInstID

        Case "M"
        Case "P"        'Patient Record
            tmpField() = Split(RcvBuffer, Chr(124))
            If UBound(tmpField()) >= 3 Then
                tmpBarCd = Trim(tmpField(3))
            End If
            pSampleInfo.ID = tmpBarCd

        Case "O"
            tmpSeqNo = "": tmpBarCd = "": tmpRack = "": tmpPos = ""
            tmpField() = Split(RcvBuffer & String(15, "|"), "|")
            
'            tmpBarCd = Trim(tmpField(2))
            tmpSeqNo = Trim(tmpField(3))
            
            tmpSpcDesc = Trim(tmpField(15))
            If tmpSpcDesc = "I" Then
                tmpKind = "QC"
            Else
                tmpKind = ""
            End If
            
            With pSampleInfo
'                .ID = UCase(tmpBarCd)
                .SEQNO = tmpSeqNo
                .KIND = tmpKind
                .SPCCD = tmpSpcDesc
            End With
            
        Case "R"        'Result Record
            '--- 결과데이타 편집
            '2:TEST ID
            '3:RESULT
            '4:UNITS
            '5:Reference Ranges
            '6:Result Abnormal Flags
            '8:Result Status(F:First,C:Rerun)
            '10:Operation ID
            tmpField() = Split(RcvBuffer & String(12, "|"), "|")

            tmpData() = Split(tmpField(2), "^")
            tmpIFCd = Trim(tmpData(3))
            
            tmpRst = Trim(tmpField(3))
            tmpUnit = Trim(tmpField(4))
            tmpFlag = Trim(tmpField(6))
            
            If Trim(tmpField(1)) = "1" Then     'Record Sequence Number
                tmpOperID = Trim(tmpField(10))
                tmpRstDT = Trim(tmpField(12))
            End If
            
            '결과정보 구조체에 저장
            With pResultInfo
                .ID = pSampleInfo.ID
                .SEQNO = pSampleInfo.SEQNO
                .RACK = pSampleInfo.RACK
                .POS = pSampleInfo.POS
                .KIND = pSampleInfo.KIND
                
                .SPCCD = pSampleInfo.SPCCD
                .INSTID = pSampleInfo.INSTID
                .INSTNM = pSampleInfo.INSTNM
                .OPERID = tmpOperID
                                
                '결과값 누적
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
                .RSTDT = .RSTDT & tmpRstDT & Chr(124)
            End With

        Case "C"        'Comment Record
            tmpField() = Split(RcvBuffer & String(3, "|"), Chr(124))
            pResultInfo.OTHER = Trim(tmpField(3))

        Case "L"
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .RSTDT, .ALARMCD, .KIND, .SPCCD, .OPERID, .INSTID, .INSTNM, .OTHER)
                End If
            End With

            Call Init_pResultInfo

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
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
        .INSTNM = ""
        .ALARMCD = ""
        .RSTDT = ""
        .OTHER = ""
        .SPCCD = ""
        .OPERID = ""
    End With
    
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

Private Sub cmdTest_Click()

    wkBuf = Text1
    Call PhaseCfg_Protocol

End Sub

Private Sub tcpClient_Connect()
    Timer1.Enabled = False
    RaiseEvent DispMsg("Connect Server...")
End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo ErrSck
    
    If Timer1.Enabled = True Then
        Timer1.Enabled = False
    End If
    
    tcpClient.GetData wkBuf

    If m_sTestMode = "77" Then
        RaiseEvent PrintRcvLog(wkBuf)
    End If
                        
    If iSpaceCnt = 30 Then
        iSpaceCnt = 0
    End If
    iSpaceCnt = iSpaceCnt + 2
    
    RaiseEvent DispMsg(Space(iSpaceCnt) & "장비와 Interface 작업 중...")
    
    Call PhaseCfg_Protocol
    
ErrSck:
    If Err <> 0 Then
        RaiseEvent DispMsg(Err.Description)
    End If
End Sub


Private Sub tcpClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If tcpClient.State = 9 Then
        RaiseEvent DispMsg("tcpClient_Error (" & Number & ") " & Description)
        
        tcpClient.Close
        Timer1.Enabled = True
    End If
End Sub

Private Sub Timer1_Timer()
        
    RaiseEvent DispMsg("Socket 연결 재시도...")
    
    tcpClient.Close
    
    Call ConnectWinSock(1)
    
End Sub

'저장소에서 속성값을 로드합니다.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

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
    m_p_sSpcCd = PropBag.ReadProperty("p_sSpcCd", m_def_p_sSpcCd)
    m_p_sCmt1 = PropBag.ReadProperty("p_sCmt1", m_def_p_sCmt1)
    m_Port = PropBag.ReadProperty("Port", m_def_Port)
    m_IPAddress = PropBag.ReadProperty("IPAddress", m_def_IPAddress)
End Sub

'속성값을 저장소에 기록합니다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

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
    Call PropBag.WriteProperty("p_sSpcCd", m_p_sSpcCd, m_def_p_sSpcCd)
    Call PropBag.WriteProperty("p_sCmt1", m_p_sCmt1, m_def_p_sCmt1)
    Call PropBag.WriteProperty("Port", m_Port, m_def_Port)
    Call PropBag.WriteProperty("IPAddress", m_IPAddress, m_def_IPAddress)
End Sub

'Public Property Let PortOpen(ByVal New_PortOpen As Boolean)
'    m_PortOpen = New_PortOpen
'    PropertyChanged "PortOpen"
'
'    '--- PortOpen시 암호 확인
'    If m_OpenPW <> pOpenPW Then
'        MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
'        Exit Property
'    End If
'    '-----------------------
'
'    '변수 초기화(E-170/H-7600)
'    RstEnd = "Y": bEndChk = True: bSTXChk = False
'
'
'    On Error GoTo ErrPortOpen
'    If m_PortOpen = True Then
'        msComm.PortOpen = True
'    End If
'    On Error GoTo 0
'ErrPortOpen:
'    If Err <> 0 Then
'        MsgBox "PortOpen Error!!! " & Err.Description, vbCritical
'        RaiseEvent DispMsg(Err.Description)
'    End If
'End Property

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
    m_p_sSpcCd = m_def_p_sSpcCd
    m_p_sCmt1 = m_def_p_sCmt1
    m_Port = m_def_Port
    m_IPAddress = m_def_IPAddress
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
'    msComm.Output = Chr(iChr)
    Call SendSckData(Chr(iChr))
    On Error GoTo 0
ErrComm:
    If Err <> 0 Then
        RaiseEvent DispMsg("Send_Chr 에러 - " & Err.Description)
    End If
End Function

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get p_sSpcCd() As Variant
    p_sSpcCd = m_p_sSpcCd
End Property

Public Property Let p_sSpcCd(ByVal New_p_sSpcCd As Variant)
    m_p_sSpcCd = New_p_sSpcCd
    PropertyChanged "p_sSpcCd"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,
Public Property Get p_sCmt1() As String
    p_sCmt1 = m_p_sCmt1
End Property

Public Property Let p_sCmt1(ByVal New_p_sCmt1 As String)
    m_p_sCmt1 = New_p_sCmt1
    PropertyChanged "p_sCmt1"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get Port() As Variant
    Port = m_Port
End Property

Public Property Let Port(ByVal New_Port As Variant)
    m_Port = New_Port
    PropertyChanged "Port"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get IPAddress() As Variant
    IPAddress = m_IPAddress
End Property

Public Property Let IPAddress(ByVal New_IPAddress As Variant)
    m_IPAddress = New_IPAddress
    PropertyChanged "IPAddress"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function Connect() As Variant

End Function

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function ConnectWinSock(Optional ByVal iGbn As Integer) As Variant
    On Error GoTo ErrRtn

    If iGbn = 0 Then
        '가장처음 Connect시 암호 확인
        If m_OpenPW <> pOpenPW Then
            MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
            Exit Function
        End If
        
        '변수 초기화(E-170/H-7600)
        RstEnd = "Y": bEndChk = True: bSTXChk = False
    End If

    tcpClient.RemotePort = Val(m_Port)
    
    tcpClient.RemoteHost = m_IPAddress
    tcpClient.Connect tcpClient.RemoteHost, tcpClient.RemotePort
    
    Timer1.Enabled = True
    
    Call Sleep(500)
    
    RaiseEvent DispMsg("WinSock State: " & tcpClient.State)
        
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("ConnectWinSock Err - " & Err.Description)
    End If
End Function

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function CloseWinSock() As Variant
    On Error GoTo ErrClose
    
    tcpClient.Close
    
ErrClose:
    If Err <> 0 Then
        RaiseEvent DispMsg("CloseWinSock Err - " & Err.Description)
    End If
End Function

