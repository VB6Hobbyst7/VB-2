VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl CFX96 
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
Attribute VB_Name = "CFX96"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_bNewMode = False
Const m_def_p_sCmt1 = ""
Const m_def_p_sSpcCd = 0
Const m_def_p_sTSVol = "0"
Const m_def_p_sRerunGbn = "0"
Const m_def_p_bSIndex = 0
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
Dim m_bNewMode As Boolean
Dim m_p_sCmt1 As String
Dim m_p_sSpcCd As Variant
Dim m_p_sTSVol As String
Dim m_p_sRerunGbn As String
Dim m_p_bSIndex As Boolean
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
Event DispMsgComm(sMsg$)
Event RequestNextOrder()
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTInstID$, sTAlarmCd$, sKind$, sTRstDT$, sOther1$)
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


'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=msComm,msComm,-1,CommPort
Public Property Get CommPort() As Integer
Attribute CommPort.VB_Description = "통신 포트 번호를 반환하거나 설정합니다."
    CommPort = MSComm.CommPort
End Property

Public Property Let CommPort(ByVal New_CommPort As Integer)
    MSComm.CommPort() = New_CommPort
    PropertyChanged "CommPort"
End Property

Private Function ConvertDataAlarmCode(ByVal sEqNm As String, ByVal sCode As String) As String
    
    Dim sTmp    As String
    
    ConvertDataAlarmCode = "": sTmp = ""
    
    Select Case UCase(sEqNm)
        Case "HITACHI7600"
            Select Case Trim(sCode)
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
        Case "DPE"
            If m_bUseBarcode = True Then
                Call PhaseCfg_Protocol_DPE          '바코드사용
            Else
                Call PhaseCfg_Protocol_DPE_Batch    'Batch Mode
            End If
        
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_DPE_Batch()
    On Error GoTo ErrRtn
    
    Dim wkDat   As String
    Dim ix1 As Integer
    Dim i   As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5      'ENQ
                        m_iPhase = 2
                        RstEnd = "Y"
                        bEndChk = True: bSTXChk = False

                        MSComm.Output = Chr(6)

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
                            Call DataEditResponse_DPE
                            RcvBuffer = ""
                        End If
                        MSComm.Output = Chr(6)

                    Case 13     'CR
                        If bEndChk = True Then
                            Call DataEditResponse_DPE
                            RcvBuffer = ""
                        End If

                    Case 4      'EOT
                        If sState = "Q" Then
                            MSComm.Output = Chr(5)
                            m_iSendPhase = 1
                            sState = ""
                        End If
'                        m_iPhase = 3
                        m_iPhase = 1
                        
                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = True
                        MSComm.Output = Chr(6)   'Send ACK

                    Case 21     'NAK
                        Call DataEditResponse_DPE

                        m_iSendPhase = 1
                        m_iFrameN = 1

                        MSComm.Output = Chr(5)   'Send ENQ

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
                        Call SendOrder_DPE_Batch

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
                        MSComm.Output = Chr(6)
                        m_iPhase = 2

                    Case 21     'NAK
                        m_iSendPhase = 1
                        m_iFrameN = 1
                        MSComm.Output = Chr(5)
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
Private Sub PhaseCfg_Protocol_DPE()
    On Error GoTo ErrRtn
    
    Dim wkDat   As String
    Dim ix1 As Integer
    Dim i   As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5      'ENQ
                        m_iPhase = 2
                        RstEnd = "Y"
                        bEndChk = True: bSTXChk = False

                        MSComm.Output = Chr(6)

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
                            If m_bNewMode = True Then
                                Call DataEditResponse_DPE_NewMode
                            Else
                                Call DataEditResponse_DPE
                            End If
                            RcvBuffer = ""
                        End If
                        MSComm.Output = Chr(6)

                    Case 13     'CR
                        If bEndChk = True Then
                            If m_bNewMode = True Then
                                Call DataEditResponse_DPE_NewMode
                            Else
                                Call DataEditResponse_DPE
                            End If
                            RcvBuffer = ""
                        End If

                    Case 4      'EOT
                        If sState = "Q" Then
                            MSComm.Output = Chr(5)
                            m_iSendPhase = 1
                        End If
                        m_iPhase = 3

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = True
                        MSComm.Output = Chr(6)   'Send ACK

                    Case 21     'NAK
                        If m_bNewMode = True Then
                            Call DataEditResponse_DPE_NewMode
                        Else
                            Call DataEditResponse_DPE
                        End If
                        
                        m_iSendPhase = 1
                        m_iFrameN = 1

                        MSComm.Output = Chr(5)   'Send ENQ

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
                        If sState = "Q" Or m_p_sRerunGbn = "R" Then
                            If m_bNewMode = True Then
                                Call SendOrder_DPE_NewMode
                            Else
                                Call SendOrder_DPE
                            End If
                        End If

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
                        MSComm.Output = Chr(6)
                        m_iPhase = 2

                    Case 21     'NAK
                        m_iSendPhase = 1
                        m_iFrameN = 1
                        MSComm.Output = Chr(5)
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

' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_DPE()
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
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$, tmpAlarmCd$, tmpInstID$
    Dim tmpRstDT$, tmpCmt$, tmpOther$


    ii = InStr(1, RcvBuffer, "|")
    If ii <> 0 Then
        RecType = Mid$(RcvBuffer, ii - 1, 1)
    Else
        Exit Sub
    End If

    Select Case RecType
        Case "H"        'Header Record
        Case "M"        'Calibration Result Record
            'M|1|PCR|      |^^^23|D2|16||2289^2625^2290^2617^^\3577^3999^3563^3976^^\^^^^^\^^^^^\^^^^^\^^^^^
            'M|1|ICR|      |ISE1||||^^^^^^^|^^^^^^^|^^^^^^^|-32.2^-35.8^-28.4^-31.8^59.2^138^140^0|-33.7^-46.1^-24.7^-37.6^58.2^4.9^4.2^0.0|135.3^139.4^130.5^134.0^-50.5^96^102^1
            tmpField() = Split(RcvBuffer, Chr(124))
                        
            tmpBarCd = Trim(tmpField(2))
            tmpKind = "CAL"
                        
            Select Case tmpBarCd
                Case "PCR"
                    tmpInstID = Trim(tmpField(5))
                    tmpAlarmCd = Trim(tmpField(6))
                    If tmpAlarmCd = "0" Then
                        tmpAlarmCd = ""
                    End If
                    
                    tmpData() = Split(tmpField(4), "^")
                    tmpIFCd = Trim(tmpData(3))
                    
                    tmpRst = Replace(Trim(tmpField(8)), vbCr, "")
                    
                    '결과정보 구조체에 저장
                    With pResultInfo
                        .ID = tmpBarCd
                        .SEQNO = "": .RACK = "": .POS = "":
                        .Kind = tmpKind
                        
                        '결과값 누적
                        .RSTCNT = 1
                        .IFCD = tmpIFCd & Chr(124)
                        .RST1 = tmpRst & Chr(124)
                        .RST2 = Chr(124)
                        .UNIT = tmpUnit & Chr(124)
                        .FLAG = tmpFlag & Chr(124)
                        .INSTID = tmpInstID & Chr(124)
                        .RSTDT = Format(Now, "YYYYMMDDHHNNSS") & Chr(124)
                        .ALARMCD = tmpAlarmCd & Chr(124)
                    End With
                    
                Case "ICR"      'ISE(Na,K,Cl 동시에 넘어옴...)
                    tmpInstID = Trim(tmpField(4))
                    
                    '결과정보 구조체에 저장
                    With pResultInfo
                        .ID = tmpBarCd
                        .SEQNO = "": .RACK = "": .POS = "":
                        .Kind = tmpKind
                         
                        '결과값 누적
                        .RSTCNT = 3
                        .IFCD = "989" & Chr(124) & "990" & Chr(124) & "991" & Chr(124)
                        .RST1 = Replace(Trim(tmpField(11)), vbCr, "") & Chr(124) & Replace(Trim(tmpField(12)), vbCr, "") & Chr(124) & Replace(Trim(tmpField(13)), vbCr, "") & Chr(124)
                        .RST2 = String(3, Chr(124))
                        .UNIT = String(3, Chr(124))
                        .FLAG = String(3, Chr(124))
                        .INSTID = tmpInstID & Chr(124) & tmpInstID & Chr(124) & tmpInstID & Chr(124)
                        .RSTDT = Format(Now, "YYYYMMDDHHNNSS") & Chr(124) & Format(Now, "YYYYMMDDHHNNSS") & Chr(124) & Format(Now, "YYYYMMDDHHNNSS") & Chr(124)
                        .ALARMCD = String(3, Chr(124))
                    End With
            End Select
            
        Case "P"        'Patient Record
            Call Init_pResultInfo

        Case "Q"        'Order Request Record
            tmpData() = Split(RcvBuffer, "|")
            sReqStatusCd = Trim(tmpData(12))    'Order Request Status Code
            tmpData() = Split(tmpData(2), "/")

            tmpBarCd = Trim(tmpData(1))
            tmpSeqNo = tmpData(0)
            tmpSampType = tmpData(2)    'SAMPLE TYPE...(2004/10/25 yk)
            tmpRack = tmpData(3)
            tmpPos = tmpData(4)
            tmpKind = tmpData(5)

            tmpData() = Split(tmpSeqNo, "^")
            tmpSeqNo = Trim(tmpData(2))

'            If tmpBarCd <> "" Then    'BarCode ID가 잘 넘어왔는지 검사
'                sState = "Q"
'                pSampleInfo.ID = UCase(tmpBarCd)
'            Else
'                sState = ""
'                pSampleInfo.ID = ""
'            End If
            If tmpBarCd <> "" And sReqStatusCd <> "A" Then      '2004/4/2 yk
                sState = "Q"
                pSampleInfo.ID = UCase(tmpBarCd)
            Else
                sState = ""
                pSampleInfo.ID = ""
            End If

            pSampleInfo.SEQNO = tmpSeqNo
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos
            pSampleInfo.Kind = Trim(tmpKind)
            pSampleInfo.SPCCD = Trim(tmpSampType)   'SAMPLE TYPE

        Case "O"
            tmpSeqNo = "": tmpBarCd = "": tmpRack = "": tmpPos = ""
            tmpField() = Split(RcvBuffer, "|")
            ii = InStr(1, tmpField(2), "^")
            If ii <> 0 Then
                tmpData() = Split(tmpField(2), "^")
                tmpSeqNo = Trim(tmpData(0))
                tmpBarCd = Trim(tmpData(1))
                tmpRack = Trim(tmpData(3))
                tmpPos = Trim(tmpData(4))
            End If
            tmpKind = Trim(tmpField(3))

            tmpRstDT = Trim(tmpField(22))
            
            tmpCmt = ""
            tmpCmt = Trim(tmpField(19))     'Comment...2007/1/9 yk
            If InStr(tmpCmt, "^") > 0 Then
                Erase tmpData()
                tmpData() = Split(tmpCmt, "^")
                tmpCmt = Trim(tmpData(0))
            End If

            pSampleInfo.ID = UCase(tmpBarCd)
            pSampleInfo.SEQNO = tmpSeqNo
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos
            pSampleInfo.Kind = tmpKind
            pSampleInfo.RSTDT = tmpRstDT        '결과일시
            pSampleInfo.CMT1 = tmpCmt           'Comment1

        Case "R"        'Result Record
            '--- 결과데이타 편집
            '2:TEST ID
            '3:RESULT
            '4:UNITS
            '5:Reference Ranges
            '6:Result Abnormal Flags
            '8:Result Status(F:First,C:Rerun)
            tmpData() = Split(RcvBuffer, "|")

            tmpIFCd = Trim(tmpData(2))
            tmpIFCd = Mid(tmpIFCd, 4)
            tmpIFCd = Mid(tmpIFCd, 1, InStr(1, tmpIFCd, "/") - 1)
            tmpRst = Trim(tmpData(3))
            tmpUnit = Trim(tmpData(4))
            tmpFlag = Trim(tmpData(6))
            tmpInstID = Trim(tmpData(13))

            '--- 결과값에 "^" 들어갈 경우 편집
            ii = InStr(1, tmpRst, "^")
            If ii <> 0 Then tmpRst = Mid(tmpRst, ii + 1)

            If Left$(tmpRst, 1) = "." Then
                tmpRst = "0" & tmpRst
            End If

            '결과정보 구조체에 저장
            With pResultInfo
                .ID = pSampleInfo.ID
                .SEQNO = pSampleInfo.SEQNO
                .RACK = pSampleInfo.RACK
                .POS = pSampleInfo.POS
                .Kind = pSampleInfo.Kind
                .OTHER = pSampleInfo.CMT1
                
                '결과값 누적
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
                .INSTID = .INSTID & tmpInstID & Chr(124)        'Inst ID...(2005/1/2) yk
                .RSTDT = .RSTDT & pSampleInfo.RSTDT & Chr(124)  '결과일시(2005/6/10) yk
            End With

        Case "C"        'Comment Record
            'Data Alarm 편집
            tmpData() = Split(RcvBuffer, Chr(124))

'            tmpAlarmCd = ConvertDataAlarmCode(UCase(m_EqName), Trim(tmpData(3)))
            tmpAlarmCd = Trim(tmpData(3))
            If tmpAlarmCd = "0" Then        '2005/4/20 yk
                tmpAlarmCd = ""
            End If
            pResultInfo.ALARMCD = pResultInfo.ALARMCD & tmpAlarmCd & Chr(124)

        Case "L"
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, .ALARMCD, .Kind, .RSTDT, .OTHER)
                End If
            End With

            Call Init_pResultInfo

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub

' *=====================================================*
' *               Data편집 & 응답처리 (New Mode)        *
' *=====================================================*
Private Sub DataEditResponse_DPE_NewMode()
    On Error GoTo ErrRtn

    Dim RecType As String   'Record Type
    Dim ii      As Integer
    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    Dim tmpKind     As String
    Dim tmpSampType As String
    Dim tmpContType As String
    Dim tmpField()  As String
    Dim tmpData()   As String
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$, tmpAlarmCd$, tmpInstID$
    Dim tmpRstDT$, tmpCmt$
    Dim aRow()  As String

    ii = InStr(1, RcvBuffer, "|")
    If ii <> 0 Then
        RecType = Mid$(RcvBuffer, ii - 1, 1)
    Else
        Exit Sub
    End If

    If InStr(1, RcvBuffer, Chr(13)) > 0 Then        '2007/6/22 yk
        aRow() = Split(RcvBuffer, Chr(13))
        RcvBuffer = aRow(0)
    End If

    Select Case RecType
        Case "H"        'Header Record
        Case "M"        'Calibration Result Record
        Case "P"        'Patient Record
            Call Init_pResultInfo

        Case "Q"        'Order Request Record
            'Q|1|^^______________________^1^5032^1^^S1^SC||ALL||||||||O
            tmpField() = Split(RcvBuffer, Chr(124))
            
            tmpData() = Split(tmpField(2), "^")
            tmpBarCd = Trim(tmpData(2))
            tmpSeqNo = Trim(tmpData(3))
            tmpRack = Trim(tmpData(4))
            tmpPos = Trim(tmpData(5))
            tmpSampType = Trim(tmpData(7))      'SAMPLE TYPE
            tmpContType = Trim(tmpData(8))      'Container Type
            If UBound(tmpData()) >= 9 Then
                tmpKind = Trim(tmpData(9))          'R1/R2
            End If
            
            sReqStatusCd = Trim(tmpField(12))    'Order Request Status Code

            If tmpBarCd <> "" And sReqStatusCd <> "A" Then
                sState = "Q"
                pSampleInfo.ID = UCase(tmpBarCd)
            Else
                sState = ""
                pSampleInfo.ID = ""
            End If

            pSampleInfo.SEQNO = tmpSeqNo
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos
            pSampleInfo.Kind = Trim(tmpKind)
            pSampleInfo.SPCCD = Trim(tmpSampType)       'SAMPLE TYPE
            pSampleInfo.CONTAINER = Trim(tmpContType)   'Container Type

        Case "O"
            'O|1|000003   |3^5238^3^^S1^SC|^^^2^1|R||20000529125556||||N||||1|||||||20000529125645|||F
            tmpSeqNo = "": tmpBarCd = "": tmpRack = "": tmpPos = ""
            tmpField() = Split(RcvBuffer, "|")
            
            tmpBarCd = Trim(tmpField(2))
            
            ii = InStr(1, tmpField(3), "^")
            If ii <> 0 Then
                tmpData() = Split(tmpField(3), "^")
                tmpSeqNo = Trim(tmpData(0))
                tmpRack = Trim(tmpData(1))
                tmpPos = Trim(tmpData(2))
                tmpSampType = Trim(tmpData(4))      'Rack Type(S1~5, QC:Control)
                If tmpSampType = "QC" Then
                    tmpKind = "QC"
                Else
                    tmpKind = ""
                End If
            End If
            
            tmpRstDT = Trim(tmpField(22))
            
            pSampleInfo.ID = UCase(tmpBarCd)
            pSampleInfo.SEQNO = tmpSeqNo
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos
            pSampleInfo.Kind = tmpKind
            pSampleInfo.RSTDT = tmpRstDT        '결과일시

        Case "R"        'Result Record
            'R|1|^^^2/1/not|8.60|nmol/L||N||F||BMSERV|||E11
            '--- 결과데이타 편집
            '2:TEST ID
            '3:RESULT
            '4:UNITS
            '5:Reference Ranges
            '6:Result Abnormal Flags
            '8:Result Status(F:First,C:Rerun)
            tmpField() = Split(RcvBuffer, "|")

            tmpData() = Split(tmpField(2), "^")
            tmpIFCd = Trim(tmpData(3))
            Erase tmpData()
            tmpData() = Split(tmpIFCd, "/")
            tmpIFCd = Trim(tmpData(0))
            
            tmpRst = Trim(tmpField(3))
            tmpUnit = Trim(tmpField(4))
            tmpFlag = Trim(tmpField(6))
            If tmpFlag = "N" Then tmpFlag = ""
            tmpInstID = Trim(tmpField(13))

            '--- 결과값에 "^" 들어갈 경우 편집
            ii = InStr(1, tmpRst, "^")
            If ii <> 0 Then tmpRst = Mid(tmpRst, ii + 1)

            If Left$(tmpRst, 1) = "." Then
                tmpRst = "0" & tmpRst
            End If

            '결과정보 구조체에 저장
            With pResultInfo
                .ID = pSampleInfo.ID
                .SEQNO = pSampleInfo.SEQNO
                .RACK = pSampleInfo.RACK
                .POS = pSampleInfo.POS
                .Kind = pSampleInfo.Kind
                .OTHER = pSampleInfo.CMT1
                
                '결과값 누적
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
                .INSTID = .INSTID & tmpInstID & Chr(124)        'Inst ID...(2005/1/2) yk
                .RSTDT = .RSTDT & pSampleInfo.RSTDT & Chr(124)  '결과일시(2005/6/10) yk
            End With

        Case "C"        'Comment Record
            tmpField() = Split(RcvBuffer, Chr(124))
            
            If Trim(tmpField(4)) = "G" Then         'Comment
                tmpCmt = ""
                tmpCmt = Trim(tmpField(3))
                If InStr(tmpCmt, "^") > 0 Then
                    tmpData() = Split(tmpCmt, "^")
                    tmpCmt = Trim(tmpData(0))
                End If
                If Trim(tmpCmt) <> "" Then
                    pSampleInfo.CMT1 = tmpCmt
                End If
            
            ElseIf Trim(tmpField(4)) = "I" Then     'Data Alarm 편집
                tmpData() = Split(RcvBuffer, Chr(124))
    
                tmpAlarmCd = Trim(tmpData(3))
                If tmpAlarmCd = "0" Then
                    tmpAlarmCd = ""
                End If
                pResultInfo.ALARMCD = pResultInfo.ALARMCD & tmpAlarmCd & Chr(124)
            End If
            
        Case "L"
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, .ALARMCD, .Kind, .RSTDT, .OTHER)
                End If
            End With

            Call Init_pResultInfo

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEditResponse_DPE_NewMode 오류발생 - " & Err.Description)
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
        .Kind = ""
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
'
'   환자 Order 전송
'
Private Sub SendOrder_DPE_Batch()
    On Error GoTo Err_Rtn

    Dim sSendBuff   As String
    Dim iCnt    As Integer
    Dim ChkSum  As String
    Dim sStat   As String
    
    Select Case m_iSendPhase
        Case 0
            m_iSendPhase = 1
            MSComm.Output = Chr(5)
            Exit Sub
        
        Case 1
            'Header Record
'            sSendBuff = m_iFrameN & "H|\^&|||HOST^2|||||H7600^1|TSDWN^REPLY|P|1" & vbCr
            sSendBuff = m_iFrameN & "H|\^&|||HOST^2|||||H7600^1|TSDWN^BATCH|P|1" & vbCr
            
            'Patient Record
            sSendBuff = sSendBuff & "P|1" & vbCr
                    
'            '----- 검사항목 조회
'            RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS, pSampleInfo.KIND)
            
            Call Get_OrderString
            
            'Order Record
'            sSendBuff = sSendBuff & "O|1|" & pSampleInfo.SEQNO & "^" & Left(Trim(pSampleInfo.ID) & Space(13), 13) _
'                    & "^" & pSampleInfo.SPCCD & "^" _
'                    & Trim(pSampleInfo.RACK) & "^" & Trim(pSampleInfo.POS) & "|" & pSampleInfo.KIND & "|"
            sSendBuff = sSendBuff & "O|1|" & pSampleInfo.SEQNO & "^" & Left(Trim(pSampleInfo.ID) & Space(13), 13) _
                    & "^1^" _
                    & Trim(pSampleInfo.RACK) & "^" & Trim(pSampleInfo.POS) & "|R1|"
                    
            '검사항목 Order코드 추가
            For iCnt = 1 To pSampleInfo.ORDCNT
                '정상 오더
                Select Case Trim$(pSampleInfo.IFCD(iCnt))
                    Case "989", "990", "991"
                        'ISE 항목이 있는 경우 전체 검사(ISE 검사는 2가지 조합만 오더 가능)
                        If Val(InStr(1, sSendBuff, "989" & "/")) = 0 Then
                            sSendBuff = sSendBuff & "^^^989/\"
                        End If
                        If Val(InStr(1, sSendBuff, "990" & "/")) = 0 Then
                            sSendBuff = sSendBuff & "^^^990/\"
                        End If
                        If Val(InStr(1, sSendBuff, "991" & "/")) = 0 Then
                            sSendBuff = sSendBuff & "^^^991/\"
                        End If
                        
                    Case Else
                        '일반항목
                        sSendBuff = sSendBuff & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "/\"
                End Select
            Next iCnt
            
            If pSampleInfo.SINDEX = True Then           'Serum Index
                sSendBuff = sSendBuff & "^^^992/\^^^993/\^^^994/\"
            End If
            
            If pSampleInfo.ORDCNT > 0 And Trim(sReqStatusCd) <> "A" Then
                sSendBuff = Left(sSendBuff, Len(sSendBuff) - 1)      '"\" Cutting
            End If
            
            'STAT RACK에 대한 처리추가
'            If Left(pSampleInfo.RACK, 1) = "4" Then
'                sStat = "S"
'            Else
                sStat = "R"
'            End If

            If Trim(pSampleInfo.CMT1) = "" Then
                sSendBuff = sSendBuff & "|" & sStat & "||" & Format(Now, "YYYYMMDDHHNNSS") & "||||N||^^||||||" _
                        & "^^^^||||||O" & vbCr
            Else
                'Comment 정보 존재할 경우 Comment1 란에 특정정보 전송(2005/8/1 yk)
                sSendBuff = sSendBuff & "|" & sStat & "||" & Format(Now, "YYYYMMDDHHNNSS") & "||||N||^^||||||" _
                        & Trim(pSampleInfo.CMT1) & "^^^^||||||O" & vbCr
            End If
                    
            'Terminator Record
            sSendBuff = sSendBuff & "L|1|N"


            '--- Text의 내용이 240byte를 넘어갈 경우 처리 추가...
            If Len(sSendBuff) >= 241 Then
                sNextSend = Mid(sSendBuff, 241)
                sSendBuff = Left(sSendBuff, 240)
                sSendBuff = sSendBuff & Chr(23)

                m_iFrameN = m_iFrameN + 1
                m_iSendPhase = 2
            Else
                sSendBuff = sSendBuff & Chr(13) & Chr(3)
                GoTo Send_Terminate
            End If

        Case 2
            sSendBuff = m_iFrameN & sNextSend & Chr(13) & Chr(3)
            sNextSend = ""

Send_Terminate:
            m_iSendPhase = 3

        Case 3      'EOT
            MSComm.Output = Chr(4)   'EOT
            m_iFrameN = 1
            m_iPhase = 3
            m_iSendPhase = 1

            sState = "": sReqStatusCd = ""

            'BarCode Mode가 아닌 경우 다음 오더 조회
            RaiseEvent RequestNextOrder
    
            Exit Sub
    End Select

    ChkSum = ChkSum_ASTM(sSendBuff)
    sSendBuff = sSendBuff & ChkSum
    MSComm.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10)

    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
    End If

'    '전송된 오더가 있는 경우 화면표시
'    If pSampleInfo.ORDCNT > 0 And sReqStatusCd = "O" Then
'        If Trim(sNextSend) = "" And m_iSendPhase <> 2 Then
'            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
'        End If
'    Else
'        '조회된 내용이 없는 경우 환자정보 구조체 초기화
'        Call Init_pResultInfo
'
'        RaiseEvent SendOrderOK("", "", "", "")
'    End If

Err_Rtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub
'
'   환자 Order 전송
'
Private Sub SendOrder_DPE()
    On Error GoTo Err_Rtn

    Dim sSendBuff   As String
    Dim iCnt    As Integer
    Dim ChkSum  As String
    Dim sStat   As String
    
    Select Case m_iSendPhase
        Case 1
            'Header Record
            sSendBuff = m_iFrameN & "H|\^&|||HOST^2|||||H7600^1|TSDWN^REPLY|P|1" & vbCr
                                    
            'Patient Record
            sSendBuff = sSendBuff & "P|1" & vbCr
                    
            '----- 검사항목 조회
            RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS, pSampleInfo.Kind)
            
            Call Get_OrderString
            
            'Order Record
'            sSendBuff = sSendBuff & "O|1|" & pSampleInfo.SEQNO & "^" & Left(Trim(pSampleInfo.ID) & Space(13), 13) & "^1^" _
'                    & Trim(pSampleInfo.RACK) & "^" & Trim(pSampleInfo.POS) & "|" & pSampleInfo.KIND & "|"
            '2004/10/25 yk
            sSendBuff = sSendBuff & "O|1|" & pSampleInfo.SEQNO & "^" & Left(Trim(pSampleInfo.ID) & Space(13), 13) _
                    & "^" & pSampleInfo.SPCCD & "^" _
                    & Trim(pSampleInfo.RACK) & "^" & Trim(pSampleInfo.POS) & "|" & pSampleInfo.Kind & "|"
            
            '검사항목 Order코드 추가
            For iCnt = 1 To pSampleInfo.ORDCNT
                'Request Information Code에 따라 검사항목을 추가하거나 취소한다.
                If Trim(sReqStatusCd) = "O" Then
                    '정상 오더
                    Select Case Trim$(pSampleInfo.IFCD(iCnt))
                        Case "989", "990", "991"
                            'ISE 항목이 있는 경우 전체 검사(ISE 검사는 2가지 조합만 오더 가능)
                            If Val(InStr(1, sSendBuff, "989" & "/")) = 0 Then
                                sSendBuff = sSendBuff & "^^^989/\"
                            End If
                            If Val(InStr(1, sSendBuff, "990" & "/")) = 0 Then
                                sSendBuff = sSendBuff & "^^^990/\"
                            End If
                            If Val(InStr(1, sSendBuff, "991" & "/")) = 0 Then
                                sSendBuff = sSendBuff & "^^^991/\"
                            End If
                            
                        Case Else
                            '일반항목
                            sSendBuff = sSendBuff & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "/\"
                    End Select
                    
                ElseIf Trim(sReqStatusCd) = "A" Then
                    '오더 취소
                    pSampleInfo.SINDEX = False
                End If
            Next iCnt
            
            If pSampleInfo.SINDEX = True Then           'Serum Index
                sSendBuff = sSendBuff & "^^^992/\^^^993/\^^^994/\"
            End If
            
            If pSampleInfo.ORDCNT > 0 And Trim(sReqStatusCd) <> "A" Then
                sSendBuff = Left(sSendBuff, Len(sSendBuff) - 1)      '"\" Cutting
            End If
            
            'STAT RACK에 대한 처리추가
            If Left(pSampleInfo.RACK, 1) = "4" Then
                sStat = "S"
            Else
                sStat = "R"
            End If

            If Trim(pSampleInfo.CMT1) = "" Then
                sSendBuff = sSendBuff & "|" & sStat & "||" & Format(Now, "YYYYMMDDHHNNSS") & "||||N||^^||||||" _
                        & "^^^^||||||O" & vbCr
            Else
                'Comment 정보 존재할 경우 Comment1 란에 특정정보 전송(2005/8/1 yk)
                sSendBuff = sSendBuff & "|" & sStat & "||" & Format(Now, "YYYYMMDDHHNNSS") & "||||N||^^||||||" _
                        & Trim(pSampleInfo.CMT1) & "^^^^||||||O" & vbCr
            End If
                    
            'Terminator Record
            sSendBuff = sSendBuff & "L|1|N"


            '--- Text의 내용이 240byte를 넘어갈 경우 처리 추가...
            If Len(sSendBuff) >= 241 Then
                sNextSend = Mid(sSendBuff, 241)
                sSendBuff = Left(sSendBuff, 240)
                sSendBuff = sSendBuff & Chr(23)

                m_iFrameN = m_iFrameN + 1
                m_iSendPhase = 2
            Else
                sSendBuff = sSendBuff & Chr(13) & Chr(3)
                GoTo Send_Terminate
            End If

        Case 2
            sSendBuff = m_iFrameN & sNextSend & Chr(13) & Chr(3)
            sNextSend = ""

Send_Terminate:
            m_iSendPhase = 3

        Case 3      'EOT
            MSComm.Output = Chr(4)   'EOT
            m_iFrameN = 1
            m_iPhase = 3
            m_iSendPhase = 1

            sState = "": sReqStatusCd = ""

            Exit Sub
    End Select

    ChkSum = ChkSum_ASTM(sSendBuff)
    sSendBuff = sSendBuff & ChkSum
    MSComm.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10)

    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
    End If

    '전송된 오더가 있는 경우 화면표시
    If pSampleInfo.ORDCNT > 0 And sReqStatusCd = "O" Then
        If Trim(sNextSend) = "" And m_iSendPhase <> 2 Then
            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
        End If
    Else
        '조회된 내용이 없는 경우 환자정보 구조체 초기화
        Call Init_pResultInfo

        RaiseEvent SendOrderOK("", "", "", "")
    End If

Err_Rtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub

'
'   환자 Order 전송 (New Mode)
'
Private Sub SendOrder_DPE_NewMode()
    On Error GoTo Err_Rtn

    Dim sSendBuff   As String
    Dim iCnt    As Integer
    Dim ChkSum  As String
    Dim sStat   As String
    Dim sSpcGbn As String
    
    Dim aDilInfo()  As String
    Dim sDilInfo    As String
    
    Dim sTOrder As String
    
    Select Case m_iSendPhase
        Case 1
            'Header Record
            sSendBuff = m_iFrameN & "H|\^&|||HOST^2|||||H7600^1|TSDWN^REPLY|P|1" & vbCr
                                    
            'Patient Record
            sSendBuff = sSendBuff & "P|1" & vbCr
                    
            '----- 검사항목 조회
            RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS, pSampleInfo.Kind)
            
            Call Get_OrderString
            
            'Order Record
            sSendBuff = sSendBuff & "O|1|" & Left(Trim(pSampleInfo.ID) & Space(22), 22) & Chr(124)      'ID 22자리로...
            sSendBuff = sSendBuff & pSampleInfo.SEQNO & "^" & Trim(pSampleInfo.RACK) & "^" & Trim(pSampleInfo.POS) & "^^" _
                        & pSampleInfo.SPCCD & "^" & pSampleInfo.CONTAINER & Chr(124)

            '검사항목 Order코드 추가
            sTOrder = ""
            For iCnt = 1 To pSampleInfo.ORDCNT
                'Request Information Code에 따라 검사항목을 추가하거나 취소한다.
                If Trim(sReqStatusCd) = "O" Then
                    '정상 오더
                    Select Case Trim$(pSampleInfo.IFCD(iCnt))
                        Case "989", "990", "991"
                            'ISE 항목이 있는 경우 전체 검사(ISE 검사는 2가지 조합만 오더 가능)
                            If Val(InStr(1, sTOrder, "989^")) = 0 Then
                                sTOrder = sTOrder & "^^^989^\"
                            End If
                            If Val(InStr(1, sTOrder, "990^")) = 0 Then
                                sTOrder = sTOrder & "^^^990^\"
                            End If
                            If Val(InStr(1, sTOrder, "991^")) = 0 Then
                                sTOrder = sTOrder & "^^^991^\"
                            End If
                            
                        Case Else
                            If InStr(pSampleInfo.IFCD(iCnt), "^") = 0 Then
                                '일반항목
                                sTOrder = sTOrder & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "^\"
                            Else
                                '<S--- Dilution 정보 추가
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
                                
                                sTOrder = sTOrder & "^^^" & Trim(aDilInfo(0)) & "^" & sDilInfo & "\"
                                '>E----------------------
                            End If
                    End Select
                    
                ElseIf Trim(sReqStatusCd) = "A" Then
                    '오더 취소
                    pSampleInfo.SINDEX = False
                End If
            Next iCnt
            
            If pSampleInfo.SINDEX = True Then           'Serum Index
                sTOrder = sTOrder & "^^^992^\^^^993^\^^^994^\"
            End If
            
            If pSampleInfo.ORDCNT > 0 And Trim(sReqStatusCd) <> "A" Then
                sTOrder = Left(sTOrder, Len(sTOrder) - 1)      '"\" Cutting
            End If
            
            'ISE 항목 오더전송시 오류 수정...2011/9/29 yk
            sSendBuff = sSendBuff & sTOrder
            
            'STAT RACK에 대한 처리추가
            If Left(pSampleInfo.RACK, 1) = "4" Then
                sStat = "S"
            Else
                sStat = "R"
            End If

            'Specimen Descriptor 설정
            Select Case Trim(pSampleInfo.SPCCD)
                Case "S1": sSpcGbn = "1"
                Case "S2": sSpcGbn = "2"
                Case "S3": sSpcGbn = "3"
                Case "S4": sSpcGbn = "4"
                Case "S5": sSpcGbn = "5"
                Case Else: sSpcGbn = "1"
            End Select

            sSendBuff = sSendBuff & "|" & sStat & "||" & Format(Now, "YYYYMMDDHHNNSS") & "||||A||||" & sSpcGbn _
                        & "||||||||||O" & vbCr
                        
            'Comment Record
            If Trim(pSampleInfo.CMT1) <> "" Then
                'Comment 정보 존재할 경우 Comment1 란에 특정정보 전송
                sSendBuff = sSendBuff & "C|1|L|" & Trim(pSampleInfo.CMT1) & "^^^^|G" & vbCr
            End If
                    
            'Terminator Record
            sSendBuff = sSendBuff & "L|1|N"


            '--- Text의 내용이 240byte를 넘어갈 경우 처리 추가...
            If Len(sSendBuff) >= 241 Then
                sNextSend = Mid(sSendBuff, 241)
                sSendBuff = Left(sSendBuff, 240)
                sSendBuff = sSendBuff & Chr(23)

                m_iFrameN = m_iFrameN + 1
                m_iSendPhase = 2
            Else
                sSendBuff = sSendBuff & Chr(13) & Chr(3)
                GoTo Send_Terminate
            End If

        Case 2
            sSendBuff = m_iFrameN & sNextSend & Chr(13) & Chr(3)
            sNextSend = ""

Send_Terminate:
            m_iSendPhase = 3

        Case 3      'EOT
            MSComm.Output = Chr(4)   'EOT
            m_iFrameN = 1
            m_iPhase = 3
            m_iSendPhase = 1

            sState = "": sReqStatusCd = ""

            Exit Sub
    End Select

    ChkSum = ChkSum_ASTM(sSendBuff)
    sSendBuff = sSendBuff & ChkSum
    MSComm.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10)

    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
    End If

    '전송된 오더가 있는 경우 화면표시
    If pSampleInfo.ORDCNT > 0 And sReqStatusCd = "O" Then
        If Trim(sNextSend) = "" And m_iSendPhase <> 2 Then
            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
        End If
    Else
        '조회된 내용이 없는 경우 환자정보 구조체 초기화
        Call Init_pResultInfo

        RaiseEvent SendOrderOK("", "", "", "")
    End If

Err_Rtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
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
        .SINDEX = m_p_bSIndex
        .ORDCNT = m_p_iOrdCnt
        
        ReDim .IFCD(.ORDCNT)
        iCnt = 0
        For ii = 1 To .ORDCNT
'            If Trim(tmpData(ii - 1)) <> "" Then
            If Trim(tmpData(ii - 1)) <> "" And Trim(tmpData(ii - 1)) <> "." Then    '계산식은 '.' 로 표시...2011/2/9 yk
                iCnt = iCnt + 1
                .IFCD(iCnt) = tmpData(ii - 1)
            End If
        Next ii
        .ORDCNT = iCnt      '실제 검사 가능한 항목 갯수
        
        .CMT1 = m_p_sCmt1
    End With
    
End Sub
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=msComm,msComm,-1,RTSEnable
Public Property Get RTSEnable() As Boolean
Attribute RTSEnable.VB_Description = "전송 요청 줄이 가능한지의 여부를 결정합니다."
    RTSEnable = MSComm.RTSEnable
End Property

Public Property Let RTSEnable(ByVal New_RTSEnable As Boolean)
    MSComm.RTSEnable() = New_RTSEnable
    PropertyChanged "RTSEnable"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=msComm,msComm,-1,RThreshold
Public Property Get RThreshold() As Integer
Attribute RThreshold.VB_Description = "수신할 문자의 수를 반환하거나 설정합니다."
    RThreshold = MSComm.RThreshold
End Property

Public Property Let RThreshold(ByVal New_RThreshold As Integer)
    MSComm.RThreshold() = New_RThreshold
    PropertyChanged "RThreshold"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=msComm,msComm,-1,Settings
Public Property Get Settings() As String
Attribute Settings.VB_Description = "전송 속도, 패리티, 데이터 비트, 중단 비트 매개 변수를 반환하거나 설정합니다."
    Settings = MSComm.Settings
End Property

Public Property Let Settings(ByVal New_Settings As String)
    MSComm.Settings() = New_Settings
    PropertyChanged "Settings"
End Property

Private Sub cmdTest_Click()

    wkBuf = Text1
    Call PhaseCfg_Protocol

End Sub

Private Sub msComm_OnComm()
        
    Select Case MSComm.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                                ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
            wkBuf = MSComm.Input
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintRcvLog(wkBuf)
            End If
                                
            If iSpaceCnt = 30 Then
                iSpaceCnt = 0
            End If
            iSpaceCnt = iSpaceCnt + 2
            
            RaiseEvent DispMsgComm(Space(iSpaceCnt) & "장비와 Interface 작업 중...")
            
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

    MSComm.CommPort = PropBag.ReadProperty("CommPort", 1)
    MSComm.RTSEnable = PropBag.ReadProperty("RTSEnable", False)
    MSComm.RThreshold = PropBag.ReadProperty("RThreshold", 0)
    MSComm.Settings = PropBag.ReadProperty("Settings", "9600,n,8,1")
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
    m_p_bSIndex = PropBag.ReadProperty("p_bSIndex", m_def_p_bSIndex)
    m_p_sRerunGbn = PropBag.ReadProperty("p_sRerunGbn", m_def_p_sRerunGbn)
    m_p_sTSVol = PropBag.ReadProperty("p_sTSVol", m_def_p_sTSVol)
    m_p_sSpcCd = PropBag.ReadProperty("p_sSpcCd", m_def_p_sSpcCd)
    m_p_sCmt1 = PropBag.ReadProperty("p_sCmt1", m_def_p_sCmt1)
    m_bNewMode = PropBag.ReadProperty("bNewMode", m_def_bNewMode)
End Sub

'속성값을 저장소에 기록합니다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("CommPort", MSComm.CommPort, 1)
    Call PropBag.WriteProperty("RTSEnable", MSComm.RTSEnable, False)
    Call PropBag.WriteProperty("RThreshold", MSComm.RThreshold, 0)
    Call PropBag.WriteProperty("Settings", MSComm.Settings, "9600,n,8,1")
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
    Call PropBag.WriteProperty("p_bSIndex", m_p_bSIndex, m_def_p_bSIndex)
    Call PropBag.WriteProperty("p_sRerunGbn", m_p_sRerunGbn, m_def_p_sRerunGbn)
    Call PropBag.WriteProperty("p_sTSVol", m_p_sTSVol, m_def_p_sTSVol)
    Call PropBag.WriteProperty("p_sSpcCd", m_p_sSpcCd, m_def_p_sSpcCd)
    Call PropBag.WriteProperty("p_sCmt1", m_p_sCmt1, m_def_p_sCmt1)
    Call PropBag.WriteProperty("bNewMode", m_bNewMode, m_def_bNewMode)
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
    
    '변수 초기화(E-170/H-7600)
    RstEnd = "Y": bEndChk = True: bSTXChk = False
    
    
    On Error GoTo ErrPortOpen
    If m_PortOpen = True Then
        MSComm.PortOpen = True
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
'    m_iStartSampleNo = m_def_iStartSampleNo
    m_p_bSIndex = m_def_p_bSIndex
    m_p_sRerunGbn = m_def_p_sRerunGbn
    m_p_sTSVol = m_def_p_sTSVol
    m_p_sSpcCd = m_def_p_sSpcCd
    m_p_sCmt1 = m_def_p_sCmt1
    m_bNewMode = m_def_bNewMode
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
    MSComm.Output = Chr(iChr)
    On Error GoTo 0
ErrComm:
    If Err <> 0 Then
        RaiseEvent DispMsg("Send_Chr 에러 - " & Err.Description)
    End If
End Function
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0,0,0,0
Public Property Get p_bSIndex() As Boolean
    p_bSIndex = m_p_bSIndex
End Property

Public Property Let p_bSIndex(ByVal New_p_bSIndex As Boolean)
    m_p_bSIndex = New_p_bSIndex
    PropertyChanged "p_bSIndex"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sRerunGbn() As String
    p_sRerunGbn = m_p_sRerunGbn
End Property

Public Property Let p_sRerunGbn(ByVal New_p_sRerunGbn As String)
    m_p_sRerunGbn = New_p_sRerunGbn
    PropertyChanged "p_sRerunGbn"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sTSVol() As String
    p_sTSVol = m_p_sTSVol
End Property

Public Property Let p_sTSVol(ByVal New_p_sTSVol As String)
    m_p_sTSVol = New_p_sTSVol
    PropertyChanged "p_sTSVol"
End Property

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
'MemberInfo=0,0,0,False
Public Property Get bNewMode() As Boolean
    bNewMode = m_bNewMode
End Property

Public Property Let bNewMode(ByVal New_bNewMode As Boolean)
    m_bNewMode = New_bNewMode
    PropertyChanged "bNewMode"
End Property

