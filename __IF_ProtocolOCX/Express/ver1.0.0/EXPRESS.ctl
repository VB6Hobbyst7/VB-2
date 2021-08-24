VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl EXPRESS 
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
Attribute VB_Name = "EXPRESS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
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

'For EXPRESS
Public miSendSeqNo  As Integer
Dim mbOrderFlg   As Boolean
Dim mbRstFlg     As Boolean
Dim msOldReq As String       'Req Number
Dim msState As String

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
        Case "EXPRESS_DUPLEX"
            Call PhaseCfg_Protocol_EXPRESS_DUPLEX
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub

Private Sub PhaseCfg_Protocol_EXPRESS_DUPLEX()
    
    Dim wkdat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkdat = Mid$(wkBuf, ix1, 1)
       
        Select Case iPhase
            Case 1      'SOH
                Select Case Asc(wkdat)
                    Case 2          'STX
                        RcvBuffer = ""
                        iPhase = 2
                    End Select
                
            Case 2      'STX
                Select Case Asc(wkdat)
                    Case 2          'STX
                        RcvBuffer = ""
                                        
                    Case 4          'EOT
                        'ACK SEND
                        'msComm.Output = Chr(2) + Chr(6) + Chr(3) + "0B" + Chr(4)
                        
                        Call DataEditResponse_EXPRESS_DUPLEX
                    
                        RcvBuffer = ""
                        
                    Case 23         'ETB
                    
                    Case Else
                        RcvBuffer = RcvBuffer & wkdat
                        iPhase = 2
                End Select

        End Select
    Next ix1
    
End Sub

' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*

Private Sub DataEditResponse_EXPRESS_DUPLEX()
    On Error GoTo ErrRtn
    
    Dim tmpBarCd$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$, tmpFlag$, sTmp$
    Dim iPos%, iRstCnt%, i%
    
    Dim tmpData()   As String
    Dim sBC         As String
    Dim sSend       As String
    Dim iBCpos      As Integer
    Dim iLCpos      As Integer
    Dim sReqNo As String
    Dim iRange As Integer
        
    If Asc(Left(RcvBuffer, 1)) = 6 Then
        If msState = "RSLTS_RDY" Then
            sSend = "ALL_DATA_REQ" & Chr(28) & Chr(30) & "S" & Chr(29) & miSendSeqNo & Chr(29) & Chr(29) & Chr(29) & Chr(28) & Chr(30)
                    
            Call Send_Response(sSend)
            
            msState = ""
            
        ElseIf msState = "BATCH_WKLST_REQ" Then
            
            RaiseEvent RequestNextOrder
            
            Call Get_OrderString
            
            Call SendOrder_EXPRESS
            
            msState = "RECV_WKLST"
        Else
            'ACK SEND
            msComm.Output = Chr(2) + Chr(6) + Chr(3) + "0B" + Chr(4)
            
        End If
        
        Exit Sub
    End If
    
    tmpData() = Split(RcvBuffer, Chr(28))
    
    sBC = Trim(tmpData(0))
    
    Select Case sBC
    
        Case "RESET"                'RESET CHANNEL
            sSend = "SUCCESS" & Chr(28) & Chr(30) & "H/R" & Chr(29) & "HOST" & Chr(29) & Chr(29) & Chr(29) & Chr(28) _
                    & "S" & Chr(29) & miSendSeqNo & Chr(29) & Chr(29) & Chr(29) & Chr(28) & Chr(30)
                    
            Call Send_Response(sSend)
            
            Call Init_pResultInfo
            
        Case "ID_REQ"               'INQUIRY
            sSend = "SUCCESS" & Chr(28) & Chr(30) & "aMOD" & Chr(29) & "HOST" & Chr(29) & Chr(29) & Chr(29) & Chr(28) _
                    & "DEVID" & Chr(29) & "90" & Chr(29) & Chr(29) & Chr(29) & Chr(28) _
                    & "KEY" & Chr(29) & "CUP" & Chr(29) & Chr(29) & Chr(29) & Chr(28) _
                    & "TERM" & Chr(29) & "Y" & Chr(29) & Chr(29) & Chr(29) & Chr(28) _
                    & "S" & Chr(29) & miSendSeqNo & Chr(29) & Chr(29) & Chr(29) & Chr(28) & Chr(30)
            Call Send_Response(sSend)
            
            msState = ""
                            
        Case "BATCH_WKLST_REQ"      'WorkList 요구
            iPos = InStr(RcvBuffer, "TERM")
            
            RcvBuffer = Mid(RcvBuffer, iPos + 4)
            Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr(29))
            sTmp = Trim(GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr(29)))
            
            If sTmp = "Y" Then
                msState = "BATCH_WKLST_REQ"
            End If
                        
            sSend = "SUCCESS" & Chr(28) & Chr(30) & "S" & Chr(29) & miSendSeqNo & Chr(29) & Chr(29) & Chr(29) & Chr(28) & Chr(30)
            Call Send_Response(sSend)
               
        Case "SUCCESS"
            If msState = "RECV_WKLST" Then
            
                RaiseEvent RequestNextOrder
            
                Call Get_OrderString
                
                Call SendOrder_EXPRESS
                
                msState = "RECV_WKLST"
                
            Else
                sSend = "ALL_DATA_REQ" & Chr(28) & Chr(30) & "S" & Chr(29) & miSendSeqNo & Chr(29) & Chr(29) & Chr(29) & Chr(28) & Chr(30)
                Call Send_Response(sSend)
            End If
        
        Case "RSLTS_RDY"        'RESULT READY
            mbRstFlg = True
            sSend = "SUCCESS" & Chr(28) & Chr(30) & "S" & Chr(29) & miSendSeqNo & Chr(29) & Chr(29) & Chr(29) & Chr(28) & Chr(30)
            Call Send_Response(sSend)
            
            msState = "RSLTS_RDY"
        
        Case "RECV_RSLTS"       '결과 DATA 편집
            Call Init_pResultInfo
                        
            iBCpos = InStr(RcvBuffer, "REQ")
            If iBCpos = 0 Then
                Exit Sub
            End If
            
            RcvBuffer = Mid(RcvBuffer, iBCpos + 3)
            
            'REQ NUM 편집
            Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr(29))
            sTmp = GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr(29))
            If Trim(sTmp) <> "" Then
                sReqNo = Trim(sTmp)
                pResultInfo.SEQNO = sReqNo
            End If
            
            '검사코드/결과값 편집
            Do
                iLCpos = InStr(RcvBuffer, "TS=")
                If iLCpos = 0 Then
                    Exit Do
                End If
            
                RcvBuffer = Mid(RcvBuffer, iLCpos + 3)
                sTmp = GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr(29))
                pResultInfo.IFCD = pResultInfo.IFCD & Trim(sTmp) & "|"
                
                sTmp = GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr(29))
                pResultInfo.RST1 = pResultInfo.RST1 & Trim(sTmp) & "|"
                pResultInfo.RST2 = pResultInfo.RST2 & "" & "|"
                pResultInfo.UNIT = pResultInfo.UNIT & "" & "|"
                pResultInfo.FLAG = pResultInfo.FLAG & "" & "|"
                
                pResultInfo.RSTCNT = pResultInfo.RSTCNT + 1
            Loop ''
            
             With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, .ALARMCD, .KIND, .RSTDT, .OTHER)
                End If
            End With
            
            Call Init_pResultInfo
            
            sSend = "SUCCESS" & Chr(28) & Chr(30) & "S" & Chr(29) & miSendSeqNo & Chr(29) & Chr(29) & Chr(29) & Chr(28) & Chr(30)
            Call Send_Response(sSend)
    
        Case "RSLTS_DONE"       '결과전송완료
            sSend = "SUCCESS" & Chr(28) & Chr(30) & "S" & Chr(29) & miSendSeqNo & Chr(29) & Chr(29) & Chr(29) & Chr(28) & Chr(30)
            Call Send_Response(sSend)
            
        Case "SEQUENCE"
            sSend = "WKLST_DONE" & Chr(28) & Chr(30) & "S" & Chr(29) & miSendSeqNo & Chr(29) & Chr(29) & Chr(29) & Chr(28) & Chr(30)
            Call Send_Response(sSend)
            
    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub

Private Sub SendOrder_EXPRESS()
    Dim sSend As String
    Dim sTest As String
    Dim i     As Integer
    
    If pSampleInfo.ORDCNT > 0 Then
        'Send Order
        'RECV_WKLSTrCUP1REQ1rTYPEPATTS=GLU1TS=CHO1TS=GOT1TS=GPT1TS=GGT1TS=TP1TS=ALB1TS=TBI1TS=DBI1TS=BUN1TS=CREA1TS=ALP1TS=LDH1TS=TG1TS=UA1TS=HDL1TS=LDL1S5AF
        
        sSend = "RECV_WKLST" & Chr(28) & Chr(30) & "rCUP" & Chr(29) & pSampleInfo.SEQNO & Chr(29) & Chr(29) & Chr(29) & Chr(28) & _
                "REQ" & Chr(29) & pSampleInfo.SEQNO & Chr(29) & Chr(29) & Chr(29) & Chr(28) & _
                "rTYPE" & Chr(29) & "PAT" & Chr(29) & Chr(29) & Chr(29) & Chr(28)
                
        For i = 1 To pSampleInfo.ORDCNT
            sTest = sTest & "TS=" & pSampleInfo.IFCD(i) & Chr(29) & "1" & Chr(29) & Chr(29) & Chr(29) & Chr(28)
        Next
        
        sSend = sSend & sTest & "S" & Chr(29) & miSendSeqNo & Chr(29) & Chr(29) & Chr(29) & Chr(28) & Chr(30)
        
        RaiseEvent SendOrderOK(pResultInfo.ID, pResultInfo.SEQNO, pResultInfo.RACK, pResultInfo.POS)
    Else
        
        sSend = "WKLST_DONE" & Chr(28) & Chr(30) & "S" & Chr(29) & miSendSeqNo & Chr(29) & Chr(29) & Chr(29) & Chr(28) & Chr(30)
        
    End If
        
    Call Send_Response(sSend)

End Sub

Public Sub Send_Response(ByVal sSend As String)
            
    Dim sCS     As String
        
    sSend = Chr(2) & sSend & Chr(3)
    sCS = Chk_Sum(sSend)
        
    Call Sleep(500)
    
    msComm.Output = sSend & sCS & Chr(4)

    miSendSeqNo = miSendSeqNo + 1
    If miSendSeqNo > 9 Then
        miSendSeqNo = 0
    End If
    
    If m_sTestMode = 77 Then
        RaiseEvent PrintSendLog(sSend & sCS & Chr(4))
    End If
    
End Sub

Public Sub Send_Worklist()
    Dim sCS     As String
    Dim sSend   As String
    
    'WORKLIST READY
    sSend = Chr(2) & "WKLST_RDY" & Chr(28) & Chr(30) & "S" & Chr(29) & miSendSeqNo & Chr(29) & Chr(29) & Chr(29) & Chr(28) & Chr(30) & Chr(3)
    
    sCS = Chk_Sum(sSend)
        
    Call Sleep(500)
    
    msComm.Output = sSend & sCS & Chr(4)

    miSendSeqNo = miSendSeqNo + 1
    If miSendSeqNo > 9 Then
        miSendSeqNo = 0
    End If
    
    If m_sTestMode = 77 Then
        RaiseEvent PrintSendLog(sSend & sCS & Chr(4))
    End If
    
End Sub

'
'   CheckSum 계산
'
Private Function Chk_Sum(ByVal Para As String)

    Dim i  As Integer
    Dim Tmp   As Integer
    Dim ChkS1  As Integer
    Dim ChkS2  As String
    
    For i = 1 To Len(Para)
        Tmp = Asc(Mid$(Para, i, 1))
        ChkS1 = ChkS1 + Tmp
    Next i
    
    If ChkS1 > 16 Then
        ChkS1 = ChkS1 - 16
    End If
    
    ChkS1 = ChkS1 Mod 256
    ChkS2 = Right$("0" & Hex$(ChkS1), 2)
    
    Chk_Sum = ChkS2
    
End Function

Public Function GetByOneUserSymbol(ByVal tStr As String, sOriginal As String, ByVal sUserSymbol As String) As String
    Dim POS%

    POS = InStr(tStr, sUserSymbol)

    If POS = 0 Then
    Else
        GetByOneUserSymbol = Trim$(Mid$(tStr, 1, POS - 1))
        sOriginal = Trim$(Mid$(sOriginal, POS + 1, Len(sOriginal) - POS))
    End If
End Function


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
            If Trim(tmpData(ii - 1)) <> "" Then
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
            
            If m_sTestMode = "77" Then
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
    m_p_bSIndex = PropBag.ReadProperty("p_bSIndex", m_def_p_bSIndex)
    m_p_sRerunGbn = PropBag.ReadProperty("p_sRerunGbn", m_def_p_sRerunGbn)
    m_p_sTSVol = PropBag.ReadProperty("p_sTSVol", m_def_p_sTSVol)
    m_p_sSpcCd = PropBag.ReadProperty("p_sSpcCd", m_def_p_sSpcCd)
    m_p_sCmt1 = PropBag.ReadProperty("p_sCmt1", m_def_p_sCmt1)
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
    Call PropBag.WriteProperty("p_bSIndex", m_p_bSIndex, m_def_p_bSIndex)
    Call PropBag.WriteProperty("p_sRerunGbn", m_p_sRerunGbn, m_def_p_sRerunGbn)
    Call PropBag.WriteProperty("p_sTSVol", m_p_sTSVol, m_def_p_sTSVol)
    Call PropBag.WriteProperty("p_sSpcCd", m_p_sSpcCd, m_def_p_sSpcCd)
    Call PropBag.WriteProperty("p_sCmt1", m_p_sCmt1, m_def_p_sCmt1)
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
'    m_iStartSampleNo = m_def_iStartSampleNo
    m_p_bSIndex = m_def_p_bSIndex
    m_p_sRerunGbn = m_def_p_sRerunGbn
    m_p_sTSVol = m_def_p_sTSVol
    m_p_sSpcCd = m_def_p_sSpcCd
    m_p_sCmt1 = m_def_p_sCmt1
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

